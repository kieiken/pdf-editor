"use strict";

/* ====== ライブラリ参照 ====== */
const { PDFDocument, degrees } = PDFLib;
pdfjsLib.GlobalWorkerOptions.workerSrc = "lib/pdf.worker.min.js";

/* ====== 状態 ====== */
const sources = {};      // docId -> { name, color, bytes(Uint8Array), pdfjsDoc, numPages }
let pages = [];          // { id, docId, pageIndex(0-based), rotation, selected, _thumb }
const PALETTE = ["#3b6ef5","#1f9d55","#e5484d","#d97706","#8b5cf6","#0891b2","#db2777","#65a30d","#475569","#ea580c"];
let colorIdx = 0;
let sortable = null;
const WARN_MB = 100;     // この合計サイズを超えたら確認

/* ====== ユーティリティ ====== */
let _uid = 0; const uid = () => "u" + (++_uid) + "_" + (performance.now()|0);
const $ = (s) => document.querySelector(s);
const grid = $("#grid");

function toast(msg, isErr){
  const t = $("#toast"); t.textContent = msg; t.className = "on" + (isErr ? " err" : "");
  // 長い通知（音声除外・隠しスライド等）は読める時間まで表示を延長
  clearTimeout(toast._t); toast._t = setTimeout(()=> t.className = "", Math.max(3200, String(msg).length * 80));
}
function busy(on, msg, sub){
  $("#busy").classList.toggle("on", !!on);
  if(msg) $("#busyMsg").textContent = msg;
  $("#busySub").textContent = sub || "";
}

/* ====== ライブラリの遅延読み込み（Word/PowerPoint 用は必要になった時だけ読む） ====== */
const _libLoaded = {};
function loadScript(src){
  return _libLoaded[src] || (_libLoaded[src] = new Promise((res, rej)=>{
    const s = document.createElement("script");
    s.src = src;
    s.onload = ()=>res();
    s.onerror = ()=>{ delete _libLoaded[src]; s.remove(); rej(new Error(src + " の読み込みに失敗しました")); };
    document.head.appendChild(s);
  }));
}
async function ensureWordLibs(){
  await loadScript("lib/mammoth.browser.min.js");
  await loadScript("lib/html2canvas.min.js");
}
async function ensurePptxLibs(){
  // 依存順を維持するため直列に読み込む
  await loadScript("lib/jquery-1.11.3.min.js");
  await loadScript("lib/filereader.js");
  await loadScript("lib/jszip.min.js");
  await loadScript("lib/pptxjs.min.js");
  await loadScript("lib/divs2slides.min.js");
  await loadScript("lib/html2canvas.min.js");
}

/* ====== PDFを開く（パスワード保護に対応） ====== */
async function openPdfJs(bytes, name){
  let password;
  for(;;){
    try{
      return await pdfjsLib.getDocument({ data: bytes.slice(0), password }).promise;
    }catch(err){
      if(err && err.name === "PasswordException"){
        password = prompt(`「${name}」はパスワードで保護されています。\nパスワードを入力してください：`);
        if(password === null) throw new Error("パスワードが入力されなかったため読み込みを中止しました");
        continue;
      }
      throw err;
    }
  }
}

/* ====== ファイル読み込み ====== */
// 同時に複数回呼ばれても順番に処理する（ビジー表示・トースト・状態の競合を防ぐ）
let _addChain = Promise.resolve();
function addFiles(fileList){
  const files = [...fileList];
  return (_addChain = _addChain.then(()=> _addFiles(files)).catch(()=>{}));
}
async function _addFiles(fileList){
  const files = [...fileList].filter(f => /\.(pdf|docx|pptx|png|jpe?g|webp|gif|bmp|svg)$/i.test(f.name));
  const rejected = [...fileList].length - files.length;
  if(!files.length){ toast("PDF・Word（.docx）・PowerPoint（.pptx）・画像（PNG/JPEG など）を選んでください", true); return; }
  const totalMB = files.reduce((s,f)=> s + (f.size||0), 0) / 1048576;
  if(totalMB > WARN_MB && !confirm(`読み込むファイルの合計が約 ${Math.round(totalMB)}MB あります。\n大きいファイルはブラウザの動作が重くなることがあります。続けますか？`)) return;
  let added = 0, audioTotal = 0, videoTotal = 0, hiddenTotal = 0, hiddenIncludedAny = false;
  try{
    for(let n=0;n<files.length;n++){
      const file = files[n];
      busy(true, "読み込み中…", `${file.name}（${n+1}/${files.length}）`);
      let bytes;
      if(/\.docx$/i.test(file.name)){
        busy(true, "Word を PDF に変換中…", file.name);
        await ensureWordLibs();
        bytes = await wordToPdfBytes(await file.arrayBuffer(), file.name);
      }else if(/\.(png|jpe?g|webp|gif|bmp|svg)$/i.test(file.name)){
        busy(true, "画像を PDF に変換中…", file.name);
        bytes = await imageToPdfBytes(file);
      }else if(/\.pptx$/i.test(file.name)){
        busy(true, "PowerPoint を PDF に変換中…", file.name);
        await ensurePptxLibs();
        const r = await pptxToPdfBytes(await file.arrayBuffer(), file.name);
        bytes = r.bytes; audioTotal += r.audioCount; videoTotal += r.videoCount;
        hiddenTotal += r.hiddenCount; if(r.hiddenIncluded) hiddenIncludedAny = true;
      }else{
        bytes = new Uint8Array(await file.arrayBuffer());
      }
      // pdf.js には複製を渡す（元バッファの detach を防ぐ）
      const pdfjsDoc = await openPdfJs(bytes, file.name);
      const docId = uid();
      sources[docId] = { name:file.name, color:PALETTE[colorIdx++ % PALETTE.length], bytes, pdfjsDoc, numPages:pdfjsDoc.numPages };
      for(let i=0;i<pdfjsDoc.numPages;i++)
        pages.push({ id:uid(), docId, pageIndex:i, rotation:0, selected:false });
      added += pdfjsDoc.numPages;
    }
    render();
    let msg = `${added} ページを追加しました`;
    const ex = [];
    if(audioTotal) ex.push(`音声 ${audioTotal} 件`);
    if(videoTotal) ex.push(`動画 ${videoTotal} 件`);
    if(ex.length) msg += `（${ex.join("・")}は含めず除外しました）`;
    if(hiddenTotal) msg += hiddenIncludedAny
      ? `（隠しスライド ${hiddenTotal} 件もページに含まれています。不要なら削除してください）`
      : `（隠しスライド ${hiddenTotal} 件は含まれていません）`;
    if(rejected) msg += `（非対応 ${rejected} 件は除外）`;
    toast(msg);
  }catch(err){
    console.error(err);
    render();   // 失敗前に追加できたページはUIに反映する
    toast("読み込みに失敗：" + (err.message || err), true);
  }finally{
    busy(false);
  }
}

/* ====== Word(.docx) → PDF バイト列（mammoth → html2canvas → pdf-lib で確実に） ====== */
async function wordToPdfBytes(arrayBuffer, name){
  const { value:html, messages } = await mammoth.convertToHtml({ arrayBuffer });
  if(messages && messages.length) console.warn("mammoth:", name, messages);
  const CW = 760, PAD = 48;
  const wrap = document.createElement("div");
  wrap.style.cssText =
    "position:fixed;left:0;top:0;z-index:1;width:"+CW+"px;padding:"+PAD+"px;"+
    "background:#fff;color:#000;font-family:'Hiragino Sans','Yu Gothic',sans-serif;font-size:15px;line-height:1.85";
  wrap.innerHTML =
    "<style>"+
    "*{box-sizing:border-box}img{max-width:100%}"+
    "h1{font-size:25px;margin:.2em 0 .6em}h2{font-size:19px;margin:1em 0 .4em;border-bottom:1px solid #ccc;padding-bottom:3px}"+
    "h3{font-size:16px;margin:.8em 0 .3em}p{margin:.55em 0}"+
    "ul,ol{margin:.5em 0;padding-left:1.6em}"+
    "table{border-collapse:collapse;width:100%;margin:.6em 0}td,th{border:1px solid #888;padding:5px 8px}"+
    "</style>" + (html || "<p>（空の文書）</p>");
  document.body.appendChild(wrap);
  try{
    const canvas = await html2canvas(wrap, { scale:2, backgroundColor:"#ffffff", useCORS:true, windowWidth:CW });
    const pdf = await PDFDocument.create();
    const A4W = 595.28, A4H = 841.89;
    const fullW = canvas.width, fullH = canvas.height;
    const pxPerPage = Math.floor(fullW * (A4H / A4W));   // A4比での1ページ分の高さ(px)
    let y = 0;
    while(y < fullH){
      const sliceH = Math.min(pxPerPage, fullH - y);
      const slice = document.createElement("canvas");
      slice.width = fullW; slice.height = sliceH;
      const sctx = slice.getContext("2d");
      sctx.fillStyle = "#fff"; sctx.fillRect(0, 0, fullW, sliceH);
      sctx.drawImage(canvas, 0, y, fullW, sliceH, 0, 0, fullW, sliceH);
      const png = await pdf.embedPng(dataUrlToBytes(slice.toDataURL("image/png")));
      const page = pdf.addPage([A4W, A4H]);
      const drawH = sliceH * (A4W / fullW);
      page.drawImage(png, { x:0, y:A4H - drawH, width:A4W, height:drawH });
      y += sliceH;
    }
    return await pdf.save();
  } finally {
    wrap.remove();
  }
}
function dataUrlToBytes(dataUrl){
  const bin = atob(dataUrl.split(",")[1]);
  const arr = new Uint8Array(bin.length);
  for(let i=0;i<bin.length;i++) arr[i] = bin.charCodeAt(i);
  return arr;
}

/* ====== 画像 → PDF バイト列（1画像=1ページ。EXIFの回転情報も反映） ====== */
async function imageToPdfBytes(file){
  // デコード：createImageBitmap は EXIF 回転を反映できる。SVG 等で失敗したら <img> にフォールバック
  let bmp = null, imgEl = null, w, h;
  try{
    bmp = await createImageBitmap(file, { imageOrientation: "from-image" });
    w = bmp.width; h = bmp.height;
  }catch(e){
    imgEl = await new Promise((res, rej)=>{
      const url = URL.createObjectURL(file);
      const im = new Image();
      im.onload = ()=>{ URL.revokeObjectURL(url); res(im); };
      im.onerror = ()=>{ URL.revokeObjectURL(url); rej(new Error(`「${file.name}」を画像として読み込めませんでした`)); };
      im.src = url;
    });
    w = imgEl.naturalWidth; h = imgEl.naturalHeight;
  }
  if(!(w > 0) || !(h > 0)) throw new Error(`「${file.name}」の画像サイズを取得できませんでした（SVG は width/height 指定が必要な場合があります）`);
  try{
    // 巨大画像は長辺 4000px まで縮小（PDFの肥大とメモリを抑える）
    const MAX = 4000;
    const scale = Math.min(1, MAX / Math.max(w, h));
    const cw = Math.max(1, Math.round(w * scale)), ch = Math.max(1, Math.round(h * scale));
    const canvas = document.createElement("canvas");
    canvas.width = cw; canvas.height = ch;
    const ctx = canvas.getContext("2d");
    // 透明画素が実際にある場合のみ PNG、それ以外は白地の JPEG（軽量）
    const src = bmp || imgEl;
    let usePng = false;
    if(/\.(png|gif|svg|webp)$/i.test(file.name)){
      try{
        const probe = document.createElement("canvas");
        probe.width = probe.height = 64;
        const pctx = probe.getContext("2d");
        pctx.drawImage(src, 0, 0, 64, 64);
        const d = pctx.getImageData(0, 0, 64, 64).data;
        for(let i = 3; i < d.length; i += 4){ if(d[i] < 255){ usePng = true; break; } }
      }catch(e){ usePng = true; }   // 判定できなければ安全側（PNG）に倒す
    }
    if(!usePng){ ctx.fillStyle = "#fff"; ctx.fillRect(0, 0, cw, ch); }
    ctx.drawImage(src, 0, 0, cw, ch);

    const pdf = await PDFDocument.create();
    const dataUrl = canvas.toDataURL(usePng ? "image/png" : "image/jpeg", 0.92);
    const embedded = usePng
      ? await pdf.embedPng(dataUrlToBytes(dataUrl))
      : await pdf.embedJpg(dataUrlToBytes(dataUrl));
    // ページサイズは画像の縦横比のまま（px → pt、長辺を A4 長辺程度に収める）
    const LONG = 841.89;
    const k = LONG / Math.max(cw, ch);
    const pw = cw * k, ph = ch * k;
    const page = pdf.addPage([pw, ph]);
    page.drawImage(embedded, { x: 0, y: 0, width: pw, height: ph });
    return await pdf.save();
  }finally{
    if(bmp && bmp.close) bmp.close();   // 例外時も ImageBitmap を確実に解放
  }
}

/* ====== PowerPoint(.pptx) → PDF バイト列（PPTXjs でスライドを描画 → html2canvas → pdf-lib） ======
   音声・動画は画像化パイプライン上 PDF に持ち込めない。件数を数えてトーストで報告し、
   レンダラが生成した <audio>/<video> は画像化の前に DOM から除去する。元の pptx は変更しない。 */
const PPTX_AUDIO_EXT = /\.(mp3|wav|m4a|wma|aac|ogg|mid|midi)$/i;
const PPTX_VIDEO_EXT = /\.(mp4|mov|wmv|avi|m4v|mkv|webm|mpg|mpeg|3gp)$/i;

function pptxScanMedia(bytes){
  // JSZip 2.x（PPTXjs 同梱）の同期 API で走査。
  // 音声/動画は ppt/media/ の拡張子に加えてスライドの .rels（audio/video relationship）からも検出し、
  // ファイル名で重複排除（拡張子が変えられた埋め込みや外部リンク音声も拾う）。
  // あわせてスライド総数と隠しスライド（<p:sld show="0">）も数える。
  const audio = new Set(), video = new Set();
  let slideCount = 0, hiddenCount = 0;
  try{
    const zip = new JSZip(bytes);
    for(const name of Object.keys(zip.files || {})){
      if(/^ppt\/media\//i.test(name)){
        const base = name.toLowerCase().replace(/^.*\//,"");
        if(PPTX_AUDIO_EXT.test(name)) audio.add(base);
        else if(PPTX_VIDEO_EXT.test(name)) video.add(base);
      }else if(/^ppt\/slides\/slide\d+\.xml$/i.test(name)){
        slideCount++;
        try{ if(/<p:sld[^>]*\bshow="0"/.test(zip.files[name].asText())) hiddenCount++; }catch(e){}
      }else if(/^ppt\/slides\/_rels\/[^/]+\.rels$/i.test(name)){
        try{
          const txt = zip.files[name].asText();
          const re = /<Relationship\b[^>]*Type="[^"]*\/(audio|video)"[^>]*>/gi;
          let m;
          while((m = re.exec(txt))){
            const tgt = (m[0].match(/Target="([^"]+)"/i) || [])[1] || "";
            const base = tgt.toLowerCase().replace(/^.*\//,"") || ("rel@" + name + "#" + re.lastIndex);
            (m[1].toLowerCase() === "audio" ? audio : video).add(base);
          }
        }catch(e){}
      }
    }
  }catch(e){ console.warn("pptx media scan:", e); }
  return { audioCount: audio.size, videoCount: video.size, slideCount, hiddenCount };
}

function pptxSlideSizePt(bytes){
  // ppt/presentation.xml の <p:sldSz cx cy>（EMU）→ pt（÷12700）。読めなければ 16:9 既定。
  try{
    const xml = new JSZip(bytes).file("ppt/presentation.xml").asText();
    const m = xml.match(/<p:sldSz[^>]*\bcx="(\d+)"[^>]*\bcy="(\d+)"/);
    if(m){
      const w = parseInt(m[1],10) / 12700, h = parseInt(m[2],10) / 12700;
      if(w > 10 && h > 10) return { w, h };
    }
  }catch(e){ console.warn("pptx sldSz:", e); }
  return { w: 960, h: 540 }; // 16:9 既定
}

async function pptxToPdfBytes(arrayBuffer, name){
  const bytes = new Uint8Array(arrayBuffer);
  const { audioCount, videoCount, slideCount, hiddenCount } = pptxScanMedia(bytes);
  const { w:slideWpt, h:slideHpt } = pptxSlideSizePt(bytes);

  // レンダリング用コンテナ（可視だが busy オーバーレイ z-index:60 の下に隠れる）
  const host = document.createElement("div");
  const hostId = "pptxHost_" + uid();   // uid はカウンタ入りで同一ミリ秒でも衝突しない
  host.id = hostId;
  host.style.cssText = "position:fixed;left:0;top:0;z-index:1;background:#fff";
  document.body.appendChild(host);

  // バイト列を blob URL にして PPTXjs に渡す（XHR 先は blob:。CSP connect-src blob: で許可済み）
  const blobUrl = URL.createObjectURL(new Blob([bytes], { type:"application/vnd.openxmlformats-officedocument.presentationml.presentation" }));

  try{
    jQuery("#" + hostId).pptxToHtml({
      pptxFileUrl: blobUrl,
      slideMode: false,        // スライドショーにしない（全スライドを縦に並べる）
      keyBoardShortCut: false,
      mediaProcess: false      // 音声/動画の埋め込み処理を行わない
    });

    // 完了イベントが無いため .slide の生成数が安定するまでポーリング（最大30秒）
    const slides = await pptxWaitSlides(host, 30000);
    if(!slides.length) throw new Error(`「${name}」のスライドを読み込めませんでした（対応していない内容の可能性があります）`);

    // 画像化の前に音声/動画要素を DOM から除去
    host.querySelectorAll("audio,video").forEach(el => el.remove());

    const pdf = await PDFDocument.create();
    for(let i=0;i<slides.length;i++){
      busy(true, "PowerPoint を PDF に変換中…", `${name}（${i+1}/${slides.length} スライド）`);
      const el = slides[i];
      const canvas = await html2canvas(el, { scale:2, backgroundColor:"#ffffff", useCORS:true });
      const png = await pdf.embedPng(dataUrlToBytes(canvas.toDataURL("image/png")));
      const page = pdf.addPage([slideWpt, slideHpt]);
      page.drawImage(png, { x:0, y:0, width:slideWpt, height:slideHpt });
    }
    // 隠しスライドがレンダリング結果に含まれたか（描画枚数がスライド総数以上なら含まれている）
    const hiddenIncluded = hiddenCount > 0 && slideCount > 0 && slides.length >= slideCount;
    return { bytes: await pdf.save(), audioCount, videoCount, hiddenCount, hiddenIncluded };
  } finally {
    host.remove();
    URL.revokeObjectURL(blobUrl);
  }
}

// .slide が現れ、件数が安定（連続して同数）するまで待つ
function pptxWaitSlides(host, timeoutMs){
  return new Promise((resolve)=>{
    const t0 = performance.now();
    let last = -1, stable = 0;
    const tick = ()=>{
      const slides = host.querySelectorAll(".slide");
      const n = slides.length;
      if(n > 0 && n === last){
        if(++stable >= 3){ resolve([...slides]); return; }
      }else{
        stable = 0; last = n;
      }
      if(performance.now() - t0 > timeoutMs){ resolve([...slides]); return; }
      setTimeout(tick, 200);
    };
    tick();
  });
}

/* ====== 描画 ====== */
const observer = new IntersectionObserver((entries)=>{
  for(const e of entries){
    if(e.isIntersecting){ drawThumb(e.target); observer.unobserve(e.target); }
  }
}, { rootMargin:"300px" });

function render(){
  $("#empty").style.display = pages.length ? "none" : "block";
  grid.style.display = pages.length ? "grid" : "none";

  grid.innerHTML = "";
  pages.forEach((p, idx)=>{
    const src = sources[p.docId];
    const card = document.createElement("div");
    card.className = "card" + (p.selected ? " selected" : "");
    card.dataset.id = p.id;
    card.innerHTML =
      '<div class="thumb"><div class="ph"><div class="spinner"></div></div></div>'+
      '<div class="topbar">'+
        '<div class="chk'+(p.selected?" on":"")+'" data-act="sel" title="選択"></div>'+
        '<div class="num">'+(idx+1)+'</div>'+
      '</div>'+
      '<div class="foot">'+
        '<div class="fname" title="'+escapeAttr(src.name)+'（元 '+(p.pageIndex+1)+' ページ目）">'+
          '<span class="fdot" style="background:'+src.color+'"></span>'+
          '<span class="nm">'+escapeHtml(src.name)+'</span>'+
          '<span class="pg">p.'+(p.pageIndex+1)+'</span>'+
        '</div>'+
        '<div class="acts">'+
          '<button class="iconbtn" data-act="rotL" title="左90°回転">↺</button>'+
          '<button class="iconbtn" data-act="rotR" title="右90°回転">↻</button>'+
          '<button class="iconbtn" data-act="dup" title="複製">⧉</button>'+
          '<button class="iconbtn del" data-act="del" title="削除">🗑</button>'+
        '</div>'+
      '</div>';
    grid.appendChild(card);
    if(p._thumb && p._thumb[p.rotation]){ showThumb(card, p._thumb[p.rotation]); }
    else observer.observe(card);
  });
  initSortable();
  updateToolbar();
}

function escapeHtml(s){ return String(s).replace(/[&<>]/g, c=>({"&":"&amp;","<":"&lt;",">":"&gt;"}[c])); }
function escapeAttr(s){ return String(s).replace(/"/g,"&quot;"); }

function showThumb(card, dataUrl){
  const thumb = card.querySelector(".thumb");
  thumb.innerHTML = '<img alt="ページ" src="'+dataUrl+'">';
}

async function drawThumb(card){
  const p = pages.find(x=>x.id === card.dataset.id);
  if(!p) return;
  if(p._thumb && p._thumb[p.rotation]){ showThumb(card, p._thumb[p.rotation]); return; }
  try{
    const src = sources[p.docId];
    const page = await src.pdfjsDoc.getPage(p.pageIndex + 1);
    const baseRot = page.rotate || 0;
    const vp = page.getViewport({ scale:1, rotation:(baseRot + p.rotation) % 360 });
    const scale = Math.min(320 / vp.width, 452 / vp.height) * (window.devicePixelRatio>1?1.4:1.1);
    const v = page.getViewport({ scale, rotation:(baseRot + p.rotation) % 360 });
    const canvas = document.createElement("canvas");
    canvas.width = Math.ceil(v.width); canvas.height = Math.ceil(v.height);
    const ctx = canvas.getContext("2d");
    // intent:"print" は rAF ではなく setTimeout 駆動になるため、タブが裏でもサムネ生成が止まらない
    await page.render({ canvasContext:ctx, viewport:v, intent:"print" }).promise;
    const url = canvas.toDataURL("image/png");
    (p._thumb || (p._thumb = {}))[p.rotation] = url;
    if(card.isConnected && card.dataset.id === p.id) showThumb(card, url);
  }catch(err){
    console.error("thumb", err);
    const thumb = card.querySelector(".thumb");
    if(thumb) thumb.innerHTML = '<div class="ph">⚠️<span>表示できません</span></div>';
  }
}

/* ====== 並べ替え ====== */
function initSortable(){
  if(sortable) sortable.destroy();
  sortable = new Sortable(grid, {
    animation:150, draggable:".card",
    filter:".iconbtn,.chk", preventOnFilter:false,
    ghostClass:"sortable-ghost", chosenClass:"sortable-chosen",
    onEnd(){
      const order = [...grid.children].map(c=>c.dataset.id);
      pages.sort((a,b)=> order.indexOf(a.id) - order.indexOf(b.id));
      [...grid.children].forEach((c,i)=>{ const n=c.querySelector(".num"); if(n) n.textContent=i+1; });
    }
  });
}

/* ====== カード操作（イベント委譲） ====== */
grid.addEventListener("click", (e)=>{
  const btn = e.target.closest("[data-act]");
  if(!btn) return;
  const card = e.target.closest(".card");
  const p = pages.find(x=>x.id === card.dataset.id);
  if(!p) return;
  const act = btn.dataset.act;
  if(act === "sel"){ p.selected = !p.selected; render(); }
  else if(act === "rotL"){ rotate(p, -90); }
  else if(act === "rotR"){ rotate(p, 90); }
  else if(act === "del"){ pages = pages.filter(x=>x !== p); render(); }
  else if(act === "dup"){
    const copy = { ...p, id:uid(), selected:false, _thumb:p._thumb };
    pages.splice(pages.indexOf(p)+1, 0, copy); render();
  }
});

function rotate(p, deg){
  p.rotation = ((p.rotation + deg) % 360 + 360) % 360;
  render();
}

/* ====== ツールバー ====== */
function selected(){ return pages.filter(p=>p.selected); }
function usedDocs(){ const u={}; pages.forEach(p=>u[p.docId]=1); return u; }
function updateToolbar(){
  const sel = selected().length, has = pages.length>0;
  $("#count").textContent = has ? `全 ${pages.length} ページ / ${Object.keys(usedDocs()).length} ファイル` : "ページなし";
  $("#selInfo").innerHTML = `選択 <b>${sel}</b> 枚`;
  $("#exportBtn").disabled = !has;
  $("#clearAll").disabled = !has;
  $("#selAll").disabled = !has;
  $("#selNone").disabled = sel===0;
  ["rotL","rotR","dupSel","delSel"].forEach(id=> $("#"+id).disabled = sel===0);
}

$("#selAll").onclick = ()=>{ const all = pages.every(p=>p.selected); pages.forEach(p=>p.selected=!all); render(); };
$("#selNone").onclick = ()=>{ pages.forEach(p=>p.selected=false); render(); };
$("#rotL").onclick = ()=>{ selected().forEach(p=>p.rotation=((p.rotation-90)%360+360)%360); render(); };
$("#rotR").onclick = ()=>{ selected().forEach(p=>p.rotation=((p.rotation+90)%360+360)%360); render(); };
$("#delSel").onclick = ()=>{ pages = pages.filter(p=>!p.selected); render(); };
$("#dupSel").onclick = ()=>{
  const out=[]; pages.forEach(p=>{ out.push(p); if(p.selected) out.push({...p,id:uid(),selected:false,_thumb:p._thumb}); });
  pages = out; render();
};
$("#clearAll").onclick = ()=>{
  if(!confirm("すべてのページを消去します。よろしいですか？")) return;
  for(const k in sources){
    try{ sources[k].pdfjsDoc && sources[k].pdfjsDoc.destroy(); }catch(e){}  // pdf.jsのリソース解放
    delete sources[k];
  }
  pages = []; colorIdx = 0; render();
};

/* ====== 書き出し ====== */
$("#exportBtn").onclick = exportPdf;
async function exportPdf(){
  if(!pages.length) return;
  try{
    busy(true, "PDF を書き出し中…", `${pages.length} ページ`);
    const out = await PDFDocument.create();
    const loaded = {};
    for(const docId of Object.keys(usedDocs())){
      loaded[docId] = await PDFDocument.load(sources[docId].bytes, { ignoreEncryption:true });
    }
    for(let i=0;i<pages.length;i++){
      const p = pages[i];
      if(i % 5 === 0) busy(true, "PDF を書き出し中…", `${i+1} / ${pages.length} ページ`);
      let copied;
      try{
        [copied] = await out.copyPages(loaded[p.docId], [p.pageIndex]);
      }catch(e){
        throw new Error(`「${sources[p.docId].name}」の ${p.pageIndex+1} ページ目を処理できませんでした（${e.message||e}）`);
      }
      const orig = copied.getRotation().angle || 0;
      copied.setRotation(degrees(((orig + p.rotation) % 360 + 360) % 360));
      out.addPage(copied);
      if(i % 8 === 0) await new Promise(r=>setTimeout(r));   // UIを固めない
    }
    const bytes = await out.save();
    const blob = new Blob([bytes], { type:"application/pdf" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "結合PDF_" + stamp() + ".pdf";
    a.click();
    setTimeout(()=>URL.revokeObjectURL(a.href), 4000);
    toast(`書き出し完了：${pages.length} ページ`);
  }catch(err){
    console.error(err);
    toast("書き出しに失敗：" + (err.message || err), true);
  }finally{
    busy(false);
  }
}
function stamp(){
  const d = new Date(), z = n=>String(n).padStart(2,"0");
  return `${d.getFullYear()}${z(d.getMonth()+1)}${z(d.getDate())}_${z(d.getHours())}${z(d.getMinutes())}`;
}

/* ====== ファイル選択 / ドラッグ&ドロップ ====== */
const fileInput = $("#fileInput");
["#addBtn","#addBtn2"].forEach(s=> $(s).onclick = ()=> fileInput.click());
fileInput.onchange = ()=>{ if(fileInput.files.length) addFiles(fileInput.files); fileInput.value=""; };

let dragDepth = 0;
const overlay = $("#dropOverlay");
window.addEventListener("dragenter", e=>{ if(hasFiles(e)){ e.preventDefault(); dragDepth++; overlay.classList.add("on"); }});
window.addEventListener("dragover", e=>{ if(hasFiles(e)) e.preventDefault(); });
window.addEventListener("dragleave", e=>{ if(hasFiles(e)){ dragDepth--; if(dragDepth<=0){dragDepth=0;overlay.classList.remove("on");} }});
window.addEventListener("drop", e=>{
  if(!hasFiles(e)) return;
  e.preventDefault(); dragDepth=0; overlay.classList.remove("on");
  if(e.dataTransfer.files.length) addFiles(e.dataTransfer.files);
});
function hasFiles(e){ return e.dataTransfer && [...e.dataTransfer.types].includes("Files"); }

/* 編集中の不意の離脱を警告 */
window.addEventListener("beforeunload", e=>{ if(pages.length){ e.preventDefault(); e.returnValue=""; } });

render();
