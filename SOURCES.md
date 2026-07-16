# 同梱ライブラリの来歴（Supply Chain Record）

このツールは外部 CDN を実行時に読み込まず、以下のライブラリを `lib/` に**ローカル同梱**しています。
これによりサプライチェーン汚染（CDN 改ざん）と外部送信のリスクを避けています。

- 取得日: 2026-06-08（PDF/Word 系）／2026-06-13（PowerPoint 系：PPTXjs ほか）
- 検証コマンド: `shasum -a 256 lib/*.js`（下表の SHA-256 と一致することを確認）

### PDF / Word 変換（既存）

| ファイル | バージョン | ライセンス | SHA-256 | 入手元 |
|---|---|---|---|---|
| pdf-lib.min.js | 1.17.1 | MIT | `0f9a5cad07941f0826586c94e089d89b918c46e5c17cf2d5a3c6f666e3bc694f` | https://unpkg.com/pdf-lib@1.17.1/dist/pdf-lib.min.js |
| pdf.min.js | 3.11.174 | Apache-2.0 | `5b5799e6f8c680663207ac5b42ee14eed2a406fa7af48f50c154f0c0b1566946` | https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js |
| pdf.worker.min.js | 3.11.174 | Apache-2.0 | `feabdf309770ed24bba31a5467836cdc8cf639c705af27d52b585b041bb8527b` | https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js |
| mammoth.browser.min.js | 1.6.0 | BSD-2-Clause | `596ef52239e52d8ee3cee10b2ee4a72596abf900d0e4f468593f956e9f1809b0` | https://unpkg.com/mammoth@1.6.0/mammoth.browser.min.js |
| html2canvas.min.js | 1.4.1 | MIT | `e87e550794322e574a1fda0c1549a3c70dae5a93d9113417a429016838eab8cb` | https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js |
| Sortable.min.js | 1.15.2 | MIT | `ca68430703c4f5960e90735867c6e94d29b5a3de37107d8100e5a301007e9e6e` | https://cdnjs.cloudflare.com/ajax/libs/Sortable/1.15.2/Sortable.min.js |

### PowerPoint(.pptx) 変換（2026-06-13 追加）

PPTXjs（github.com/meshesha/PPTXjs）はタグ `v1.21.1` を固定で取得。jQuery プラグインとして
スライドを HTML/SVG に描画し、その DOM を html2canvas で画像化して PDF に焼く。チャート用 d3 / nv.d3 は
**不採用**（文字・図形・画像の描画には不要。チャートは崩れる可能性ありとして許容）。いずれも実行時 CDN は
参照せず `lib/` から読み込む。読み込み順は jquery → filereader → jszip → pptxjs → divs2slides（index.html 参照）。

| ファイル | バージョン | ライセンス | SHA-256 | 入手元（取得日 2026-06-13） |
|---|---|---|---|---|
| jquery-1.11.3.min.js | 1.11.3 | MIT | `ecb916133a9376911f10bc5c659952eb0031e457f5df367cde560edbfba38fb8` | https://cdn.jsdelivr.net/gh/meshesha/PPTXjs@v1.21.1/js/jquery-1.11.3.min.js |
| filereader.js | 0.99 | MIT | `96df9b7a2c5801e64fbf4917ea8d10167e22b90de17767352e286e8202737079` | https://cdn.jsdelivr.net/gh/meshesha/PPTXjs@v1.21.1/js/filereader.js |
| jszip.min.js | 2.x（JSZip 2 系・PPTXjs 同梱） | MIT or GPLv3 | `215fb2537b13d82daabd46e1ee59ffe4dce90abd0acb0ac5432e77071f422e9c` | https://cdn.jsdelivr.net/gh/meshesha/PPTXjs@v1.21.1/js/jszip.min.js |
| pptxjs.min.js | 1.21.1（**1 箇所パッチ済・下記参照**） | MIT | `d3fc5f85f629e99736e98b03dece236942dac082390fa01ba6fcb4458ebeeb7b` | https://cdn.jsdelivr.net/gh/meshesha/PPTXjs@v1.21.1/js/pptxjs.min.js |
| divs2slides.min.js | 1.3（PPTXjs 同梱） | MIT | `aebd8d03ecd91d13ca35ada9376dac62a61c56d4e8737d18779c5237ac0ec5d0` | https://cdn.jsdelivr.net/gh/meshesha/PPTXjs@v1.21.1/js/divs2slides.min.js |
| pptxjs.css | 1.21.1（PPTXjs 同梱） | MIT | `c488fb0ec604387bb2a6b8cc8f1e2427dc6e24b2e38cf239b38b10a480195ed3` | https://cdn.jsdelivr.net/gh/meshesha/PPTXjs@v1.21.1/css/pptxjs.css |

> **CSP 適合**: 採用した全ファイルに `eval(` / `new Function` は存在しないことを確認済み。
> 現行 CSP（`script-src 'self'`、`unsafe-eval` なし）のままで動作する。

#### pptxjs.min.js への適用パッチ（1 箇所）

PPTXjs 1.21.1 には、効果（影など）の付かない単純な図形を SVG 描画する際に属性区切りの
シングルクォートが 1 つ余分に出力されるバグがある（`_name='…''  style=` の `''`）。これにより
不正な `data:image/svg+xml` 画像のロードが失敗し、ブラウザのコンソールに
「Error loading svg …」エラーが図形ごとに大量に出る。出力 PDF への影響は無い（文字は別レイヤーで描画される）が、
コンソールエラー 0 件を満たすため、配布元の該当 1 文字を修正した：

- 変更内容: `_name='" + o + "'' style='` → `_name='" + o + "' style='`（余分な `'` を 1 個削除）
- 影響: 図形が壊れた `<img>` ではなく正しいインライン `<svg>` として描画されるようになり、画像ロードエラーが消える
- 確認: `node --check lib/pptxjs.min.js` が通ること、上記 SHA-256（パッチ後）と一致すること

オリジナル（未パッチ）の SHA-256 は `845555ec4179f557f0b78822baeffbaa6aa14c303eaaba1def7f608367eaca46`。
PPTXjs を更新する場合は、新バージョンに同種の typo が残っていないか確認し、残っていれば同じ修正を再適用する。

## 更新する場合

1. 上記 URL（またはバージョンを上げた URL）から新しいファイルを取得し `lib/` に置く。
2. PPTXjs を更新したときは、上記「適用パッチ」の typo が残っていないか確認し、必要なら再適用する。
3. `shasum -a 256 lib/*.js` を実行し、この表のハッシュを更新する。
4. ブラウザで一通り（PDF 読み込み・回転・並べ替え・書き出し・Word 変換・PowerPoint 変換）動作確認する。
   PowerPoint 系ライブラリを差し替えたら、`index.html` 内のキャッシュ無効化用クエリ（`app.js?v=`, `pptxjs.min.js?v=`）の番号も上げる。
