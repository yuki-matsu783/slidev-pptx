# repair-bisect

Windows の PowerPoint で `samples/shapes.pptx` を開くと「修復」を求められる。原因を実機で絞るための PPTX 一式です。
どのファイルも、`00-full.pptx` から要因を 1 つだけ変えています（17・18 は別のデッキの比較用）。

## 開き方

1. Windows の PowerPoint で、1 つずつ開く（エクスプローラーでダブルクリック、または「ファイル」→「開く」）
2. 「保護ビュー」の帯が出たら「編集を有効にする」を押す。修復のダイアログはその後に出ることがある
3. 「PowerPoint が … の内容に問題を検出しました。修復を試みますか？」のようなダイアログが出たかを、下の表に書く
4. 修復した場合は、スライドを見て、消えた図形・形の変わった図形があればメモする
5. 保存せずに閉じてから次のファイルを開く（修復後の保存は、下の「お願い」の分だけ）

## 各ファイル

check は `packages/slidev-addon-pptx/src/opc/check.ts` の結果です（error の件数）。大きさはおおよその値です。

| ファイル | 何を変えたか | check error | 大きさ |
|---|---|---|---|
| `00-full.pptx` | 今の `tests/fixtures/deck/shapes.md` の書き出しそのまま（再現の確認用。`samples/shapes.pptx` と日時以外は同じ） | 0 | 51 KB |
| `01-slide1.pptx` | shapes.md の 1 枚目だけ（図形 1/7: accentBorderCallout1 〜 blockArc） | 0 | 21 KB |
| `02-slide2.pptx` | 2 枚目だけ（図形 2/7: borderCallout1 〜 curvedRightArrow） | 0 | 21 KB |
| `03-slide3.pptx` | 3 枚目だけ（図形 3/7: curvedUpArrow 〜 flowChartManualOperation） | 0 | 21 KB |
| `04-slide4.pptx` | 4 枚目だけ（図形 4/7: flowChartMerge 〜 irregularSeal2） | 0 | 21 KB |
| `05-slide5.pptx` | 5 枚目だけ（図形 5/7: leftArrow 〜 pentagon） | 0 | 21 KB |
| `06-slide6.pptx` | 6 枚目だけ（図形 6/7: pie 〜 star10） | 0 | 21 KB |
| `07-slide7.pptx` | 7 枚目だけ（図形 7/7: star12 〜 wedgeRoundRectCallout） | 0 | 21 KB |
| `08-slide8.pptx` | 8 枚目だけ（見本: 調整値・反転・矢じり・文字の枠） | 0 | 21 KB |
| `09-slide9.pptx` | 9 枚目だけ（境界の入力: 調整値の丸めと範囲外・知らない矢じりと図形） | 0 | 20 KB |
| `10-no-patches.pptx` | 後処理を全部外す（PptxGenJS の生の出力。無圧縮なので大きい）。check に error が 9 件ある（C2 実在しない slideMaster2〜9 への Override が 8 件、C14 `<a:pPr>` の重複が 1 件）ので、図形と関係なく修復になりうる | 9 | 395 KB |
| `10b-min-patches.pptx` | 後処理を dedupeParagraphProps と rebuildContentTypes の 2 つだけにする（10 の check error を消す最小限。図形名の付け替え・空の placeholder の除去・縮小率・調整値・ノートの段落分け・マスターの差し替えを外す） | 0 | 52 KB |
| `11-no-adjust.pptx` | 調整値の後処理（applyShapeAdjust）だけ外す。`<a:avLst>` が全部空になり、roundRect の角の半径も既定になる | 0 | 51 KB |
| `12-no-flip.pptx` | 全要素の flipH / flipV を外す | 0 | 51 KB |
| `13-no-arrows.pptx` | 全要素の矢じり（図形の arrow、線の head / tail）を外す | 0 | 51 KB |
| `14-no-connectors.pptx` | コネクタ 9 種（bentConnector2〜5、curvedConnector2〜5、straightConnector1）と lineInv の図形、線（prst="line"）の要素を取り除く | 0 | 51 KB |
| `15-no-shape-text.pptx` | 図形の中の文字を全部外す（全図形が文字なしの図形として出る）。図形の下の名前の文字枠は残す | 0 | 50 KB |
| `16-no-new-prst.pptx` | 以前の 11 種（rect roundRect ellipse line rightArrow leftArrow upArrow downArrow diamond triangle hexagon）以外の図形を取り除く。知らない図形 "foo" は rect で出るので残る | 0 | 47 KB |
| `17-plain.pptx` | `tests/fixtures/deck/plain.md` の書き出し（図形なしの比較用） | 0 | 79 KB |
| `18-components.pptx` | `tests/fixtures/deck/components.md` の書き出し | 0 | 33 KB |
| `19-only-connectors.pptx` | コネクタ 9 種と line・lineInv だけを 1 枚に並べる（文字・調整値・矢じり・反転なし） | 0 | 19 KB |
| `20-single-foldedCorner.pptx` | foldedCorner を 1 つだけ置いた 1 枚（文字・調整値なし） | 0 | 19 KB |
| `20-single-bentConnector3.pptx` | bentConnector3 を 1 つだけ置いた 1 枚（文字・調整値・矢じりなし） | 0 | 19 KB |
| `20-single-chartPlus.pptx` | chartPlus を 1 つだけ置いた 1 枚（文字・調整値なし） | 0 | 19 KB |
| `20-single-lineInv.pptx` | lineInv を 1 つだけ置いた 1 枚（文字・調整値なし） | 0 | 19 KB |
| `20-single-roundRect-adj.pptx` | roundRect に adj 50000 を付けて 1 つだけ置いた 1 枚（文字なし） | 0 | 19 KB |

## 結果（ここに書き込んでください）

| ファイル | 修復が出たか | 修復後に消えた・変わった図形があればメモ |
|---|---|---|
| `00-full.pptx` |  |  |
| `01-slide1.pptx` |  |  |
| `02-slide2.pptx` |  |  |
| `03-slide3.pptx` |  |  |
| `04-slide4.pptx` |  |  |
| `05-slide5.pptx` |  |  |
| `06-slide6.pptx` |  |  |
| `07-slide7.pptx` |  |  |
| `08-slide8.pptx` |  |  |
| `09-slide9.pptx` |  |  |
| `10-no-patches.pptx` |  |  |
| `10b-min-patches.pptx` |  |  |
| `11-no-adjust.pptx` |  |  |
| `12-no-flip.pptx` |  |  |
| `13-no-arrows.pptx` |  |  |
| `14-no-connectors.pptx` |  |  |
| `15-no-shape-text.pptx` |  |  |
| `16-no-new-prst.pptx` |  |  |
| `17-plain.pptx` |  |  |
| `18-components.pptx` |  |  |
| `19-only-connectors.pptx` |  |  |
| `20-single-foldedCorner.pptx` |  |  |
| `20-single-bentConnector3.pptx` |  |  |
| `20-single-chartPlus.pptx` |  |  |
| `20-single-lineInv.pptx` |  |  |
| `20-single-roundRect-adj.pptx` |  |  |

## お願い: 修復後の PPTX

修復が出たファイルは、修復後に PowerPoint で「名前を付けて保存」した PPTX をこのフォルダに置いてもらえると助かります
（例: `01-slide1.repaired.pptx`）。元のファイルと XML の差分を取れば、PowerPoint が何を直したか（消した要素・書き換えた属性）を見られます。

## 作り方

製品のコード（`packages/`・`tests/`）は変えていません。書き出しと同じ部品（`collect` → `build` → `postProcess(PATCHES)` → `check`）を
使い捨てのスクリプトから直接呼び、Capture（収集結果）の段階か、`build` と `postProcess` の間で 1 つの要因だけを変えて作りました。
`01`〜`09` は Capture のスライドを絞ったもので、`--range` と同じ結果です（PPTX の中の番号は 1 から振り直し）。
