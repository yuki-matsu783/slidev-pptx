# repair-bisect

Windows の PowerPoint で `samples/shapes.pptx` を開くと「修復」を求められる。原因を実機で絞るための PPTX 一式です。
どのファイルも、`00-full.pptx` から要因を 1 つだけ変えています（17・18 は別のデッキの比較用。32 は 19 から、36 は 31 から 1 つだけ変えています）。

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
| `17-plain.pptx` | `tests/fixtures/deck/plain.md` の書き出し（図形なしの比較用）。新しい prst・調整値の後処理・図形の反転・図形の矢じりの経路を通らない。以前開けたときと同じ経路だけを使う比較用 | 0 | 79 KB |
| `18-components.pptx` | `tests/fixtures/deck/components.md` の書き出し。新しい prst・調整値の後処理・図形の反転・図形の矢じりの経路を通らない。以前開けたときと同じ経路だけを使う比較用 | 0 | 33 KB |
| `19-only-connectors.pptx` | コネクタ 9 種と line・lineInv だけを 1 枚に並べる（文字・調整値・矢じり・反転なし） | 0 | 19 KB |
| `20-single-foldedCorner.pptx` | foldedCorner を 1 つだけ置いた 1 枚（文字・調整値なし） | 0 | 19 KB |
| `20-single-bentConnector3.pptx` | bentConnector3 を 1 つだけ置いた 1 枚（文字・調整値・矢じりなし） | 0 | 19 KB |
| `20-single-chartPlus.pptx` | chartPlus を 1 つだけ置いた 1 枚（文字・調整値なし） | 0 | 19 KB |
| `20-single-lineInv.pptx` | lineInv を 1 つだけ置いた 1 枚（文字・調整値なし） | 0 | 19 KB |
| `20-single-roundRect-adj.pptx` | roundRect に adj 50000 を付けて 1 つだけ置いた 1 枚（文字なし） | 0 | 19 KB |
| `30-cxnsp.pptx` | 00 のうち、prst がコネクタ 9 種・line・lineInv の `<p:sp>` 16 個を `<p:cxnSp>` に書き換える（`p:nvSpPr` → `p:nvCxnSpPr`、`p:cNvSpPr` → `p:cNvCxnSpPr`）。id・name・xfrm（反転）・ln（矢じり）・avLst・塗りはそのまま。txBody を持つものは無かった | 0 | 51 KB |
| `31-cxnsp-no-fill.pptx` | 30 から、その 16 個の `<p:spPr>` の直下の塗り（solidFill 14 個、noFill 2 個）を外す。PowerPoint が自分で書くコネクタと同じく、塗りの要素が無い形 | 0 | 51 KB |
| `32-only-connectors-cxnsp.pptx` | 19 の 11 個を、30 と同じ規則で `<p:cxnSp>` にする（塗りはそのまま） | 0 | 19 KB |
| `33-sp-no-fill.pptx` | 00 のうち、同じ 16 個を `<p:sp>` のまま、`<p:spPr>` の直下の塗り（solidFill 14 個、noFill 2 個）だけを外す。スライドの中身は、31 の要素名を `<p:sp>` に戻したものと同じ | 0 | 51 KB |
| `34-cxnsp-connectors-only.pptx` | 00 のうち、コネクタ 9 種の 13 個だけを 30 と同じ規則で `<p:cxnSp>` にする（塗りは残す）。line 2 個・lineInv 1 個は `<p:sp>` のまま | 0 | 51 KB |
| `35-cxnsp-lines-only.pptx` | 00 のうち、line 2 個・lineInv 1 個だけを 30 と同じ規則で `<p:cxnSp>` にする（塗りは残す）。コネクタ 9 種の 13 個は `<p:sp>` のまま | 0 | 51 KB |
| `36-cxnsp-no-fill-style.pptx` | 31 の 16 個の `<p:cxnSp>` に、PowerPoint が書くコネクタと同じ形の `<p:style>` を `<p:spPr>` の直後に 1 つずつ足す（下の「36 の `<p:style>`」） | 0 | 52 KB |

### 30〜32 の読み方

仮説は「PowerPoint はコネクタと線を `<p:cxnSp>` で書くのに、こちらは `<p:sp>` で出している」です。30〜32 は、要素名（31 は塗りも）だけを変えています。
図形の数・id・名前・座標・反転・prst は 00 / 19 と同じで、ECMA-376 の XSD の検証結果も 00 と同じです（`presentation.xml` の `notesMasterIdLst` の 1 件だけ）。

- 00 で修復が出て、30 で出なければ: `<p:sp>` で出していることが原因
- 30 でも出て、31 で出なければ: `<p:cxnSp>` に塗りを持たせていることが原因（XSD では許されるが、PowerPoint は書かない形）
- 30 と 31 の両方で出れば: この仮説は外れか、別の原因も残っている。14-no-connectors の結果と合わせて見る
- 19 で出て、32 で出なければ: コネクタだけの 1 枚でも同じ結論。19 で出ないなら、32 は判断の材料にならない
- 19 と 32 の両方で出れば: 30・31 の両方で出たときと同じく、この仮説は外れか、別の原因も残っている

### 33〜36 の読み方

この節の箇条書きは、「30〜32 の読み方」の箇条書きと合わせて読んでください。1 つの結果に、両方の節の文が当てはまることがあります。
例えば 30 と 31 で出て 33 で出ないときは、「30〜32 の読み方」の「30 と 31 の両方で出れば」と、この節の「33 で出なければ」が同時に成り立ちます。

33〜36 も、00（36 は 31）から 1 つだけ変えています。図形の数・id・名前・座標・反転・prst は基準と同じで、XSD の検証結果も 00 と同じです。
どれも、00 で修復が出ることが前提です。00 で出なければ、30〜36 は判断の材料になりません。

**30・31・33 と 00 の 2×2**（対象は 30 と同じ 16 個）

| | 塗りあり | 塗りなし |
|---|---|---|
| `<p:sp>` | 00 | 33 |
| `<p:cxnSp>` | 30 | 31 |

- 33 で出なければ: `<p:sp>` のままでも、塗りを外すと出なくなる。塗りが関わっている
- 30 で出なければ: 塗りを残したままでも、要素名を変えると出なくなる。要素名が関わっている
- 30 と 33 の両方で出て、31 で出なければ: 片方だけ変えても足りず、要素名と塗りの両方を変えると出なくなる
- 4 つとも出れば: 要素名と塗りのほかに原因がある（30〜32 の読み方の「両方で出れば」と同じ）

**34・35: コネクタと線を分けたもの**（30 で出なかったときに見る。30 で出たなら判断の材料にならない）

- 34 で出ず、35 で出れば: `<p:sp>` のままのコネクタ 9 種が関わっている
- 35 で出ず、34 で出れば: `<p:sp>` のままの line・lineInv が関わっている
- 34・35 の両方で出なければ: どちらか片方を `<p:cxnSp>` にするだけで出なくなる。これだけでは、どちらが原因かは決まらない
- 34・35 の両方で出れば: コネクタと線の両方を `<p:cxnSp>` にしないと出なくなる

**36: 30 と 31 の両方で出たときに試すもの**

- 31 で出て、36 で出なければ: `<p:cxnSp>` に `<p:style>` が無いことが関わっている
- 36 でも出れば: ECMA-376 の例の値の `<p:style>` を足しても出なくならない、とまでは言える。python-pptx の解析文書にある PowerPoint の出力例の値（`lnRef idx="2"`・`effectRef idx="1"`）は試していないので、`<p:style>` の有無そのものを原因から外すことはできない

上のどれにも当てはまらない結果もありえます。例えば、30 で出ないのに 31 で出る、33 で出ないのに 31 で出る、などです。1 つずつなら出なくなる変更を、重ねたらまた出るようになった場合も含みます。
このときは、それぞれの変更の効き目を単純に足し合わせても説明できません。結果はそのまま下の表に書いて伝えてください。こちらで、ファイルの生成に意図しない変更が混ざっていないかを確かめ直します。

#### 36 の `<p:style>`

ECMA-376 の cxnSp の例（PowerPoint が書いたコネクタ）の値に合わせています。

```xml
<p:style>
  <a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef>
  <a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef>
  <a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>
  <a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef>
</p:style>
```

python-pptx の解析文書（shp-connector）に載っている PowerPoint の出力例は、`lnRef idx="2"`・`effectRef idx="1"` で、この 2 つの値が違います。
線の太さ・色・矢じりは、31 と同じく `<p:spPr>` の `<a:ln>` に書いてあります。

## 結果（ここに書き込んでください）

まず 00・01〜07・14・19・30・31・33 を試してください。残りは、その結果を見てからで足ります。

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
| `10-no-patches.pptx`（注意: PptxGenJS の生の出力で、図形と関係ない check error が 9 件ある。修復が出ても図形のせいとは限らない。10b と見比べる） |  |  |
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
| `30-cxnsp.pptx` |  |  |
| `31-cxnsp-no-fill.pptx` |  |  |
| `32-only-connectors-cxnsp.pptx` |  |  |
| `33-sp-no-fill.pptx` |  |  |
| `34-cxnsp-connectors-only.pptx` |  |  |
| `35-cxnsp-lines-only.pptx` |  |  |
| `36-cxnsp-no-fill-style.pptx` |  |  |

## お願い: 修復後の PPTX

修復が出たファイルは、修復後に PowerPoint で「名前を付けて保存」した PPTX をこのフォルダに置いてもらえると助かります
（例: `01-slide1.repaired.pptx`）。元のファイルと XML の差分を取れば、PowerPoint が何を直したか（消した要素・書き換えた属性）を見られます。

## 作り方

製品のコード（`packages/`・`tests/`）は変えていません。書き出しと同じ部品（`collect` → `build` → `postProcess(PATCHES)` → `check`）を
使い捨てのスクリプトから直接呼び、Capture（収集結果）の段階か、`build` と `postProcess` の間で 1 つの要因だけを変えて作りました。
`01`〜`09` は Capture のスライドを絞ったもので、`--range` と同じ結果です（PPTX の中の番号は 1 から振り直し）。
`30`〜`36` は、`postProcess` に渡す PATCHES の最後に、スライドの XML の書き換えを 1 つ足して作りました。
