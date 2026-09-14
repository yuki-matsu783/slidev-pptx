# PPT 部品と Markdown 変換の設計（i0001-02）

`native-export.md` の §2（Capture）に何をどう入れるかを決める。置き場は `packages/slidev-addon-pptx/components/`。
`slides.md` からは `addons: [slidev-addon-pptx]` で読み込み、`<PptText>` のように使う。

改訂（i0001-06）: 実装 i0001-04 で実物（Slidev 52.19.1 の DOM、PptxGenJS 4.0.1）と食い違った箇所と、設計に無かった判断を書き戻した。
「実測」はその実装で確かめた結果を指し、検査は `tests/README.md` にある。

---

## 1. 部品の props

### 1.1 共通の props

| prop | 型 | 既定 | 意味 |
|---|---|---|---|
| `x` `y` `w` `h` | `number`（キャンバス px） | 未指定 | 座標指定。1 つでも指定すると `position: absolute` で置き、指定した値は測らずにそのまま Capture の `box` に使う。未指定の辺は実測 |
| `export` | `'native' \| 'image'` | `'native'` | `'image'` なら部品全体を画像への置き換えにする |
| `name` | `string` | `''` | PPTX の図形名。PowerPoint の「選択ウィンドウ」に出る。省略時は `<種類> <連番>`（native-export.md §4.5） |

- 座標の原点はスライド左上。`.slidev-page` が `position: absolute; inset: 0` なので、`.slidev-layout` の padding の影響を受けない。
- `x`/`y` だけ指定して `w`/`h` を省くと、幅は内容に従い（`width: max-content`、上限はスライド右端まで）、高さも内容に従う。その実測値を使う。
- 得るもの: 「だいたい実測、要るところだけ座標」で書ける。
- 失うもの: 座標指定した部品は Slidev の流し込みから外れ、他の要素と重なりうる。Slidev の画面でも重なるので気付ける。
- PPT 部品と frontmatter `zoom` の組み合わせは**未定義**。`data-ppt-box` `padding` `radius` は props の生 px で Capture に入り、実測の値は zoom を掛けた後の px なので（§2.3）、同じスライドで混ぜると単位が揃わない。`zoom` を使うスライドには PPT 部品を置かない。

### 1.2 `PptText`

| prop | 型 | 既定 | Capture |
|---|---|---|---|
| `align` | `'left' \| 'center' \| 'right' \| 'justify'` | 実測 | 全段落の `align` |
| `valign` | `'top' \| 'middle' \| 'bottom'` | `'top'` | `valign`。**`h` を指定したときだけ効く**（高さが内容で決まる枠では CSS でも効かないので、Capture にも載せない。Slidev と PPTX の見た目を揃えるため） |
| `size` | `number`（px） | 実測 | 全 run の `size`（CSS でも同じ大きさで描く） |
| `color` | `string` | 実測 | 全 run の `color` |
| `bold` `italic` | `boolean` | 実測 | 全 run |
| `fill` | `string`（色） | なし | `frame.fill` |
| `line` | `string \| { color, width, dash }` | なし | `frame.line`。文字列は色（幅 1 px、実線） |
| `radius` | `number`（px） | 0 | `frame.radius` |
| `padding` | `number \| [t, r, b, l]`（px） | 0 | `frame.inset` |

中身は default slot。Slidev の流儀で、タグの前後に空行を置けば中の Markdown が描画される。
slot の中は §3 の流し込みブロックの規則で段落に変換する。表・画像・引用などの区切りブロック（§3.3）は `W-PPT-CONTENT` を出して無視する。コードブロックは `kind: 'code'` の段落になる。
描画は `<div data-ppt="text">` で、指定した props を CSS にも反映する（Slidev の見た目と PPTX の見た目を同じにするため）。

props を「書式の上書き」に絞り、Markdown の中身は slot に任せる。
- 得るもの: 文章の書き方が普通の Markdown と同じ。部品は「枠」を足すだけ。
- 失うもの: 段落ごと・run ごとの書式は props では付けられない（Markdown の `**` や HTML で書く）。

### 1.3 `PptShape`

| prop | 型 | 既定 | 意味 |
|---|---|---|---|
| `type` | `string`（PowerPoint の図形の名前。187 種） | `'rect'` | ECMA-376 の `prst` 名をそのまま使う（`rect` `roundRect` `ellipse` `rightArrow` `wedgeRectCallout` `flowChartDecision` `bentConnector3` など。一覧は `src/shapes/presets.ts` のキーと、全種を並べたデッキ `tests/fixtures/deck/shapes.md`）。`line` だけは対角線（下の描画）。未知の名前は `console.warn`（名前ごとに 1 回）を出して `rect` で描き、書き出しでも `rect` にして `W-SHAPE` |
| `adj` | `Record<string, number>` | なし | 調整値。キーは定義の `avLst` の名前（`adj` `adj1` …）、値は ECMA の単位（例 `50000`。大半は 100000 分率、角度は 60000 分の 1 度。何に対する比かは図形ごとに定義の数式が決める）。指定しないキーは定義の既定。定義に無いキー・数でない値・丸めた後に 32 bit 整数の範囲外の値は捨て、残りは整数に丸める。この判定は `src/shapes/adjust.ts` の共通の関数 `normalizeAdjust` 1 つにまとめ、**Slidev の描画と PPTX の変換が同じ値を使う**（同じ入力なら同じ形になる）。捨てた値は書き出しで `W-SHAPE` |
| `flipH` `flipV` | `boolean` | `false` | 左右・上下反転。図形は反転し、文字は左右反転しない。`flipV` のとき文字は 180 度回る（PowerPoint と同じ）。**`type: 'line'` では無視する**（Slidev の描画でも PPTX でも。線の向きは `x y w h` で決まる） |
| `fill` | `string \| 'none'` | `'#ffffff'` | 塗り。`'none'` で塗りなし |
| `line` | `string \| 'none' \| { color, width, dash, head, tail }` | `'#000000'`（1 px） | 線。`'none'` で線なし。`head` `tail` は `'none' \| 'arrow' \| 'stealth' \| 'triangle' \| 'oval' \| 'diamond'`。**`line` 以外の図形でも効き**、開いた path（`Z` で閉じない輪郭。コネクタ・円弧・かっこなど）の始点（`head`）と終点（`tail`）に付く |
| `radius` | `number`（px） | 8 | `roundRect` の角丸。定義の式（角の半径 = 短辺 × `adj` / 100000）で `adj.adj` に換算する（`src/shapes/adjust.ts` の `radiusToAdj`。0〜50000 に収めて整数に丸める。書き出しの変換も同じ関数）。`adj.adj` は、有限で、丸めて 32 bit 整数に収まるときだけ `radius` に勝つ（書き出しの変換と同じ判定。native-export.md §4.3） |
| `rotate` | `number`（度） | 0 | 回転 |
| `padding` | `number \| [t, r, b, l]`（px） | `[0, 8, 0, 8]` | 図形の中の文字の内側余白。CSS の padding と PPTX の inset（`frame.inset`）の両方に使う |
| `align` `valign` | `PptText` と同じ | `'center'` / `'middle'` | 図形の中の文字。既定は `PptText`（左・上）と違い PowerPoint の図形と同じ中央 |
| `size` `color` | `PptText` と同じ | 実測 | 図形の中の文字 |

中身は default slot（図形の中の文字。段落の規則は `PptText` と同じ）。`type: 'line'` の slot は描かない（線の上に文字は置けない）。

- `align` / `valign` の既定を中央にする理由: PowerPoint で図形に文字を入れると中央に置かれる。「PptText と同じ」（左・上）にすると、Slidev の画面と PowerPoint で新しく足した図形の見た目が食い違う（実装 i0001-04 で変更）。
- 得るもの: Slidev で書いた図形の文字が PowerPoint の図形の流儀と同じ位置に出る。
- 失うもの: `PptText` と `PptShape` で既定が違うので、覚えることが 1 つ増える。

**描画（Slidev の画面）**。`type: 'line'` は `x y w h` で対角線を引く（`w` か `h` が 0 なら水平・垂直）。それ以外はすべて、PowerPoint の図形の定義（ECMA-376 Part 1 の `presetShapeDefinitions.xml` を `src/shapes/presets.ts` に写したもの。docs/adr/0002）を枠の実寸 px で評価し（`src/shapes/geometry.ts` の `evalPreset`）、inline SVG を背景に敷いて文字を上に重ねる。**形は PowerPoint の定義の数式どおり**で、縦横比による変形（矢印の頭の長さ、角丸の半径など）も PowerPoint と同じになる。
- `rect` `roundRect` も SVG で描く（CSS の border は使わない）。線は path の中心に引くので、太い線の外半分は枠の外にはみ出す（PowerPoint と同じ）。
- SVG の viewBox は枠の実寸 px で、引き伸ばさない（形が縦横比で変わるため）。大きさは `w` `h` があればそれ、無い辺は ResizeObserver で測る。SVG には class `ppt-shape-svg` を付け、収集器は中身を拾わない。
- 1 つの図形は path を複数持つ。path の `fill` が `none` なら塗らない。`darken` `darkenLess` `lighten` `lightenLess` は塗りの上に黒 40% / 黒 20% / 白 40% / 白 20% を重ねる。`stroke="false"` の path は線を引かない。`fill` prop は `fill="none"` でない path すべてに効く（`fill="none"` の path を持つ図形でも、他の path は塗る）。
- **塗りを全部描いてから線を描く**。定義の path の順に描くと、`chartPlus` など線だけの path が塗りの path より先に定義されている図形で、線が塗りに隠れる。
- 矢じりは SVG の marker（部品ごとに一意な id。`/print` は全スライドを同時に描く。`line` の SVG と図形の SVG で同じ定義を使う）。**形は値ごと**: `arrow` は開いた V 字（線だけ）、`stealth` は後ろに切り欠きのある塗り、`triangle` は塗った三角、`oval` は円、`diamond` は菱形。**未知の値は `arrow`（開いた V 字）**で、書き出しの変換（未知の値を `arrow` にして `W-SHAPE`）と同じ形にする。**大きさは線幅のおよそ 3 倍、下限 6 px**（PowerPoint の既定の大きさ「中」（med）の近似）。
- 反転は SVG の中で、枠の中心に対して行う。
- 評価器の近似: 0 割りは 0、`sqrt` の負数は 0 にする。範囲内の `adj` でもそうなる定義がある（`circularArrow` の `adj5 = 0`、`curvedUpArrow`）。PowerPoint がそのとき何を描くかは確かめていない。
- 定義の誤記 1 件（`pie` の文字の枠。`t` と `r` の名前が入れ替わっていて枠が図形の外に出る）は、生成時に補正してある（`scripts/gen-presets.mjs` の `RECT_FIXES`）。PowerPoint の実際の挙動は確かめていない。

**文字の枠**。
- **`w` と `h` がどちらも props にあるときだけ**、定義の文字の枠（`rect` の l t r b。楕円なら内接する矩形、三角なら下半分）に文字を置き、その内側に `padding` を空ける。反転したときは、反転した形の上での位置に置く。
- どちらかが無ければ（幅か高さが内容で決まるときは）枠全体に置き、`padding` だけを空ける。文字の枠の余白は枠の大きさに比例し、枠の大きさは文字の量で決まるので循環するため。幅が内容で決まる図形で文字の枠を使うと、余白が幅に比例して増えて測り直しが止まらず、書き出しが壊れた（レビューで実測。幅が 3300 万 px に達した）。
- `flipV` のとき文字は 180 度回る。`valign` の上下も入れ替わって見える（PowerPoint と同じ）。

Capture: `kind: 'shape'`（`line` は `kind: 'line'`）。`shape` は `type` の名前のまま。`adj` は `data-ppt-opts` の値のまま（数でない値も入れ、捨てて `W-SHAPE` を出すのは変換の `normalizeAdjust`。収集で捨てると警告が出ない）、`flipH` `flipV` は `true` のときだけ、`line.head` `line.tail` は `'none'` を除いて `arrow` に入れる（native-export.md §2）。`rotate` があるとき `box` は回転前の枠（`data-ppt-box` と `data-ppt-opts.rotate` から復元。回転後の外接矩形は使わない）。`rotate` と実測の大きさを併用するときは、大きさが変わるたびに回転前の枠を測り直す（大きさの変わらない位置だけの移動は追わない）。
PPTX: 名前をそのまま `<a:prstGeom prst>` に出し、調整値は後処理で `<a:avLst>` に書く（native-export.md §3.2、§4.3）。PowerPoint は同じ定義で描くので、形が一致し、開いた後に調整ハンドルでも直せる。

`type` を PowerPoint の図形の名前にし、形を定義の数式から描く（docs/adr/0002）。
- 得るもの: PowerPoint の図形 187 種（吹き出し・フローチャート記号・コネクタ・星・アクションボタンなど）が 1 つの部品で使え、Slidev の画面と PowerPoint で形が一致する。調整値と反転も PowerPoint の図形と 1:1。
- 失うもの: `type` は `string` なので、綴り違いは型では止まらない。画面の `console.warn`（と `rect` の見た目）と、書き出しの `W-SHAPE` で気付く。`adj` のキーと単位は図形ごとの定義を読まないと分からない。
- 失うもの: `w` と `h` の両方を指定しない図形（`h` だけ書いた図形を含む）は、Slidev では文字が枠全体に置かれ、PowerPoint では定義の文字の枠に置かれるので、PowerPoint で開くと文字の位置が少しずれる（楕円や三角で目立つ）。
- 失うもの: 矢じりの大きさは PowerPoint の「中」の近似（線幅のおよそ 3 倍）で、細い線では下限 6 px のぶん PowerPoint より大きく見える。
- 失うもの: 形は定義どおりに描くだけで、見た目の補正はしない。`cloudCallout` は調整値で泡（吹き出しの先）を雲の近くに置くと、泡が雲に重なり、PowerPoint と描き分けが変わる（既知の差。`PptShape.vue` のコメント）。
- 失うもの: スライドからはみ出して枠を縮めた図形は、PowerPoint では縮めた枠に調整値が効くので、形が Slidev とずれる（native-export.md §4.3）。
- 失うもの（Slidev の画面と PPTX の既知の差。どれも直していない）:
  - `rotate` が 360 度以上: DOM は `rotate(390deg)` のまま描き、PPTX は `rot="1800000"`（30 度）に正規化される。見た目は同じ。
  - `w` か `h` が 0 の図形: PPTX に出ない。部品は SVG と文字の div を子に持つので §2.2 の 1「rect が空」には当たらず Capture に入り、変換の `clampBox` が `w <= 0 || h <= 0` を `dropped` にして飛ばす。そのとき `W-HIDDEN` は出るが、文面は「スライドの外にあるので飛ばした」で、図形名（`name`）が付かず `elementId` だけ（名前を決める前に出すため）。
  - `w` か `h` を省いた図形: Slidev の SVG の大きさは `offsetWidth` / `offsetHeight`（整数 px）で描き、PPTX の枠は実測の小数なので、1 px 未満ずれる（最大 0.39 px を実測）。

### 1.4 `PptImage`

| prop | 型 | 既定 | 意味 |
|---|---|---|---|
| `src` | `string` | 必須 | `/public` からの相対、絶対 URL、data URL |
| `alt` | `string` | `''` | 代替文字。PPTX の `descr` に入れる（PptxGenJS `altText`） |
| `fit` | `'contain' \| 'cover' \| 'fill'` | `'contain'` | 枠の中での収め方。CSS `object-fit` と PptxGenJS `sizing.type` に同じ値を渡す |

描画は `<img>` に `object-fit: fit`。Capture は `kind: 'image'`、`src` は絶対 URL に解決したもの（Node が fetch。SVG は Node が撮影に切り替える）。PPTX 側は `contain` / `cover` だけ `sizing` に変換し、`fill` は `sizing` 無し（native-export.md §4.3。`type: 'fill'` を渡すと `write()` が落ちる）。

`fit` を持たせて `<img>` そのままにしない。
- 得るもの: 写真を枠に合わせて切る書き方が Slidev と PPTX で同じになる。
- 失うもの: `sizing: crop` の切り出し位置（`x y` のオフセット）は使えない。要るなら prop を足す。

### 1.5 `PptTable`

| prop | 型 | 既定 | 意味 |
|---|---|---|---|
| `rows` | `string[][]` | なし | データで表を作るとき。1 行目が見出し |
| `header` | `boolean` | `true` | `rows` の 1 行目を見出し扱い（`thead` に描く）。**slot の Markdown 表には効かない**（markdown-it が作る `thead` はそのまま。`headerRows` は常に `thead > tr` の数で、収集器は `opts.header` を読まない） |
| `colW` | `number[]`（px） | 実測 | 列幅。指定すれば描画にも効かせる（slot の表には `<colgroup>` を差し込めないので、`table-layout: fixed` にして 1 行目のセルに幅を当てる） |
| `border` | `string \| { color, width }` | 実測 | 全セルの罫線 |
| `fill` | `string` | なし | 全セルの塗り |
| `size` `align` `valign` | `PptText` と同じ | 実測 | 全セル |

中身は default slot の Markdown 表か、`rows`。両方あれば slot を優先し `rows` は無視する。
セルの中は run の規則（§3.4）だけを使う（段落は 1 つ）。
Capture は `kind: 'table'`。`colW` は 1 行目の各セル、`rowH` は各 `tr` の実測（§3.5）。セルの罫線は computed の 4 辺を別々に取り、無い辺は `{ width: 0, color: '' }` で埋める。

slot と `rows` の両方を持つ。
- 得るもの: 手で書く表は Markdown、`v-for` で回す表は `rows` と使い分けられる。
- 失うもの: `rows` はセルの中に書式を持てない（文字列だけ）。両方あるときの優先も覚える必要がある。

### 1.6 ブラウザで測る方法: `data-ppt-*` 属性（登録簿は使わない）

部品は自分のルート要素に次の属性を出す。収集器は DOM だけを見る。

| 属性 | 値 |
|---|---|
| `data-ppt` | `text` `shape` `image` `table` |
| `data-ppt-export` | `native` `image` |
| `data-ppt-name` | `name` prop |
| `data-ppt-box` | 座標指定した辺だけの JSON（`{"x":100,"y":80}`）。無い辺は実測 |
| `data-ppt-opts` | 書式の props を JSON にしたもの（`type` `fill` `line` `radius` `rotate` `fit` `header` `valign` など、CSS から読み戻しにくいもの）。`PptShape` は `type` `fill` `line`（`head` `tail` を含む）`rotate` `padding` `valign` `align` を常に、`radius` は `type: 'roundRect'` のときだけ、`adj` は指定したときだけ、`flipH` `flipV` は `true` のときだけ出す |

| | 属性 | 登録簿（`window.__ppt__` に部品が登録） |
|---|---|---|
| 得るもの | `/print` で全スライドが同時に描画されても、要素と情報が同じ場所にある。SSR / クリック状態に依らない。DevTools で見える | props をそのまま（関数も）渡せる |
| 失うもの | JSON にできる props しか渡せない。DOM が少し太る | スライド番号と要素の対応を自分で持つ必要がある。`v-if` で消えた部品の登録解除を追う必要がある |

JSON にできない props は無いので、属性を選ぶ。

---

## 2. 収集器の歩き方

### 2.1 入口

1. `#print-content > .print-slide-container` を順に取る。`canvas.width/height` はこの要素の rect。`backgroundColor` はこの要素の computed。
2. その中の `[data-slidev-no]` から `no` と `lang`。`.slidev-page` の `scale` を `zoom` に。
3. `[data-slidev-no]` から子孫へ降りて**根の列**を作る。`.slidev-layout` に当たったらそれを根の 1 つにして、その下へは降りない。`.slidev-layout` でなく、子孫にも `.slidev-layout` を持たない要素のうち最上位のもの（`image-right` の右半分 `div[style*=background-image]` など）を根の 1 つにする。外側の `div.grid` は子孫に `.slidev-layout` を持つので根にならない。根の列は文書順。
4. 根が `.slidev-layout` のときは、**根自身の装飾も読む**。computed `background-image` に `url()` があれば画像要素にする（`layout: image` が根の style に背景を置く。ただし `cover` / `intro` の背景は `frontmatter.background` から Node が取るので（native-export.md §5.5）、根に `.cover` / `.intro` の class があれば読まない）。computed `background-color` が透明でなければ `SlideCapture.backgroundColor` を上書きする（`layout: end` の黒、`layoutClass` の塗り）。
5. 根ごとに、子を文書順に歩く（§2.2）。`.col-left` と `.col-right` に当たったら（`two-cols`、`two-cols-header`）、その 2 つを副領域にして別々に歩く（`.col-right` の要素は `roleHint` の候補が `body2` になる）。`two-cols-header` の `.col-header` / `.col-bottom` は副領域にせず、根の領域として歩く（見出しは `title` 候補のまま）。

### 2.2 要素の分類（歩く順に判定。最初に当たった行で決まる）

「装飾」= 背景色が透明でない、`border` 幅が 0 でない、`box-shadow` がある、`transform` がある、のいずれか。`background-image` は装飾に含めず 4 で別に扱う。

| # | 条件 | 扱い |
|---|---|---|
| 0 | Slidev の UI（`.slidev-code-copy`、`.slidev-icon`、`.slidev-icon-btn`、`SlideTop` / `SlideBottom` が描く要素、`[data-slidev-clicks-start]` の印だけの空要素） | 警告なしで飛ばす。コードブロックのコピーボタンは `opacity: 0` で全スライドにあるので、ここで消さないと雑音になる |
| 1 | 表示されていない（`display:none`、`visibility:hidden`、`opacity: 0`（祖先の積）、rect が空、キャンバスの外） | 飛ばす。`opacity: 0` とキャンバス外は `W-HIDDEN`、それ以外は無警告。「rect が空」は幅 0 かつ（高さ 0 か `hr` でない）かつ子が無い。`hr` は高さ 0〜1 px なので、`w === 0 \|\| h === 0` で判定すると 10 に届かない（実測） |
| 2 | `[data-ppt]` | PPT 部品（§1）。中に `[data-ppt]` があれば `W-NESTED-PPT` で無視 |
| 3 | `[data-ppt-export="image"]`（部品以外の要素に手で付けたもの） | 画像への置き換え、`reason: 'explicit'` |
| 4 | `background-image` を持つ要素（`img` 以外） | 子孫に可視テキストが無ければ: `url()` があれば画像要素（URL を取得。`linear-gradient` が重なっていれば捨てて `W-CSS`）、`url()` が無ければ画像への置き換え `reason: 'gradient'`。子孫に可視テキストがあれば: 要素ごと画像への置き換え `reason: 'explicit'`（文字も絵になるが見た目は合う。文字を編集したければ背景を `PptShape` か `PptImage` に分けて書く）。どちらも**中は歩かない** |
| 5 | `.katex-display`、`.mermaid`、`svg`、`canvas`、`iframe`、`video`、`audio`、`object`、`embed`。**`p` の中身が 5 か 6 の要素 1 つ（と空白）だけなら、その `p` ごと**（markdown-it は `![]()` を `<p><img></p>` に、Slidev はブロック数式を `div.slidev-katex-wrapper > p > span.katex-display` で描くので、`p` を 11 で拾うと中身に到達しない。実装 i0001-04 は 5 のタグと `.katex-display` / `.mermaid` で実装済み（`canvas` `iframe` などの `reason` は `'unknown-element'`）。`img` は未対応で、11 が先に当たって `W-INLINE` になる。次の実装で足す） | 画像への置き換え（`reason: 'math' \| 'mermaid' \| 'svg'`。`p` の中身が `img` なら 6 の画像要素）。`.mermaid` の中は Shadow DOM で歩けないが、撮影はホスト要素で足りる |
| 6 | `img` | 画像要素（`src` を取得。`a > img` なら `link` を要素に。§3.3） |
| 7 | `table` | 表要素（§3.5） |
| 8 | `pre` | コードブロック（§3.6）、区切りブロック |
| 9 | `blockquote` | 引用（§3.7）、区切りブロック |
| 10 | `hr` | 線要素（実測の上辺、computed `border-top-color`） |
| 11 | `h1`–`h6` `p` `ul` `ol` | 流し込みブロック（§3.1）。段落の中に `svg` `img` `table` `pre` `.katex-display` `.mermaid` などの区切りブロックがあれば、その要素は `W-INLINE` を出して無視する（画像への置き換えは枠の単位で、段落の途中には置けない） |
| 12 | `div` `section` `article` `main` `aside` `header` `footer` `span` `figure` `a`（および `strong` `em` `code` などの inline タグ）で、**装飾を持たない**。Slidev の `.slidev-layout a` は `border-b border-dashed` を持つので、画像を包む `a` は装飾ありとして 14 を通り（`W-CSS: border` が 1 つ出る）、中の `img` が 6 に届く。`a` の下線を装飾に数えない改善は次の実装の候補 | 透明な入れ物。中を歩く（`.col-left` `.col-right` もここ）。**子に可視のテキストノードや inline 要素を直接持つなら、文書順に、連続する inline の連なりごとに 1 つの `plain` 段落として流し込みブロックに数える**（ブロックの子が間に入れば、そこで連なりが切れる。markdown-it の HTML ブロックは `<p>` に包まれない。`slides.md` の `<span>Space / → で次へ</span>` や `<div v-click>クリックのたびに…</div>` がこれ） |
| 13 | 12 のタグで装飾を持ち、**装飾を持たない入れ物（12）だけを通って**到達する子孫が 11・裸のテキスト・inline 要素だけ（子が無い場合も含む。ただし 4 で先に拾われる） | 自前のテキスト枠（区切りブロック）。中の入れ物は透明に抜けて段落になる。装飾は `frame` に変換し、変換できないものは `W-CSS` |
| 14 | 12 のタグで装飾を持ち、子孫に 3–10 か 13（装飾つきの入れ物）がある | 装飾を `W-CSS` で捨てて、透明な入れ物として中を歩く |
| 15 | それ以外（`button` `input` `details` `kbd` を含む未知タグ、Vue 部品が描く任意の要素）。`kbd` は inline で現れたときは §3.4 の run になり、ブロックとして現れたときだけここ | 画像への置き換え、`reason: 'unknown-element'` |

- 得るもの: Tailwind の `div` 入れ子（`grid`、`pt-12`）は透明に抜け、中の文字はネイティブになる。色付きの箱は塗りのあるテキスト枠になる。
- 失うもの: 14 のとき、箱の見た目（背景・角丸）が消える。`slides.md` の「アニメーション」スライドの 3 つの箱（`div.p-6.rounded-lg.bg-blue-500/20 > div.text-3xl + div`）は、子が装飾の無い `div` なので 13 に当たり、2 段落のテキスト枠として残る。

`slides.md` を通した確認（実装フェーズの e2e の期待値の元）:
- 表紙: `.slidev-layout.cover > div.my-auto > h1 + p + div.pt-12 > span` → `h1` は 12 を通って到達するので `title` 候補。`p` と、`div.pt-12`（12）の裸テキスト `span` の文は、間に区切りブロックが無いので**同じ枠**（`body` 候補）の 2 段落。入れ物の境界では枠を閉じない。
- 「Slidev とは」: `h2` → title 候補、`ul`（`v-clicks` はラッパを作らない）→ body 候補。
- 「コードのハイライト」: `h2` → title 候補、`p` → body 候補、`div.slidev-code-wrapper`（12）> `pre`（8）→ 自由配置の code 枠、`button.slidev-code-copy`（0）→ 飛ばす。
- 「数式」: `p`（インライン `.katex` を含む）→ body 候補に `W-MATH-INLINE`、`div.slidev-katex-wrapper > p > .katex-display` → 5 で画像。
- 「Vue コンポーネント」: `div.mt-6 > button` → 12 を通って `button` が 15 で画像。

### 2.3 位置と大きさ

- `box = elementRect − containerRect`。`getBoundingClientRect()` は変換後の値を返すので、`.slidev-page` に `scale`（frontmatter `zoom`）があっても **rect はそのままキャンバス px**（`SlideWrapper.vue` は幅を `100% / zoom` に広げてから `scale` で縮めている）。
- 逆に **`getComputedStyle` から取る長さ**（`font-size` `line-height` `margin` `padding` `border-width` `border-radius` `letter-spacing`）は変換前の値なので、`zoom` を**掛ける**。
- 回転した要素（CSS `transform: rotate`）は rect が回転後の外接矩形になる。PPT 部品以外の回転は再現せず `W-CSS`（`box` は外接矩形のまま）。`PptShape rotate` は §1.3 の方法で回転前の枠を使う。
- 流し込みブロックの枠（§3.1）は、含むブロックの border-box の和集合（左端の最小、右端の最大、最初の上端、最後の下端）。段落の margin は枠に入れず、`spaceBefore/After` に変換する。
- `position: fixed` の要素はコンテナ基準になる（`translate-0` のため）。`fixed` を使った部品も同じ式で測れる。

---

## 3. 普通の Markdown の変換規則

### 3.1 流し込みブロックとテキスト枠

連続する流し込みブロック（`h1`–`h6` `p` `ul` `ol`、および 2.2 の 13 の中身）を 1 つのテキスト枠にまとめる。
区切りブロック（表・コードブロック・引用・画像・置き換え・PPT 部品・装飾つきの箱）が来たら枠を閉じ、次の流し込みブロックから新しい枠を開く。

`roleHint`（候補）の決め方。収集器は DOM の事実だけを記録し、placeholder に入れるかどうかは Node が決める（native-export.md §4.2。収集器はレイアウト名も対応表も知らない）:
- スライドの中で文書順に最初の `h1` か `h2` で、根（または `.col-left`）から**装飾を持たない入れ物（§2.2 の 12）だけを通って**到達するもの → `title` 候補。この見出しだけで 1 枠。theme-default の `cover` と client の `center` は slot を `div.my-auto` で包むので、「直下」に限ると表紙の見出しが候補にならない。
- **`title` の直後に、区切りブロックを挟まずに**開く枠（同じ領域）→ `body` 候補。間に区切りブロックがあれば `body` 候補は無い（placeholder の固定位置に飛んで、実測配置の要素と重なるため）。
- `.col-right` の最初の要素が流し込みブロックなら、その枠 → `body2` 候補。最初の要素が区切りブロックなら右列に候補は無い。
- それ以外の枠には `roleHint` を付けない（自由配置）。
- Node 側: レイアウトが対応表に無い（`blank`）か、対応表のそのレイアウトに同名の placeholder が無ければ、候補は捨てて自由配置にする。

- 得るもの: 収集器が「測って記録する」に保たれ、対応表の変更が Node に閉じる。
- 失うもの: Capture を見ただけでは placeholder に入るか分からない。受入テストの fixture は Capture と `options.data` 相当を対で持つ。

親チケットの「見出しはタイトル枠」は主見出し 1 つを指すと解した。2 つ目以降の見出しと h3 以降は、枠の中の見出し段落になる。

例（`slides.md` の「始め方」、two-cols）:
```
.col-left : h2「始め方」→ title 候補 / pre → 自由配置の code 枠
.col-right: h2「画面」+ ul → body2 候補（見出しは body2 の中の heading 段落）
```
`body` 候補は無い。

例（「Vue コンポーネントを埋め込む」）:
```
h2 → title 候補 / p → body 候補 / div.mt-6 > Counter(button) → image(unknown-element) / pre → 自由配置の code 枠 / p → 自由配置
```

- 得るもの: 段落と箇条書きが 1 つの枠で流れるので、PowerPoint で文を足すと下の行が押し下がる。
- 失うもの: 区切りブロックの後の文は別の枠になり、前の枠で文を足しても後ろは動かない。

### 3.2 ブロックの規則

| Slidev / HTML | Paragraph | 備考 |
|---|---|---|
| `h1`–`h6` | `kind: 'heading'`, `level: n` | 大きさは run の `size`（computed）。太さも computed（Slidev の見出しは太字でない） |
| `p` | `'plain'` | `my-4` は `spaceBefore/After` に |
| `ul > li` | `'bullet'`, `level: 入れ子の深さ` | `li` の直下のテキストと inline を 1 段落。`li > p` があれば `p` ごとに段落（最初だけ箇条書き記号、続きは同じ `level` の `plain`）。字下げ幅は測らず PptxGenJS の `indentLevel` に任せる（native-export.md §4.2） |
| `ol > li` | `'number'`, `numberStart: ol の start` | 同上 |
| `li > ul/ol` | `level + 1` | 入れ子 |
| `li > pre` | `'code'` の段落として同じ枠に流す | |
| `li > table / img / blockquote / .katex-display` | 無視して `W-LI-BLOCK` | 箇条書きの中で枠を割らない |
| `br` | 前の run に `breakAfter` | |
| 半透明のブロック（theme-default の `h1 + p { opacity-50 }` など） | 特別扱いしない。§3.4 の computed `opacity` の式で run の `transparency` になる（0.5 なら 50） | |

### 3.3 区切りブロックの規則

| Slidev / HTML | Element | 備考 |
|---|---|---|
| `table` | table | §3.5 |
| `pre`（Shiki の `.slidev-code`） | text（`kind: 'code'` の段落、1 行 1 段落）、`frame.fill` = computed 背景色、`frame.inset` = padding、`frame.radius` | §3.6 |
| `blockquote` | text、`frame.fill` = computed 背景色 | §3.7 |
| `img` | image | `src` 絶対化。`a > img` なら `link` を要素に |
| `.katex-display` | image（`reason: 'math'`） | 2 倍で撮る |
| `.mermaid` `svg` | image（`reason: 'mermaid' \| 'svg'`） | 同上 |
| 装飾つきの箱（2.2 の 13） | text | `frame` に背景色・border・角丸・padding・opacity |
| `background-image` の要素（2.2 の 4、§2.1 の根自身） | image | `image-right` の右半分、`layout: image` の根など。`cover` / `intro` の背景は `frontmatter.background` から Node が取るので収集器は見ない（native-export.md §5.5） |

### 3.4 run の規則（inline）

| HTML | Run |
|---|---|
| テキストノード | `text`（連続する空白は 1 つに。`pre` の中は保つ）。畳む空白は CSS の `white-space: normal` と同じ `[ \t\r\n\f]` だけ。全角空白 U+3000 と NBSP は文字として保つ（`\s` で畳むと「全角　空白」が潰れる。実測） |
| `strong` `b`、computed `font-weight >= 600` | `bold` |
| `em` `i`、computed `font-style: italic` | `italic` |
| `code` `kbd` | `code: true`、`highlight` = computed 背景色 |
| `a[href]` | `href` は属性の生の値で読む（`a.href` は正規化で末尾に `/` が付く）。`#N` `##N` `/N` なら `{ slide: N }`（Slidev は `[x](#3)` を `href="##3"` にする。実測）。スキーム付き（`https:` `mailto:` など）なら `link: { url }`。それ以外（相対パス、数字でないアンカー）は開発サーバの URL に解決されて PPTX に残るので外し、`W-LINK` |
| `del` `s` | `strike` |
| `u`、computed `text-decoration-line` に underline | `underline`（`a` の下線は除く。Slidev の `a` は border-bottom で下線を描く。ただし `a` の中の `u` は本物の下線） |
| `sup` `sub` | `sup` / `sub` |
| `mark` | `highlight` = computed 背景色 |
| `.katex`（インライン数式） | `.katex-html` の `textContent` を 1 run（`italic`）。`W-MATH-INLINE` |
| computed `color` | `color`（親と同じでも毎回入れる。Node 側で省く）。`rgb()` `rgba()` `#` 以外の形式（`oklch()` `color()` など）は読めないので**黒**にして `W-CSS`（スライドごとに 1 回）。`transparent` とアルファ 0 の色も黒にする（警告なし）。白にすると白地で見えなくなる |
| computed `opacity`（祖先の積）と `color` のアルファ | `transparency = 100 − opacity × 100`。Node 側で PptxGenJS に渡す直前に整数に丸める（`<a:alpha>` は整数） |
| computed `letter-spacing` | `charSpacing` |
| `span` `em` などの入れ子 | 平坦化。書式は computed から run ごとに読む |
| `text-transform` | 無視して `textContent` のまま。`W-CSS` |

### 3.5 表

- `thead > tr` の数が `headerRows`。`thead` が無ければ 0。
- セルは `td` / `th` を 1 段落。`th` は computed の太さに従う（Slidev の `th` は `font-400`）。
- `colspan` `rowspan` はそのまま。
- 罫線は computed の `border-{top,right,bottom,left}-{width,color}` を 4 辺別に。無い辺は `{ width: 0, color: '' }`。Slidev 既定は `tr` の `border-b` だけ → 下辺のみ（下辺だけの出力は実測で確認済み）。
  - セルの computed では 4 辺とも 0 になる（Slidev は `tr { border-b }` で引く。実測）。セルの辺が 0 なら、上下は `tr`、左右は `table` の同じ辺へ遡って読む。`tr` の色は `rgba(…, 0.2)` のような半透明なので、色はアルファを落として `#rrggbb` にする（罫線に透明度は渡さない）。
- `caption` は表の要素にせず、表の**前**の自由配置の段落（テキスト枠、`inset` 0）にする。PptxGenJS の表にキャプションは無い。
- `inset` はセルの padding（Slidev は `p-2 py-3`）。
- `align` は computed `text-align`、`valign` は `vertical-align`。
- `colW` は 1 行目の各セルの rect 幅、`rowH` は各 `tr` の rect 高さ。

### 3.6 コードブロック

- `pre .line` を 1 行 1 段落（`kind: 'code'`）。`.line` が無ければ `textContent` を改行で割る。
- 色付け（Shiki の `span style="--shiki-light: …"`。色は CSS 変数経由で、`emulateMedia({ colorScheme: 'light' })` が効いている前提で computed に解決される）は捨てて、`pre` の computed `color` を全 run に使う。行番号（`lineNumbers: true`）と行強調（`{2|4-6}`。print では全行に `slidev-code-highlighted highlighted` が付くので、この class は無視する）は捨てる。行頭の空白は保つ（native-export.md §11 の 2）。
- 行強調で強調されない行は `.slidev-code-dishonored`（`opacity: 0.3`。実測）になる。コード枠の中では §3.4 の opacity → `transparency` を**当てない**。行強調は捨てる規則なので、薄い行を作らない。
- `frame.fill` は `pre` の computed 背景色、`frame.inset` は padding、`frame.radius` は border-radius。
- `W-CSS` は出さない。色付けを捨てるのは規則（親チケットの決定）で、`Report.dropped['code-highlight']` に件数だけ出す。
- 画像で出したいときは `<div data-ppt-export="image">` で囲む。

### 3.7 引用

- `blockquote` の中の `p` を段落に。`frame.fill` = computed 背景色、`frame.inset` = padding、`frame.radius`。
- Slidev の左の縦線（`border-l`）は出さない（PptxGenJS の `line` は 4 辺共通）。`W-CSS` は出さず `Report.dropped['blockquote-border']` に件数を出す。
- 得るもの: 引用のたびに同じ警告が出ない。
- 失うもの: 縦線が消えたことは件数でしか分からない。縦線を出したければ、引用の左に `line` 要素を 1 本足す変換を後で入れる。

### 3.8 `::right::`（2 段組）

`.col-right` を第 2 の領域として歩く。最初の枠が `body2` 候補（`native-export.md` §5.2 の placeholder）、残りは自由配置。

---

## 4. 画像への置き換え

### 4.1 対象

2.2 の 3、4（グラデーションだけの要素と、文字を含む背景画像の箱）、5、15。`img` の SVG（Node が判定）。
`PptText` 等に `export="image"` を付けた部品もここ。

### 4.2 撮り方

1. 収集器が要素に `data-ppt-capture-id="sN-eM"` を付ける。
2. Node が `page.locator('[data-ppt-capture-id="sN-eM"]').screenshot({ type: 'png', omitBackground: true, animations: 'disabled' })` で撮る。context の `deviceScaleFactor: 2` で 2 倍解像度になる。
3. 枠は要素の実測 `box`（`overflow: visible` で外にはみ出た描画は写らない。要素に `padding` を足すのは使う側の責任）。
4. 撮影に失敗したら `W-IMAGE`、灰色の矩形。

- 得るもの: 部分の撮影なので、周りの要素が写り込まない。
- 失うもの: 要素の外にはみ出す装飾（影、absolute の子）は切れる。`omitBackground` が祖先の塗りを消さない可能性がある（native-export.md §11 の 5。消えなければ撮影の間だけ祖先の背景を透明にする）。

### 4.3 記録

`Report.replacements` に `slide` `elementId` `name` `reason` `selector`（タグ + class の先頭 2 つ）`width height`（撮った px）。

---

## 5. 発表者ノート

`native-export.md` §4.4。`<!-- -->` の中身（`SlideInfo.note`）を行ごとの段落に。`[click]` 印は消す。

---

## 6. `slides.md` での使い方（実装フェーズで例に入れる）

````md
<PptText :x="60" :y="380" :w="400" fill="#eef" :radius="8" :padding="12">

**実測配置**が既定で、`x` `y` を書けば座標指定になります。

</PptText>

<PptShape type="rightArrow" :x="500" :y="400" :w="120" :h="40" fill="#3b82f6" line="none" />

<PptImage src="/photo.jpg" :x="640" :y="120" :w="280" :h="200" fit="cover" />

<PptTable :rows="[['名前','用途'],['default','通常'],['cover','表紙']]" />

<div data-ppt-export="image">

```ts
// この部分だけ画像で出す
```

</div>
````

---

## 7. 要確認（`native-export.md` §11 に加えて）

8. Slidev の `<PptText>` slot に Markdown を書いたとき、markdown-it の HTML ブロック規則で中身が描画されること（空行の要否）。Slidev の `<div>` の流儀と同じはずだが、大文字始まりの部品名で同じかを確かめる。
