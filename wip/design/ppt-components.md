# PPT 部品と Markdown 変換の設計（i0001-02）

`native-export.md` の §2（Capture）に何をどう入れるかを決める。置き場は `packages/slidev-addon-pptx/components/`。
`slides.md` からは `addons: [slidev-addon-pptx]` で読み込み、`<PptText>` のように使う。

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

### 1.2 `PptText`

| prop | 型 | 既定 | Capture |
|---|---|---|---|
| `align` | `'left' \| 'center' \| 'right' \| 'justify'` | 実測 | 全段落の `align` |
| `valign` | `'top' \| 'middle' \| 'bottom'` | `'top'` | `valign` |
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
| `type` | `'rect' \| 'roundRect' \| 'ellipse' \| 'line' \| 'rightArrow' \| 'leftArrow' \| 'upArrow' \| 'downArrow' \| 'diamond' \| 'triangle' \| 'hexagon'` | `'rect'` | PptxGenJS `ShapeType` の名前をそのまま使う。実装はこの 11 種から始め、増やすときは対応する SVG を足す |
| `fill` | `string \| 'none'` | `'#ffffff'` | 塗り。`'none'` で塗りなし |
| `line` | `string \| 'none' \| { color, width, dash, head, tail }` | `'#000000'`（1 px） | 線。`'none'` で線なし。`head` `tail` は `'none' \| 'arrow' \| 'triangle' \| 'oval' \| 'diamond'`（`type: 'line'` のとき） |
| `radius` | `number` | 8 | `roundRect` の角丸 px |
| `rotate` | `number`（度） | 0 | 回転 |
| `align` `valign` `size` `color` | `PptText` と同じ | | 図形の中の文字 |

中身は default slot（図形の中の文字。段落の規則は `PptText` と同じ）。
描画: `rect` `roundRect` は `div` に border / background / border-radius。それ以外は inline SVG（`<svg viewBox>` に `path`）を背景に敷き、文字は上に重ねる。`type: 'line'` は `x y w h` で対角線を引く（`w` か `h` が 0 なら水平・垂直）。
Capture: `kind: 'shape'`（`line` は `kind: 'line'`）。`rotate` があるとき `box` は回転前の枠（`data-ppt-box` と `data-ppt-opts.rotate` から復元。回転後の外接矩形は使わない）。

- 得るもの: PowerPoint の図形と 1:1。
- 失うもの: 11 種以外の図形（吹き出し、フローチャート記号）は使えない。要るものが出たら 1 種ずつ足す。

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
| `header` | `boolean` | `true` | 1 行目を見出し扱い（太字、塗り） |
| `colW` | `number[]`（px） | 実測 | 列幅。指定すれば `<col>` で描画にも効かせる |
| `border` | `string \| { color, width }` | 実測 | 全セルの罫線 |
| `fill` | `string` | なし | 全セルの塗り |
| `size` `align` `valign` | `PptText` と同じ | 実測 | 全セル |

中身は default slot の Markdown 表か、`rows`。両方あれば slot を優先し `rows` は無視する。
セルの中は run の規則（§3.4）だけを使う（段落は 1 つ）。
Capture は `kind: 'table'`。`colW` `rowH` は `<col>` と `<tr>` の実測。セルの罫線は computed の 4 辺を別々に取り、無い辺は `{ width: 0, color: '' }` で埋める。

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
| `data-ppt-opts` | 書式の props を JSON にしたもの（`type` `fill` `line` `radius` `rotate` `fit` `header` `valign` など、CSS から読み戻しにくいもの） |

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
4. 根ごとに、子を文書順に歩く（§2.2）。根が `.slidev-layout` で直下に `.col-left` と `.col-right` があれば（`two-cols`）、その 2 つを副領域にして別々に歩く（`.col-right` の要素は `roleHint` の候補が `body2` になる）。

### 2.2 要素の分類（歩く順に判定。最初に当たった行で決まる）

「装飾」= 背景色が透明でない、`border` 幅が 0 でない、`box-shadow` がある、`transform` がある、のいずれか。`background-image` は装飾に含めず 4 で別に扱う。

| # | 条件 | 扱い |
|---|---|---|
| 0 | Slidev の UI（`.slidev-code-copy`、`.slidev-icon`、`.slidev-icon-btn`、`SlideTop` / `SlideBottom` が描く要素、`[data-slidev-clicks-start]` の印だけの空要素） | 警告なしで飛ばす。コードブロックのコピーボタンは `opacity: 0` で全スライドにあるので、ここで消さないと雑音になる |
| 1 | 表示されていない（`display:none`、`visibility:hidden`、`opacity: 0`（祖先の積）、rect が空、キャンバスの外） | 飛ばす。`opacity: 0` とキャンバス外は `W-HIDDEN`、それ以外は無警告 |
| 2 | `[data-ppt]` | PPT 部品（§1）。中に `[data-ppt]` があれば `W-NESTED-PPT` で無視 |
| 3 | `[data-ppt-export="image"]`（部品以外の要素に手で付けたもの） | 画像への置き換え、`reason: 'explicit'` |
| 4 | `background-image` を持つ要素（`img` 以外） | 子孫に可視テキストが無ければ: `url()` があれば画像要素（URL を取得。`linear-gradient` が重なっていれば捨てて `W-CSS`）、`url()` が無ければ画像への置き換え `reason: 'gradient'`。子孫に可視テキストがあれば: 要素ごと画像への置き換え `reason: 'explicit'`（文字も絵になるが見た目は合う。文字を編集したければ背景を `PptShape` か `PptImage` に分けて書く）。どちらも**中は歩かない** |
| 5 | `.katex-display`、`.mermaid`、`svg`、`canvas`、`iframe`、`video`、`audio`、`object`、`embed`。**`p` の中身が `.katex-display`（と空白）だけなら、その `p` ごと**（Slidev はブロック数式を `div.slidev-katex-wrapper > p > span.katex-display` で描くので、`p` を 11 で拾うと数式に到達しない） | 画像への置き換え（`reason: 'math' \| 'mermaid' \| 'svg'`）。`.mermaid` の中は Shadow DOM で歩けないが、撮影はホスト要素で足りる |
| 6 | `img` | 画像要素（`src` を取得） |
| 7 | `table` | 表要素（§3.5） |
| 8 | `pre` | コードブロック（§3.6）、区切りブロック |
| 9 | `blockquote` | 引用（§3.7）、区切りブロック |
| 10 | `hr` | 線要素（実測の上辺、computed `border-top-color`） |
| 11 | `h1`–`h6` `p` `ul` `ol` | 流し込みブロック（§3.1） |
| 12 | `div` `section` `article` `main` `aside` `header` `footer` `span` `figure` で、**装飾を持たない** | 透明な入れ物。中を歩く（`.col-left` `.col-right` もここ）。**子に可視のテキストノードや inline 要素を直接持つなら、文書順に、連続する inline の連なりごとに 1 つの `plain` 段落として流し込みブロックに数える**（ブロックの子が間に入れば、そこで連なりが切れる。markdown-it の HTML ブロックは `<p>` に包まれない。`slides.md` の `<span>Space / → で次へ</span>` や `<div v-click>クリックのたびに…</div>` がこれ） |
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
| `background-image` の要素（2.2 の 4） | image | `image-right` の右半分など。`cover` の背景は `.slidev-layout` 自身の style なので収集器は見ず、Node が `frontmatter.background` から取る（native-export.md §5.5） |

### 3.4 run の規則（inline）

| HTML | Run |
|---|---|
| テキストノード | `text`（連続する空白は 1 つに。`pre` の中は保つ） |
| `strong` `b`、computed `font-weight >= 600` | `bold` |
| `em` `i`、computed `font-style: italic` | `italic` |
| `code` `kbd` | `code: true`、`highlight` = computed 背景色 |
| `a[href]` | `link: { url }`。`href` が `#N` か `/N` なら `{ slide: N }` |
| `del` `s` | `strike` |
| `u`、computed `text-decoration-line` に underline | `underline`（`a` の下線は除く。Slidev の `a` は border-bottom で下線を描く） |
| `sup` `sub` | `sup` / `sub` |
| `mark` | `highlight` = computed 背景色 |
| `.katex`（インライン数式） | `.katex-html` の `textContent` を 1 run（`italic`）。`W-MATH-INLINE` |
| computed `color` | `color`（親と同じでも毎回入れる。Node 側で省く） |
| computed `opacity`（祖先の積） | `transparency` |
| computed `letter-spacing` | `charSpacing` |
| `span` `em` などの入れ子 | 平坦化。書式は computed から run ごとに読む |
| `text-transform` | 無視して `textContent` のまま。`W-CSS` |

### 3.5 表

- `thead > tr` の数が `headerRows`。`thead` が無ければ 0。
- セルは `td` / `th` を 1 段落。`th` は computed の太さに従う（Slidev の `th` は `font-400`）。
- `colspan` `rowspan` はそのまま。
- 罫線は computed の `border-{top,right,bottom,left}-{width,color}` を 4 辺別に。無い辺は `{ width: 0, color: '' }`。Slidev 既定は `tr` の `border-b` だけ → 下辺のみ（下辺だけの出力は実測で確認済み）。
- `inset` はセルの padding（Slidev は `p-2 py-3`）。
- `align` は computed `text-align`、`valign` は `vertical-align`。
- `colW` は 1 行目の各セルの rect 幅、`rowH` は各 `tr` の rect 高さ。

### 3.6 コードブロック

- `pre .line` を 1 行 1 段落（`kind: 'code'`）。`.line` が無ければ `textContent` を改行で割る。
- 色付け（Shiki の `span style="--shiki-light: …"`。色は CSS 変数経由で、`emulateMedia({ colorScheme: 'light' })` が効いている前提で computed に解決される）は捨てて、`pre` の computed `color` を全 run に使う。行番号（`lineNumbers: true`）と行強調（`{2|4-6}`。print では全行に `slidev-code-highlighted highlighted` が付くので、この class は無視する）は捨てる。行頭の空白は保つ（native-export.md §11 の 2）。
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
