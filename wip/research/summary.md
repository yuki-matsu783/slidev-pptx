# 調査まとめ（i0001-01）

対象: PptxGenJS 4.0.1 / jszip 3.10.2 / Slidev 52.19.1（@slidev/cli, client, parser, types）/ playwright-chromium 1.63.0。
パスの `<cli>` `<client>` `<types>` `<parser>` は `node_modules/.pnpm/@slidev+<名前>@52.19.1_.../node_modules/@slidev/<名前>` の略。
`<pg>` は `node_modules/.pnpm/pptxgenjs@4.0.1/node_modules/pptxgenjs` の略。
試作の出力（`proto1.pptx` `proto1-patched.pptx` と取り出した XML）はスクラッチパッドの `research/` にある。

---

## 1. PptxGenJS 4.0.1 の出力に XML を足す方法

### 分かったこと

**ZIP ライブラリと書き出し経路**
- PptxGenJS は `jszip`（依存指定 `^3.10.1`、実体 3.10.2）を `require('jszip')` で使う。ES ビルドも `import JSZip from 'jszip'`。
- `write({ outputType: 'nodebuffer' })` は内部で `new JSZip()` に全パートを `zip.file()` して `zip.generateAsync({ type: outputType })` を返すだけ。`compression` は `outputType` 指定時には無視され、jszip 既定（STORE）になる。
- 出力後の Buffer を `JSZip.loadAsync(buf)` で開き、`zip.file(path).async('string')` で読み、`zip.file(path, xml)` で上書きし、`zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })` で再保存できる。試作で slide2.xml の `</p:spTree>` 直前に `<p:grpSp>` を足したものが python-pptx で「GROUP」として読めた（PowerPoint 実機での開封は未確認）。
- 出力 ZIP は 2 スライド・画像 1 枚で 46 エントリ。うち 19 は `zip.folder()` による空ディレクトリエントリで、パートは 27。エントリ順は `_rels/` `docProps/` `ppt/`… のディレクトリが先で、`[Content_Types].xml` は先頭ではない（OPC 上は必須ではないが、PowerPoint 生成物とは並びが違う）。パートの一覧: `[Content_Types].xml` `_rels/.rels` `docProps/{app,core}.xml` `ppt/presentation.xml` `ppt/_rels/presentation.xml.rels` `ppt/{presProps,viewProps,tableStyles}.xml` `ppt/theme/theme1.xml` `ppt/slideMasters/slideMaster1.xml`(+rels) `ppt/slideLayouts/slideLayoutN.xml`(+rels) `ppt/slides/slideN.xml`(+rels) `ppt/notesMasters/notesMaster1.xml`(+rels) `ppt/notesSlides/notesSlideN.xml`(+rels) `ppt/media/image-<slide>-<n>.png`。
- スライドマスターは常に 1 枚（`slideMaster1.xml` 固定）。レイアウトは `defineSlideMaster` の数 + 1 で、`slideLayout1.xml` は常に組み込みの `DEFAULT`。ユーザー定義は 2 番から。`DEFAULT` は図形なしだが空ではなく、`<p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>` が必ず入り、`slide.slideNumber = {...}` を使うと `setSlideNumber` がこの `DEFAULT` レイアウトにもスライド番号 placeholder を足す。
- `<p:cNvPr id>` はスライド内で重複しうる。テキスト・画像・図形は `idx + 2`（spTree 内の順番）、表は `intTableNum * slideNum + 1` で採番するため別系列になり、試作の slide2 では `id="3"` が `Text 1` と `Table 0` で重複した。`name` も `Text 0` が 3 つ（自由配置 1 つ + 自動追加 placeholder 2 つ）。python-pptx はどちらも検出しない。
- `[Content_Types].xml` にはスライド数ぶん `/ppt/slideMasters/slideMaster{idx}.xml` の Override が出る（実在しない `slideMaster2.xml` を指す）。python-pptx は読めた。PowerPoint も通ることが多いと思われるが未確認（推測）。

**`defineSlideMaster` / `placeholder` で出せるもの**
- `SlideMasterProps.objects` に置けるのは `chart` `image` `line` `rect` `text` `placeholder` の 6 種。`placeholder` は `{ options: PlaceholderProps, text? }` で、`options.type` は `'title' | 'body' | 'pic' | 'chart' | 'tbl' | 'media'`。ただし実装はテキスト系のみで、`pic` 等はテキスト枠として書かれ、画像プレースホルダーには入らない（ソースの TODO コメント）。
- マスター側の `background`（色・画像）、`slideNumber`（位置・フォント）、`margin` は使える。
- レイアウト XML には各 placeholder が `<p:sp>` + `<p:nvPr><p:ph idx="100+n" type="title|body"/></p:nvPr>` で出る。`idx` は `100 + objects 配列内の順番`。位置は `<a:xfrm>` に EMU で入る。
- スライド側 `addText(text, { placeholder: 'name' })` は、レイアウトの同名 placeholder から x/y/w/h と一部書式を継承し、`<p:ph idx=... type=... hasCustomPrompt="1"/>` 付きの `<p:sp>` を出す。
- **`addSlide({ masterName })` で作ったスライドが埋めなかった placeholder は、書き出し時に空テキストの placeholder 図形として自動追加される**（`addPlaceholdersToSlideLayouts`）。試作のスライド 2（自由配置のみ）にも title/body の空枠が入った。`masterName` なしの `addSlide()` は `_slideObjects` を持たない臨時の `DEFAULT` レイアウト参照になるので何も足されない。
- 箇条書き: `bullet: true` と `indentLevel: n` で `<a:pPr marL=... indent=...>` + `<a:buChar>`。`lvl` 属性は `indentLevel >= 1` のときだけ出る（0 では出ない）。階層は 1 段目 `marL=342900`、以後 `+342900` ずつ。
- 表: `addTable` → `<p:graphicFrame>` + `<a:tbl>`。セルごとに `<a:rPr>` の中に `<a:latin>/<a:ea>/<a:cs>` が入る。セル結合は `colspan`/`rowspan` → `gridSpan`/`rowSpan` と、埋め草セルの `hMerge`/`vMerge` で出せる。
- 座標: 数値は 100 未満をインチ、100 以上を EMU とみなして素通しする（`getSmartParseNumber`。`inch2Emu` は `> 100` で素通し）。丸めは `Math.round(914400 * inches)`。`'50%'` はスライド幅/高さに対する割合。
- 画像: `addImage({ data: 'data:image/png;base64,...' })` → `<p:pic>` + `ppt/media/image-<slideNo>-<n>.png` + スライド rels の `image` 関係。
- 図形: `addShape(pptx.ShapeType.roundRect, ...)` → `<p:sp>` + `<a:prstGeom prst="roundRect">`。
- ノート: `slide.addNotes(str)` → `notesSlideN.xml` の `<p:ph type="body" idx="1">` に生テキスト。改行は `\n` を CRLF に置き換えたうえで 1 つの `<a:t>` に入る（`<a:p>` にも `<a:br/>` にもならない）。段落に分けたければ後処理で `<a:p>` に割る（PowerPoint 実機で CRLF がどう見えるかは未確認）。
- ハイパーリンク: run の `hyperlink: { url, tooltip }` → `<a:hlinkClick r:id="rIdN" ...>` とスライド rels の `hyperlink` 関係（`TargetMode="External"`）。`hyperlink: { slide: n }` で内部リンク（rels の Target がスライド番号）も出せる。どちらも run に `<ahyp:hlinkClr val="tx"/>` が付き、テーマのリンク色ではなく文字色で表示される。
- run の `lang` は既定 `en-US`（`TextBaseProps.lang` で個別指定可）。

**出せないもの**
- グループ化: 型にも実装にも API がない（`grpSp` の出現は spTree のルート 4 箇所のみ）。足すなら `ppt/slides/slideN.xml` の `<p:spTree>` 直下に `<p:grpSp>`（`<p:grpSpPr><a:xfrm>` に off/ext/chOff/chExt）を書き、子 `<p:sp>` を中に入れる。`cNvPr id` はスライド内で一意にする。
- アニメーション: API なし。足すなら `slideN.xml` の `</p:cSld>` の後、`<p:clrMapOvr/>` の後ろに `<p:timing>` を書く。対象図形は `cNvPr id` で `<p:spTgt spid="..."/>` から指す。
- 画面切り替え: `<p:transition>` も API なし（ソースに出現しない）。Slidev の `transition` frontmatter は落ちる。足すなら `</p:cSld>` `<p:clrMapOvr/>` の後、`<p:timing>` の前。
- 自動縮小 `fit`: `'shrink'` → `<a:bodyPr ...><a:normAutofit/></a:bodyPr>`（`fontScale` なし。ソースのコメントは「PowerPoint がリサイズ時に `fontScale` を動的に計算する」で、書き出し直後は縮まない）。`'resize'` → `<a:spAutoFit/>`。`'none'` と未指定 → 何も出ない（`noAutofit` は PPT2013 で問題が出るため意図的に出さない）。縮小率を固定したいなら `<a:normAutofit fontScale="85000" lnSpcReduction="20000"/>` を後から書く。

### 設計に効く制約

- 後処理は「PptxGenJS で出す → jszip で開く → 文字列置換または DOM 操作 → 再圧縮」の一本道。PptxGenJS の中間オブジェクトに割り込む公開 API はない。
- 図形の `cNvPr id` と `name`（`Text 0`, `Shape 4`, `Image 0` など）は PptxGenJS が採番し、**表があると `id` がスライド内で重複する**。重複 id は PowerPoint の「修復」対象になりうるので、後処理で id をスライド内一意に振り直すのを必須項目にする。図形の特定は `objectName` オプションで名前を付けておく（`name` の重複も同時に消える）。
- マスターは 1 枚固定、レイアウト 1 番は `DEFAULT`（背景参照つき、`slideNumber` 使用時は番号 placeholder つき）が必ず入る。レイアウト対応表の番号は「ユーザー定義順 + 2」になる。
- `masterName` 付きのスライドは、未使用 placeholder が空枠として必ず入る。PowerPoint では「クリックしてタイトルを入力」の枠が見える。要らないなら後処理で消すか、レイアウトを placeholder なしで作る。
- PowerPoint の「修復」に掛かる項目（重複 `cNvPr id`、実在しないパートへの Override、rels の Target 解決、`[Content_Types].xml` の Default 抜け）は python-pptx では検出できない。自前の OPC 整合チェッカが要る（この環境に LibreOffice は無い）。
- 画像プレースホルダー（`type: 'pic'`）は実装されていない。画像は座標指定の `addImage` になる。
- `[Content_Types].xml` を PptxGenJS 生成のまま使うと、実在しない `slideMasterN.xml` の Override が混ざる。マスター差し替え時はここを作り直すほうがよい。
- `write()` の `outputType` 指定時は無圧縮（STORE）で出る。再圧縮は後処理側で `DEFLATE` を指定する。

### 参照した場所

- `<pg>/package.json`（dependencies: jszip ^3.10.1）、`<pg>/node_modules/jszip/package.json`（3.10.2）
- `<pg>/dist/pptxgen.cjs.js` L4（JSZip require）、L7003-7086（exportPresentation: パート配置 L7030-7060、generateAsync L7074/L7078）
- 同 L579-602（MASTER_OBJECTS / SLIDE_OBJECT_TYPES / PLACEHOLDER_TYPES）、L1585-1607（objects の解釈、placeholder idx = 100+idx）、L2630-2642（addPlaceholdersToSlideLayouts）、L5121-5159（placeholder からの位置継承）、L6306-6318（genXmlPlaceholder）
- 同 L6062-6081（bodyPr の fit）、L5963（run の latin/ea/cs）、L3935（lang 既定 en-US）
- 同 L6350-6354（Content_Types の slideMaster Override がスライド数ぶん出る）
- 同 L5184（表の cNvPr id）、L5403 / L5531 / L5642（テキスト・画像・図形の cNvPr id = idx+2）、L5105-5108（DEFAULT レイアウトの bg）、L6941-6946（setSlideNumber）、L7244-7258（addSlide の臨時レイアウト）
- 同 L6507-6514（getNotesFromSlide の CRLF 置換）、L6527-6528（notesSlide の単一 `<a:t>`）、L5233-5290（表のセル結合）、L5970-5982（hyperlink.slide と hlinkClr）、L633-650（getSmartParseNumber）、L686-693（inch2Emu）
- レビュアー試作: スクラッチパッド `review/opccheck.mjs`（OPC 整合チェック）、`review/roundtrip.mjs`
- `<pg>/types/index.d.ts` L629（PLACEHOLDER_TYPE）、L1271-1278（PlaceholderProps）、L1816（fit）、L1836（indentLevel）、L2443-2450（WriteProps）、L593（WRITE_OUTPUT_TYPE）、L2490-2518（SlideMasterProps）、L2650（addNotes）
- 試作: スクラッチパッド `research/proto1.mjs`、出力 `proto1-ppt_slides_slide1.xml` `proto1-ppt_slides_slide2.xml` `proto1-_Content_Types_.xml` ほか

---

## 2. マスターの XML 構造

### 分かったこと

**参照の向き**（PptxGenJS 出力で確認。OOXML 一般と同じ）
- `ppt/presentation.xml`: `<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/>` と `<p:sldIdLst><p:sldId id="256.." r:id="rId2.."/>`、`<p:notesMasterIdLst>`、`<p:sldSz cx cy/>`、`<p:defaultTextStyle>`。
- `ppt/_rels/presentation.xml.rels`: rId1 → `slideMasters/slideMaster1.xml`、rId2.. → `slides/slideN.xml`、続けて notesMaster / presProps / viewProps / theme / tableStyles。
- `ppt/slideMasters/slideMaster1.xml`: `<p:cSld>`（背景・共通図形・スライド番号 placeholder）、`<p:clrMap>`、`<p:sldLayoutIdLst><p:sldLayoutId id="2147483649.." r:id="rIdN"/>`、`<p:hf>`、`<p:txStyles>`（titleStyle / bodyStyle / otherStyle、フォントは `+mj-lt` `+mj-ea` `+mn-lt` `+mn-ea` でテーマ参照）。
- `ppt/slideMasters/_rels/slideMaster1.xml.rels`: 各 `../slideLayouts/slideLayoutN.xml` と `../theme/theme1.xml`。
- `ppt/slideLayouts/slideLayoutN.xml`: `<p:cSld name="<masterName>">` の中に placeholder 図形。`<p:clrMapOvr><a:masterClrMapping/>`。rels は `../slideMasters/slideMaster1.xml` のみ。
- `ppt/slides/slideN.xml` の rels: `slideLayout`（`../slideLayouts/slideLayoutK.xml`）と `notesSlide`、加えて画像・ハイパーリンクの関係。スライドは「レイアウトだけ」を指し、マスターやテーマを直接は指さない。`rId` は固定ではない（試作の slide1 は rId1=hyperlink、rId2=slideLayout、rId3=notesSlide。画像やリンクの数で前にずれる）。書き換えるときは `Type` 属性で引く。
- `ppt/theme/theme1.xml`: `<a:clrScheme>` `<a:fontScheme>`（majorFont / minorFont。`<a:latin>` に加えて `<a:font script="Jpan" typeface=...>` の script 別指定）`<a:fmtScheme>`。マスターの rels から参照される。
- `[Content_Types].xml`: マスター・レイアウト・スライド・ノート・テーマそれぞれに `<Override PartName=... ContentType=...>` が要る。

**既存 .pptx のマスターを差し替える見立て**（手順の見立て。未試行）
1. テンプレート .pptx から `ppt/slideMasters/slideMasterM.xml`(+rels)、それが指す `ppt/slideLayouts/*.xml`(+rels)、`ppt/theme/themeT.xml`、マスターやレイアウトが参照する `ppt/media/*` を取り出す。
2. PptxGenJS 出力から `slideMaster1.xml`(+rels)、`slideLayout*.xml`(+rels)、`theme1.xml` を捨てて、1 のファイルを同じ名前体系で置く（`slideLayoutN.xml` の番号が変わるなら master rels の Target も書き換える）。
3. `ppt/slides/_rels/slideN.xml.rels` の `Type=".../slideLayout"` の Relationship を `Type` で探し、Target を対応表に従って新しいレイアウト番号に付け替える（`rId` 決め打ちにしない）。
4. `ppt/presentation.xml` の `<p:sldMasterIdLst>` は 1 枚なら変更不要（rId1 → slideMaster1）。マスターを複数持ち込むなら `sldMasterId` を増やし、`presentation.xml.rels` にも足す。
5. `[Content_Types].xml` の Override をレイアウト数・メディア拡張子に合わせて作り直す。
6. スライド側の placeholder `<p:ph idx type>` を、新レイアウトの placeholder の `idx`/`type` に合わせる（PptxGenJS の `idx=100+n` は独自採番なので、テンプレート側の idx と一致させないと継承が切れる）。
7. スライド番号 placeholder（PptxGenJS は `idx="4294967295"`）とノートマスターは PptxGenJS 生成のまま残してよいと思われる（推測）。

### 設計に効く制約

- スライドがレイアウトを指すのは `slideN.xml.rels` の 1 行だけ。マスター差し替えの「継ぎ目」は「レイアウト名 → レイアウトファイル番号」と「placeholder 名 → `<p:ph idx type>`」の 2 つの対応表に集約できる。
- PptxGenJS の placeholder `idx` は 100 始まりの独自採番。既存テンプレートに合わせるには後処理で `idx` を書き換える必要がある。
- テーマのフォントは `+mj-*` `+mn-*` 経由でマスターの `txStyles` から参照される。フォント統一を「テーマで」やるか「run ごとに `<a:latin>/<a:ea>`」でやるかで、差し替え時の壊れ方が変わる（run 指定は差し替え後も残る）。
- `[Content_Types].xml` と各 rels の整合は手で取る必要がある。検証手段は python-pptx（`uv run --with python-pptx`）で読めることまで。PowerPoint 実機での開封確認は別途要る。

### 参照した場所

- `<pg>/dist/pptxgen.cjs.js` L6470-6489（presentation.xml.rels）、L6547-6560（slideMaster1.xml）、L6596-6622（layout rels / slide rels）、L6636-6660（master rels）、L6680-6684（theme1.xml）、L6697-6720（presentation.xml）、L6340-6372（Content_Types）
- 試作出力: `proto1-ppt_presentation.xml`、`proto1-ppt__rels_presentation.xml.rels`、`proto1-ppt_slideMasters__rels_slideMaster1.xml.rels`、`proto1-ppt_slideLayouts__rels_slideLayout1.xml.rels`、`proto1-ppt_slides__rels_slide1.xml.rels`、`proto1-_Content_Types_.xml`、`proto1-ppt_theme_theme1.xml`

---

## 3. Slidev 52.19.1 の書き出し処理

### 分かったこと

**公開 API と内部**
- `@slidev/cli` の公開エクスポート（`dist/index.d.mts`）は `createServer(options, viteConfig?, serverOptions?)`、`resolveOptions(entryOptions, mode)`、`createDataUtils`、`ViteSlidevPlugin`、`parser`（= `@slidev/parser/fs`）の 5 つだけ。
- `exportSlides` `exportNotes` `getExportOptions` は `dist/export-B_maVZXE.mjs` にあり、`index.mjs` からは出ていない。package.json の `exports` に `"./*": "./*"` があるので `@slidev/cli/dist/export-B_maVZXE.mjs` を直接 import すれば読めるが、ファイル名にビルドハッシュが付く内部物で、版が上がると名前が変わる。
- `slidev export` コマンドの流れ（`cli.mjs` L372-409）: `resolveOptions({entry, theme}, 'export')` → `createServer(options, {server:{port}})` → `server.listen()` → `exportSlides({ port, ...getExportOptions(args, options) })` → `server.close()`。この 3 つのうち `resolveOptions` と `createServer` は公開。
- `exportSlides` の中の再利用したい部分は `go(no, clicks)`（L131-187）: URL を `http://localhost:{port}{base}{no}?print=true|clicks&range=...&clicks=...` に組み（一括モードでは `no='print'` なので `/print?print=...`）、`page.goto` → `emulateMedia({colorScheme, media:'screen'})` → `[data-slidev-no="{no}"]` または `body` を待ち → `.slidev-slide-loading` の消滅、`[data-waitfor]`、iframe、mermaid、monaco の待機。これは関数内のクロージャで export されていないので、同じ手順を自前で書く。
- `getExportOptions` の既定値: `scale`（deviceScaleFactor）は 2、`withClicks` は `format === 'pptx'` のとき true、`waitUntil` は `'networkidle'`、`height` は `Math.round(canvasWidth / aspectRatio)`。

**サーバのモードで有効になる機能が変わる**
- `resolveOptions(entry, mode)` の `mode` で Vite の define が変わる（`serve-*.mjs` L1104-1120）。`__DEV__` は `mode === 'dev'` のときだけ true、`__SLIDEV_FEATURE_PRINT__` は `mode === 'export'`（または `build` で `download` 有効）のときだけ true。
- `/print` と `/presenter/print` のルートは `__SLIDEV_FEATURE_PRINT__` のときだけ登録される（`client/setup/routes.ts` L61-75）。`window.__slidev__` は `__DEV__` のときだけ生える（`setup/root.ts` L31-35）。**同じサーバで両方は使えない。** レビュアー実測: export モードでは `/print?print=true` に `.print-slide-container` が 12 個並ぶが `window.__slidev__` は undefined。dev モードでは `/print` が `play` に落ちてコンテナ 0 個。
- Playwright の読み込み `importPlaywright()` は `playwright-chromium` をユーザールート → ワークスペース → グローバル → 通常解決の順で探す。アドオンからは `import('playwright-chromium')` で足りる。
- PPTX 書き出し `genPagePptx`（L327-351）は `?print=clicks` の 1 ページ（`no='print'`）でスクリーンショットを撮り、`slide.background = { data: png }` として貼るだけ。PptxGenJS の `defineLayout({ width: canvasWidth/96, height: /96 })`、`pptx.layout`、`author/company/title/subject`、`addNotes(note)` を使う。

**`slides.md` の解析結果の取り方**
- `resolveOptions(...)` の戻り `options.data: SlidevData` に `slides: SlideInfo[]`（`index` `frontmatter` `content` `note` `title` `level` `source` `importChain` `noteHTML?`）、`config: SlidevConfig`（`canvasWidth` `aspectRatio` `colorSchema` `routerMode` `export` など）、`headmatter`、`markdownFiles` が入る。`SlideInfo` にレイアウト名は直接ない。レイアウトは `frontmatter.layout ?? slides[0].frontmatter.defaults.layout ?? (index===0 ? 'cover' : 'default')` で決まり、その名前が `utils.getLayouts()`（テーマ・アドオン・ユーザーの `layouts/` を集めたもの）に無ければ警告して `default` に落ちる（`serve-*.mjs` L683-690）。
- `@slidev/parser/fs` は `load({roots, userRoot}, filepath)`（config 抜きの `SlidevData`）のほか、`save`、`parse`、`parseSync`、`stringify`、`parseRangeString`、`resolveConfig` など core の関数も再エクスポートする。`@slidev/parser/core` の `parse(markdown, filepath)` は 1 ファイルぶんの `SlidevMarkdown`（`slides: SourceSlideInfo[]`）。
- 開発サーバのエンドポイント: `GET /__slidev/slides/{no}.json` が `withRenderedNote(data.slides[idx])`（`SlideInfo` + `noteHTML`）を返す。`POST` は編集用。他の `/__slidev/` 系はない。
- クライアント側: `#slidev/slides` 仮想モジュールが `slides: ShallowRef<SlideRoute[]>` を持ち、各 `route.meta` に `layout`（frontmatter.layout そのまま。既定は undefined）、`slide`（`SlideInfo` 相当 + `filepath` `start` `sourceIndex` `id` `no`）、`__clicksContext` がある。dev モード（`mode: 'dev'`）に限り `window.__slidev__ = { nav, configs, themeConfigs }` が生え、`nav.slides` `nav.clicksTotal` `nav.currentSlideNo` `nav.go()` が取れる。export モードでは生えない。
- クリック数: `ClicksContext.total`（`route.meta.__clicksContext.total` または `useNav().clicksTotal`）。描画後にしか確定しない。`/print?print=clicks`（export モード）では、クリック状態ごとに `.print-slide-container` が並び、`id` は `{no:3桁}-{clicks+1:2桁}`（例 `003-02`）。スライドごとのコンテナ数 − 1 がクリック数。`exportSlides` の `getSlidesIndex()`（L188-201）はこの id を数えたうえで累積和を取っており、PDF の目次用の「ページ番号」を出す関数で、クリック数そのものではない。
- スライドの `frontmatter.lang` は `SlideWrapper.vue` L46 で `[data-slidev-no]` の `lang` 属性に出る。PptxGenJS の run `lang` 既定 `en-US` と揃えるなら、ここから拾う。

**DOM から位置を測る入口**
- `/print?print=true` / `/print?print=clicks`（export モードのみ）の表示: `#print-container > #print-content > .print-slide-container`（`pages/print.vue` → `PrintContainer.vue` → `PrintSlideClick.vue`）。各コンテナは `style="width:{slideWidth}px;height:{slideHeight}px"` で、scoped style は `relative overflow-hidden break-after-page translate-0 bg-main`。`translate-0` により transform は単位行列（拡縮なし）なので、子要素の `getBoundingClientRect()` からコンテナの rect を引けば Slidev キャンバス px がそのまま得られる（viewport 幅が canvasWidth 以上のとき。`exportSlides` は viewport を `width=canvasWidth, height=canvasHeight×枚数` にしている）。ただし transform がある以上コンテナは containing block になり、`position: fixed` の部品はコンテナ基準になる。
- `/{no}?print=true` は `pages/play.vue` → `SlideContainer.vue` で描画され、`.print-slide-container` は無く、`#slide-content` に `transform: scale(var(--slidev-slide-scale))` が掛かる（`isPrintMode` は true になり print 用スタイルは入る）。dev モードでも開けるが、計測には `/print` のほうが素直。
- `.print-slide-container` の中は `[data-slidev-no="{no}"]`（`SlideWrapper.vue`、class `slidev-page`、`position:absolute; inset:0`、`lang` 属性つき）→ `<InjectedLayout>` = `.slidev-layout.{layoutName}`（`default`/`cover`/`center`/`two-columns` の `.col-left`/`.col-right` など）→ 本文。`frontmatter.zoom` があると `.slidev-page` に `scale:` が掛かる。
- キャンバス高さの計算が 2 か所で違う。client の `env.ts` L14 は `Math.ceil(canvasWidth / aspectRatio)`（980 × 9/16 → 552）、`getExportOptions` は `Math.round(...)`（→ 551）。`.print-slide-container` の高さは client 側の 552 px、`exportSlides` の viewport と PPTX の `defineLayout` は 551 px。座標換算の基準をどちらにするか決める必要がある。
- `/export` ルート（ブラウザ書き出し画面、`__SLIDEV_FEATURE_BROWSER_EXPORTER__`）は `isPrintMode` を true にするが、`#slide-container/#slide-content` を `transform: scale(...)` で縮めて表示し、画像は `getDisplayMedia` の画面キャプチャで作る。計測には向かない。

### 設計に効く制約

- 使ってよい公開 API は `resolveOptions` `createServer` と `@slidev/parser`。Playwright でページを開く `go()` 相当は自前で持つ（`exportSlides` は流用できない）。`dist/export-*.mjs` の直 import は版更新で壊れる。
- サーバは `resolveOptions(entry, 'export')` で立てる。そうしないと `/print` ルートが無い。その代わり `window.__slidev__` は使えないので、クリック数もレイアウトも「`/print?print=clicks` の DOM（コンテナ id、`.slidev-layout` の class）」と「`options.data`（Node 側）」から取る前提にする。
- レイアウト名の決定規則（`frontmatter.layout` → `defaults.layout` → 1 枚目 `cover` / それ以外 `default` → `getLayouts()` に無ければ `default`）を書き出し側でも同じ順で解決する必要がある。DOM の `.slidev-layout` の class から読むほうが確実。
- クリック数は描画後にしか決まらない。`/print?print=clicks` の `.print-slide-container` をスライドごとに数える（Slidev 標準の PPTX 書き出しと同じ材料）。
- 位置の実測は `/print?print=true|clicks` で、`.print-slide-container` を基準に `getBoundingClientRect()` の差分で取ると、キャンバス px が 1:1 で得られる。`zoom` を使うスライドは `.slidev-page` の `scale` を戻す必要がある。`position: fixed` の部品はコンテナ基準になる。
- キャンバス高さは client（ceil）と export（round）で 1 px ずれうる。PPTX のスライドサイズと座標換算の分母をどちらに合わせるか決める。
- `note` は生 Markdown、`noteHTML` は `/__slidev/slides/N.json` からのみ（dev / export どちらのモードでも middleware は入る）。

### 参照した場所

- `<cli>/dist/index.d.mts`（公開 API）、`<cli>/dist/index.mjs`、`<cli>/package.json` exports
- `<cli>/dist/cli.mjs` L372-409（export コマンド）
- `<cli>/dist/export-B_maVZXE.mjs` L106-130（exportSlides 入口）、L131-187（go）、L188-201（getSlidesIndex）、L271-291（genPagePngOnePiece）、L327-351（genPagePptx）、L372-406（getExportOptions）、L407-424（importPlaywright）
- `<cli>/dist/serve-DzftUF4M.mjs` L121（`/__slidev/slides/N.json` の正規表現）、L1545-1552（GET 応答）、L683-690（レイアウト決定と default への退避）、L1104-1120（`getDefine`: `__DEV__` `__SLIDEV_FEATURE_PRINT__`）、L1456-1488（`/@slidev/slides` 仮想モジュール）、L1725-1740（`meta` の中身）
- `<client>/setup/root.ts` L31-35（`window.__slidev__`）、`<client>/setup/routes.ts` L61-75（`/print` ルートの条件）
- `<client>/internals/PrintContainer.vue`、`PrintSlide.vue`、`PrintSlideClick.vue`（id の形式、コンテナの style、`translate-0`）、`SlideWrapper.vue` L43（`data-slidev-no`）L46（`lang`）L59-71（`.slidev-page`）、`SlideContainer.vue` L137（transform）、`pages/play.vue` L91（SlideContainer）
- レビュアー試作: スクラッチパッド `review/measure.mjs` `review/measure2.mjs`（export / dev 各モードでの `/print` と `window.__slidev__` の実測）
- `<client>/composables/useNav.ts` L286-287（`isPrintMode` `isPrintWithClicks`）L296（`printRange`）、`useClicks.ts` L166-180（createFixedClicks）、`useSlideInfo.ts` L20（fetch URL）
- `<client>/env.ts` L9-14（`slideWidth` `slideHeight`、`Math.ceil`）、`<cli>/dist/export-B_maVZXE.mjs` L398（`Math.round`）、`<client>/layouts/{default,center,two-cols,cover,image-right}.vue`
- `<client>/pages/export.vue` L23-30, L60-80、`<client>/logic/screenshot.ts` L1-40
- `<types>/dist/index.d.mts` L89-103（ClicksContext）、L161（BuiltinLayouts）、L164-178（SlideInfoBase）、L206-220（SlideInfo）、L266-281（SlidevData）、L290-303（SlideRoute）、L965-976（ResolvedSlidevOptions）
- `<parser>/dist/fs.d.mts`（`load`）、`<parser>/dist/core.d.mts`（`parse` `parseSync`）

---

## 4. 游ゴシックの扱い

### 分かったこと

**PptxGenJS が出す XML**
- run に `fontFace: X` を付けると `<a:latin typeface="X" pitchFamily="34" charset="0"/><a:ea typeface="X" pitchFamily="34" charset="-122"/><a:cs typeface="X" pitchFamily="34" charset="-120"/>` の 3 つが同じ名前で出る。`'游ゴシック'` でも `'Yu Gothic'` でも書き分けはなく、指定した文字列がそのまま入る。
- 表セルの `endParaRPr` は `charset="0"` の 3 つ。スライド番号は `<a:latin>/<a:ea>/<a:cs>` を charset なしで出す。
- `pptx.theme = { headFontFace, bodyFontFace }` は theme1.xml の majorFont / minorFont の `<a:latin typeface>` だけを差し替える。`<a:ea typeface=""/>` は空のまま。`<a:font script="Jpan" typeface="游ゴシック Light"/>`（major）と `"游ゴシック"`（minor）は **PptxGenJS 4.0.1 のテンプレートに日本語名で固定**されている（Excel 用テーマは `Yu Gothic Light`/`Yu Gothic`）。
- マスターの `txStyles` は `+mj-lt/+mj-ea/+mn-lt/+mn-ea` でテーマ参照。run に `fontFace` を付けなければテーマの Jpan 指定（游ゴシック）が効く。
- run の `lang` 既定は `en-US`。`lang: 'ja-JP'` を run ごとに渡さない限り日本語文字にも `en-US` が付く。

**フォント名と OS**
- Windows（8.1 以降）: ファミリー名 `Yu Gothic`（日本語名 `游ゴシック`）。ウェイトは Light / Regular / Medium(10 以降) / Bold。PowerPoint の一覧では `游ゴシック` `游ゴシック Light` `游ゴシック Medium` が別項目として見える。
- `Yu Gothic UI` は Windows 10 で追加された別ファミリー（Light / Semilight / Regular / Semibold / Bold）。UI 用に仮名を詰めており、英数字は Segoe UI 由来。`Yu Gothic` の代替にはならない。
- macOS: OS 同梱は `YuGothic`（日本語名 `游ゴシック体`）の Medium / Bold のみ（Sierra 以降はダウンロード扱い）。Regular/Light はない。Windows の `Yu Gothic`/`游ゴシック` とは別名。
- Microsoft 365 の「クラウドフォント」に `Yu Gothic` Light/Regular/Medium/Bold と `Yu Gothic UI` 各種が含まれ、Mac 版 Office でも取得される。したがって M365 の PowerPoint for Mac なら `Yu Gothic` 名で解決しうる（クラウドフォントの取得はサインインと初回ダウンロードが前提。推測: 未取得の環境では次項の代替になる）。
- 見つからないときの PowerPoint の挙動: ファイル内のフォント名は保持したまま、表示時に代替する。Windows は `FontSubstitutes` レジストリ → PANOSE/charset/pitchFamily の情報 → テーマ既定フォント（Calibri など）の順（外部ブログの説明。Microsoft の一次資料は見つけていない）。「フォントの置換」ダイアログはファイル内で使われているフォントだけを列挙し、2 バイトフォントを 1 バイトフォントへは置換できない。
- OOXML の `typeface` は名前文字列の一致で解決される。Windows の Yu Gothic は name テーブルに英語名と日本語名の両方を持つので `Yu Gothic` `游ゴシック` どちらでも解決する（推測。フォントの name テーブルを直接は確認していない。`Yu Gothic Light` / `Yu Gothic Medium` は別ファミリー名として登録されている）。

### 設計に効く制約

- PptxGenJS は `<a:ea>` だけを別名にできない。「英数字は別フォント、和文は游ゴシック」のような書き分けは後処理で run の `<a:latin>` を書き換えるしかない。
- 「游ゴシックに統一」の実現方法は 2 通りで、どちらを既定にするか決める必要がある。
  - (a) run ごとに `fontFace` を付ける: 差し替えたマスターのテーマに関係なく残る。マスター差し替え後もフォントが固定される。
  - (b) run には付けず、テーマの Jpan（`游ゴシック` 固定）とマスターの `txStyles` に任せる: マスター差し替えでフォントも切り替わる。PptxGenJS の theme は `<a:latin>` しか変えられないので、和文の名前を変えたければ theme1.xml の後処理になる。
- フォント名は `'游ゴシック'`（日本語名）か `'Yu Gothic'`（英語名）のどちらかに揃える。PptxGenJS のテーマ側は `游ゴシック` 固定なので、合わせるなら日本語名。Mac の OS 同梱名（`YuGothic`/`游ゴシック体`）とは一致しない。
- ウェイトは Windows でも `游ゴシック Light`/`Medium` が別ファミリー名で、Mac 同梱は Medium/Bold のみ。`Regular` を前提にすると Mac で太る。
- run の `lang` を `ja-JP` にするかは別途決める（校正・ハイフネーションに影響。表示には影響しない）。Slidev 側の `frontmatter.lang` が DOM の `lang` 属性に出るので、揃えるならそれを使う。

### 参照した場所

- `<pg>/dist/pptxgen.cjs.js` L5963（run のフォント 3 要素）、L5701（スライド番号）、L6680-6684（theme の major/minor と Jpan 固定値）、L2979（Excel 用テーマの `Yu Gothic`）、L3935 ほか（`lang` 既定）
- `<pg>/types/index.d.ts` L1290-1303（ThemeProps）、L1210（lang）
- 試作出力 `proto1-ppt_slides_slide1.xml` `proto1-ppt_slides_slide2.xml` `proto1-ppt_theme_theme1.xml`
- https://learn.microsoft.com/en-us/typography/font-list/yu-gothic （ウェイト、Windows 8.1/10/11、Yu Gothic UI の追加）
- https://ja.wikipedia.org/wiki/游書体 （Windows/Mac の名前とウェイト、Yu Gothic UI の違い）
- https://support.apple.com/en-us/103197 （macOS Ventura 同梱: `YuGothic Bold` `YuGothic Medium`）
- https://support.microsoft.com/en-us/office/cloud-fonts-in-office-f7b009fe-037f-45ed-a556-b5fe6ede6adb （クラウドフォントに Yu Gothic / Yu Gothic UI）
- https://chiilabo.com/2024-01/windows-mac-yu-gochic-difference/ （Windows `游ゴシック`/`Yu Gothic`、Mac `游ゴシック体`/`YuGothic`）
- https://consul-deck-lab.com/posts/ppt-font-yu-gothic-meiryo/ （PowerPoint 上のウェイト別表示、Yu Gothic UI の位置づけ）
- https://neuxpower.com/blog/2021/6/3/fonts-not-displaying-properly-in-powerpoint 、https://wisechecker.com/powerpoint-font-substitution-fallback/ 、https://support.microsoft.com/en-us/topic/replace-font-dialog-only-shows-fonts-within-the-powerpoint-presentation-ac661d45-d145-4da2-bae4-f7935ee42f04 （代替の挙動、置換ダイアログの制限）

---

## 設計への提案（事実とは分けて書く）

- 後処理の層を「jszip で開いた ZIP に対する、パス指定の XML 変換関数の列」として切る。PptxGenJS で出せるものはそのまま、グループ・アニメーション・画面切り替え・`normAutofit fontScale`・placeholder `idx` の付け替え・`[Content_Types].xml` の作り直しをここに寄せると、マスター差し替えも同じ層の 1 変換になる。**必須の変換として「スライド内の `cNvPr id` / `name` の重複是正」を最初に入れる**（表を 1 つでも使えば重複する）。
  - メリット: PptxGenJS の版に依存する箇所と OOXML 直書きの箇所が分かれる。
  - デメリット: XML を文字列で扱う限り壊しやすい。python-pptx で読める、までしか自動検証できないので、OPC 整合チェッカ（Override とパートの突き合わせ、rels Target の解決、id 重複）を自前で持つ必要がある。
- 図形の特定は `objectName` を必ず付ける方針にする（PptxGenJS の採番 `Text 0` に頼らない）。
- フォントは、まず (a) run ごとに `fontFace: '游ゴシック'` を付ける方を既定にし、マスター差し替えの継ぎ目ができた時点で (b) テーマ任せに切り替えられるよう、フォント名の付与を 1 箇所に集める。
  - メリット: 差し替え前でも確実に游ゴシックになる。
  - デメリット: 差し替え後にテンプレートのフォントへ自動では寄らない。`<a:latin>` にも同じ名前が書かれるので英数字も游ゴシックの Latin グリフになり、「英数字は別フォント」にしたければ後処理で `<a:latin>` を書き換える必要がある。
- Playwright の待機手順（`go()` 相当）は `exportSlides` の L147-186 を写して自前に持つ。サーバは `resolveOptions(entry, 'export')` + `createServer` で立て、開く URL は `/print?print=true` または `/print?print=clicks`。`window.__slidev__` には頼らない前提にし、クリック数は `.print-slide-container` の id から、レイアウト名は `.slidev-layout` の class または `options.data` から取る。Slidev 側の更新で追従が要るが、内部ファイルの直 import よりは安全。
