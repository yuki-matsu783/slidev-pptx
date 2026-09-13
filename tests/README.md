# 受入テスト（i0001-03）

設計 `wip/design/native-export.md` と `wip/design/ppt-components.md` を検査として固定したもの。
実装（`packages/slidev-addon-pptx/`）はフェーズ 4 で入る。それまでは import 先が無いので全部 red。

## 回し方

依存（`vitest` `@xmldom/xmldom`、直接依存の `jszip` `pptxgenjs`）はフェーズ 4 の子チケットで入る。
入ったあと:

```sh
pnpm vitest run -c tests/vitest.config.ts        # 単体（ブラウザ無し）
pnpm vitest run -c tests/vitest.e2e.config.ts    # e2e と CLI（playwright-chromium と Slidev サーバを使う）
```

`package.json` の `test` / `test:e2e` scripts も同じ 2 行にする（フェーズ 4）。

## 置き場

| 場所 | 何を固定するか | 設計の節 |
|---|---|---|
| `fixtures/gen-pptx.mjs` | PptxGenJS 4.0.1 だけで作る入力 PPTX。末尾で 4 条件を assert。`node tests/fixtures/gen-pptx.mjs` で再生成 | §3.2、§11 の 6・7 |
| `fixtures/pptx/sample.pptx` | 上の生成物（コミット済み。版を上げたら再生成） | |
| `fixtures/capture/basic.json` `.data.json` | 収集結果の見本と `options.data` 相当。対で使う | §2 |
| `fixtures/deck/plain.md` | 部品を使わない e2e 用デッキ（アドオン無しで起動する） | ppt-components §2.2 の 0〜15 |
| `fixtures/deck/components.md` | PPT 部品のデッキ。フェーズ 4 まで `describe.skip` | ppt-components §1 |
| `fixtures/deck/dark.md` | `colorSchema: dark`（`W-DARK`） | §1.2 |
| `opc/check.test.ts` | C1〜C14 を「通る / 落ちる」の対で | §3.4 |
| `patch/*.test.ts` | 後処理 1〜7 を 1 変換 1 ファイル。`pipeline.test.ts` は列の順、往復の同値、DEFLATE | §3.1、§3.2 |
| `build/*.test.ts` | 単位、名前、sanitize、マスター、フォント、ノート、レイアウト解決、Capture → XML | §4、§5、§6 |
| `e2e/collect.test.ts` | 収集器を `page.evaluate(collect)` で走らせる | ppt-components §2、§3 |
| `e2e/export.test.ts` | 全経路。OPC 検査、Report、警告コード、dropped | §1、§8 |
| `cli/export.test.ts` | exit code、フラグ、scripts の綴り | §7 |

## import する名前（設計 §9 のとおり。検査で新しい名前を決めない）

`src/types.ts`: `Capture` `Report` `CheckResult` と要素の型
`src/collect/index.ts`: `collect`（自己完結。`page.evaluate` にそのまま渡す）
`src/build/`: `build`（convert.ts）、`LAYOUTS` `defineMasters` `masterFor`（masters.ts）、`resolveLayout`（layout.ts）、
`SLIDE_W_EMU` `SLIDE_W_IN` `emu` `pt` `inch` `clampBox` `toDashType` `textMargin` `cellMargin`（units.ts）、
`fontFor` `themeFonts`（fonts.ts）、`assignNames`（names.ts）、`sanitizeXmlText`（sanitize.ts）、`notesText`（notes.ts）
`src/patch/`: `Patch` `PatchContext` `PATCHES` `postProcess` `xmlPatch`（index.ts）、`findRelByType`（zip.ts）、
`patches/<name>.ts` が同名の `Patch` を 1 つ export、`replaceMaster.ts` は `MasterSwap` の型も
`src/opc/check.ts`: `check`
`src/export.ts`: `exportPptx({ entry, output, range?, lang?, check?, report? })`

検査が前提にしている引数の形と、設計に無い細部（実装が別解を選ぶなら検査を直す。設計に書き戻すべきものは design-feedback へ）:
- `build(capture, data, { assets: Record<captureId, dataUrl>, lang, layouts: string[] })` → `{ pptx, ctx }`。`layouts` は `options.utils.getLayouts()` のキー一覧
- `clampBox(box, canvas)` → `{ box, shift, dropped }`。`shift` は寄せた量と縮めた量の最大値
- `assignNames(elements, confirmedRoles: Map<id, 'title' | 'body' | 'body2'>, hasBackgroundDim)` → `string[]`。`Placeholder N` の N は spTree の 1 始まりの位置
- `resolveLayout(slideIndex, data, layoutNames: string[])` → レイアウト名
- `emu(px, canvas)` `pt(px, canvas)` `inch(px, canvas)` `textMargin(inset, canvas)` `cellMargin(inset, canvas)` `toDashType(dash)`
- `defineMasters` は `LAYOUTS` の写しを PptxGenJS に渡す（`createSlideMaster` が options を破壊するため）
- `sanitizeXmlText` は XML 1.0 で合法な DEL / C1（0x7F–0x9F）も落とす
- `notesText` は字下げ後の `- ` も `• ` にし、CRLF を LF にし、`undefined` を空文字にする
- `findRelByType` は `slideLayout` のような短い名前でも完全な URI でも引ける
- `two-cols` の `body2` placeholder の `type` は `body`
- CLI: entry が無ければ exit 1（引数の綴りは合っているので 2 ではない）
- `W-CSS` の message には捨てた property 名（`shadow` など）を含める
- `collect` は 1 関数の中に閉じる（Playwright は関数を文字列化して渡すので、モジュール先頭の定数やヘルパを参照できない）
- `Capture.canvas.height` は `.print-slide-container` の rect のまま（551〜552。丸め規則は設計に無い）
- `dropped['code-highlight']` の単位（ブロックか行か）は設計に無いので、検査は下限だけ見る
- e2e の viewport は Slidev と同じ「980 × 552 × 枚数」。待機列は `native-export.md` §1.2 を写した（`tests/e2e/helpers.ts`）

## 設計へ書き戻す候補（design-feedback の子で出す）

- ppt-components.md §2.2: 「`p` の中身が `img`（と空白）だけなら、その `p` ごと画像要素（区切りブロック）」の特例。
  Markdown の `![]()` は `<p><img></p>` になり、今の規則では 11（`p`）が先に当たって 6（`img`）に届かない。
  このデッキは裸の `<img>` で回避している
- ppt-components.md §2.2 の 1: `hr` は高さ 0〜1 px なので、「rect が空」を `w === 0 || h === 0` で判定すると規則 10 に届かない

## フェーズ 4 で最初に確かめること

- vitest が `tests/**` を ESM として動かすとき `import.meta.url` が取れること（`__dirname` は使わず `here` に統一）
- `page.evaluate(collect)` の直列化（上の「1 関数の中に閉じる」）
- `plain.md` の枚数は `tests/e2e/slides.ts` の `SLIDE_COUNT`（12）。`@slidev/parser` は `---` で始まる行を無条件に区切りにするので、水平線は `***` で書いてある
