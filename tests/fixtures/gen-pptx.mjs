// 後処理と OPC 整合チェッカの入力になる PPTX を、PptxGenJS 4.0.1 だけで作る。
// 実装（packages/）には依存しない。
//
// 使い方: node tests/fixtures/gen-pptx.mjs [出力パス]
// 既定の出力は tests/fixtures/pptx/sample.pptx。
//
// 末尾で、後処理が直す前提にしている 4 条件を assert する。PptxGenJS の版を上げて条件が
// 消えたら、後処理の検査が「何も直さなくても通る」ようになる前にここで気付く。
//   1. スライド内で <p:cNvPr id> が重複する（表を使うと別系列で採番される）
//   2. masterName 付きのスライドに、埋めていない placeholder が空の図形として入る
//   3. [Content_Types].xml に実在しない slideMasterN.xml の Override が出る（スライド数ぶん）
//   4. 1 つの <a:p> に <a:pPr> が 2 つ以上出る（run ごとに書かれる）
//
// 設計: wip/design/native-export.md §3.2（後処理の列）、§3.4（C6 / C2 / C14）

import { writeFileSync, mkdirSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { createRequire } from 'node:module'

const require = createRequire(import.meta.url)
const PptxGenJS = require('pptxgenjs')
const JSZip = require('jszip')

const here = dirname(fileURLToPath(import.meta.url))
const out = resolve(process.argv[2] ?? resolve(here, 'pptx/sample.pptx'))

// 1×1 の PNG（透明）
const PNG_1x1 =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='

export async function generate() {
  const pptx = new PptxGenJS()
  pptx.layout = 'LAYOUT_WIDE'

  // 設計 §5.2 の default と同じ形（座標はインチ）
  pptx.defineSlideMaster({
    title: 'default',
    background: { color: 'FFFFFF' },
    objects: [
      { placeholder: { options: { name: 'title', type: 'title', x: 0.76, y: 0.54, w: 11.81, h: 0.6 } } },
      { placeholder: { options: { name: 'body', type: 'body', x: 0.76, y: 1.31, w: 11.81, h: 5.66 } } },
    ],
  })

  // slide 1: title placeholder + 自由配置（太字とリンクを含む段落）+ 表 + 画像 + ノート
  const s1 = pptx.addSlide({ masterName: 'default' })
  s1.addText('見出し', { placeholder: 'title' })
  s1.addText(
    [
      { text: '太字', options: { bold: true, bullet: true, breakLine: false } },
      { text: 'と', options: {} },
      { text: 'リンク', options: { hyperlink: { url: 'https://sli.dev' }, breakLine: true } },
      { text: '2 段落目', options: { bullet: true } },
    ],
    { x: 1, y: 3, w: 6, h: 1.5, fontFace: '游ゴシック', lang: 'ja-JP', fit: 'shrink', objectName: 'free-text' },
  )
  s1.addTable(
    [
      [{ text: '名前' }, { text: '用途' }],
      [{ text: 'default' }, { text: '通常' }],
    ],
    { x: 1, y: 5, w: 6, colW: [3, 3], border: [{ type: 'none' }, { type: 'none' }, { type: 'solid', pt: 1, color: '999999' }, { type: 'none' }] },
  )
  s1.addImage({ data: PNG_1x1, x: 8, y: 3, w: 1, h: 1, altText: '1x1' })
  s1.addNotes('1 行目\n2 行目\n3 行目')

  // slide 2: 自由配置だけ（placeholder を埋めない → 空の title / body が自動追加される）
  const s2 = pptx.addSlide({ masterName: 'default' })
  s2.addText('自由配置だけ', { x: 1, y: 1, w: 4, h: 1 })
  s2.addShape(pptx.ShapeType.line, { x: 1, y: 3, w: 4, h: 0, line: { color: '000000', width: 1 } })

  const buf = await pptx.write({ outputType: 'nodebuffer' })
  return buf
}

export async function assertFixtureConditions(buf) {
  const zip = await JSZip.loadAsync(buf)
  const read = (p) => zip.file(p).async('string')
  const slide1 = await read('ppt/slides/slide1.xml')
  const slide2 = await read('ppt/slides/slide2.xml')
  const ct = await read('[Content_Types].xml')
  const problems = []

  // 1. cNvPr id の重複
  const ids = [...slide1.matchAll(/<p:cNvPr id="(\d+)"/g)].map((m) => m[1])
  if (new Set(ids).size === ids.length) problems.push('条件 1: slide1.xml の cNvPr id が重複していない')

  // 2. 空の placeholder
  // PptxGenJS は <p:ph の直後に改行とタブを入れるので、空白類で区切る
  const emptyPh = [...slide2.matchAll(/<p:sp>(?:(?!<\/p:sp>).)*<p:ph[\s/>](?:(?!<\/p:sp>).)*<\/p:sp>/gs)].filter((m) => !/<a:t>[^<]+<\/a:t>/.test(m[0]))
  if (emptyPh.length === 0) problems.push('条件 2: slide2.xml に空の placeholder が無い')

  // 3. 実在しない slideMaster の Override
  const masters = [...ct.matchAll(/PartName="\/ppt\/slideMasters\/slideMaster(\d+)\.xml"/g)].map((m) => `ppt/slideMasters/slideMaster${m[1]}.xml`)
  const missing = masters.filter((p) => !zip.file(p))
  if (missing.length === 0) problems.push('条件 3: 実在しない slideMaster の Override が無い')

  // 4. 1 つの <a:p> に <a:pPr> が 2 つ以上
  const paras = [...slide1.matchAll(/<a:p>(?:(?!<\/a:p>).)*<\/a:p>/gs)].map((m) => m[0])
  const multi = paras.filter((p) => (p.match(/<a:pPr/g) ?? []).length >= 2)
  if (multi.length === 0) problems.push('条件 4: <a:pPr> が 2 つ以上ある <a:p> が無い')

  return problems
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const buf = await generate()
  const problems = await assertFixtureConditions(buf)
  if (problems.length) {
    console.error('fixture の前提が崩れている:\n  ' + problems.join('\n  '))
    process.exit(1)
  }
  mkdirSync(dirname(out), { recursive: true })
  writeFileSync(out, buf)
  console.log(`wrote ${out} (${buf.length} bytes)`)
}
