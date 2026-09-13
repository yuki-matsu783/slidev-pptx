// 書き出しの全経路（native-export.md §1、§7、§8）
import { mkdtempSync, existsSync, readFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { beforeAll, describe, expect, it } from 'vitest'
import { exportPptx } from '../../packages/slidev-addon-pptx/src/export'
import { check } from '../../packages/slidev-addon-pptx/src/opc/check'
import type { Report } from '../../packages/slidev-addon-pptx/src/types'
import { openPptx, els } from '../helpers/pptx'
import type { OpenedPptx } from '../helpers/pptx'
import { DARK, PLAIN } from './helpers'
import { S, SLIDE_COUNT } from './slides'

const tmp = () => mkdtempSync(join(tmpdir(), 'slidev-pptx-'))

describe('exportPptx: plain.md', () => {
  let report: Report
  let output: string
  let p: OpenedPptx

  beforeAll(async () => {
    output = join(tmp(), 'plain.pptx')
    report = await exportPptx({ entry: PLAIN, output, lang: 'ja-JP' })
    p = await openPptx(readFileSync(output))
  }, 180_000)

  it('書き出せて、OPC 検査の error が 0、report.check は同じ結果', async () => {
    expect(existsSync(output)).toBe(true)
    const results = await check(readFileSync(output))
    expect(results.filter((r) => r.level === 'error')).toEqual([])
    expect(report.check).toEqual(results)
  })

  it('Report の形（§8.1）', () => {
    expect(report).toMatchObject({ output, slides: SLIDE_COUNT })
    expect(typeof report.generatedAt).toBe('string')
    expect(report.slidev).toMatch(/^\d+\./)
    expect(report.pptxgenjs).toMatch(/^4\./)
    expect(report.native + report.replaced).toBeGreaterThan(0)
    expect(Array.isArray(report.replacements)).toBe(true)
    expect(Array.isArray(report.warnings)).toBe(true)
  })

  it('記録の JSON は <output>.report.json に、Report と同じ内容で出る', () => {
    const json = JSON.parse(readFileSync(output + '.report.json', 'utf8'))
    expect(json).toEqual(JSON.parse(JSON.stringify(report)))
  })

  it('画像への置き換えの一覧: math / mermaid / unknown-element / explicit / gradient', () => {
    const reasons = new Set(report.replacements.map((r) => r.reason))
    for (const r of ['math', 'mermaid', 'unknown-element', 'explicit', 'gradient']) expect(reasons.has(r as never), r).toBe(true)
    for (const r of report.replacements) {
      expect(r.width).toBeGreaterThan(0)
      expect(r.height).toBeGreaterThan(0)
      expect(r.selector).toMatch(/^[a-z]/)
      expect(r.name).toMatch(/^Replaced \d+$/)
    }
  })

  it('置き換え画像は 2 倍解像度（撮った px は枠の 2 倍）', async () => {
    const math = report.replacements.find((r) => r.reason === 'math')!
    const doc = await p.xml(`ppt/slides/slide${math.slide}.xml`)
    const pic = els(doc, 'p', 'pic').find((x) => els(x, 'p', 'cNvPr')[0].getAttribute('name') === math.name)!
    const cx = Number(els(pic, 'a', 'ext')[0].getAttribute('cx'))
    const px = cx / (12192000 / 980)
    expect(Math.round(math.width / px)).toBe(2)
  })

  it('警告コード（§8.2）: plain.md で出せるものが全部出る', () => {
    const codes = new Set(report.warnings.map((w) => w.code))
    for (const c of ['W-CSS', 'W-MATH-INLINE', 'W-LI-BLOCK', 'W-LAYOUT', 'W-OVERFLOW', 'W-OFFSLIDE', 'W-IMAGE', 'W-HIDDEN', 'W-TRANSITION', 'W-INLINE', 'W-LINK']) expect(codes.has(c), c).toBe(true)
    // UnoCSS が効かないまま測ったときの W-RENDER は出ていない（番人が効いている）
    expect(report.warnings.some((w) => w.code === 'W-RENDER' && /UnoCSS/.test(w.message))).toBe(false)
    expect(report.warnings.filter((w) => w.code === 'W-TRANSITION')).toHaveLength(1)
    expect(report.warnings.some((w) => w.code === 'W-IMAGE' && w.slide === S.boxes)).toBe(true)
    expect(report.warnings.some((w) => w.code === 'W-OVERFLOW' && w.slide === S.offslide)).toBe(true)
    expect(report.warnings.some((w) => w.code === 'W-OFFSLIDE' && w.slide === S.offslide)).toBe(true)
    // W-LAYOUT は section と image-right の両方で鳴る
    expect(report.warnings.some((w) => w.code === 'W-LAYOUT' && w.slide === S.section)).toBe(true)
    expect(report.warnings.some((w) => w.code === 'W-LAYOUT' && w.slide === S.imageRight)).toBe(true)
    for (const w of report.warnings) {
      expect(w.slide).toBeGreaterThanOrEqual(0)
      expect(typeof w.message).toBe('string')
    }
  })

  it('規則として捨てたものは警告ではなく dropped の件数', () => {
    // ネイティブに残るコードブロックは 2 つ（two-cols と影つきの箱の中）。件数の単位（ブロック / 行）は設計に無いので下限だけ
    expect(report.dropped['code-highlight']).toBeGreaterThanOrEqual(2)
    expect(report.dropped['blockquote-border']).toBe(1)
    expect(report.dropped['transition']).toBe(1)
    expect(report.dropped['background-crop']).toBe(1)
    expect(report.warnings.some((w) => /highlight|blockquote/i.test(w.message))).toBe(false)
  })

  it('zoom を使ったスライドが記録に出る', () => {
    expect(report.zoom[S.zoom]).toBeCloseTo(0.8, 5)
  })

  it('cover の背景画像: bg が blip で、Background dim が最初の図形', async () => {
    const doc = await p.xml(`ppt/slides/slide${S.coverBg}.xml`)
    expect(els(els(doc, 'p', 'bg')[0], 'a', 'blip')).toHaveLength(1)
    const first = els(els(doc, 'p', 'spTree')[0], 'p', 'cNvPr')[1]
    expect(first.getAttribute('name')).toBe('Background dim')
  })

  it('image-right の右半分は画像として入る（/bg.png をサーバから取得）', async () => {
    const doc = await p.xml(`ppt/slides/slide${S.imageRight}.xml`)
    expect(els(doc, 'p', 'pic').length).toBeGreaterThanOrEqual(1)
    expect(report.warnings.some((w) => w.code === 'W-IMAGE' && w.slide === S.imageRight)).toBe(false)
  })

  it('取得できない画像は灰色の矩形（p:sp、W-IMAGE）で、pic にはならない', async () => {
    const doc = await p.xml(`ppt/slides/slide${S.boxes}.xml`)
    // 置き換え画像（button / gradient / explicit）の 3 つだけが pic
    expect(els(doc, 'p', 'pic')).toHaveLength(3)
    const spTree = els(doc, 'p', 'spTree')[0]
    const imageNamed = Array.from(spTree.childNodes).filter((n): n is Element => n.nodeType === 1 && /^Image \d+$/.test(els(n as Element, 'p', 'cNvPr')[0]?.getAttribute('name') ?? ''))
    expect(imageNamed).toHaveLength(1)
    expect(imageNamed[0].localName).toBe('sp')
  })

  it('ノートが段落に割れて入る（本文の placeholder だけを見る。PptxGenJS はスライド番号のフィールドも別の <p:sp> に出す）', async () => {
    const doc = await p.xml(`ppt/notesSlides/notesSlide${S.cover}.xml`)
    const body = els(doc, 'p', 'sp').find((sp) => els(sp, 'p', 'ph')[0]?.getAttribute('type') === 'body')!
    const texts = els(body, 'a', 'p').map((x) => els(x, 'a', 't').map((t) => t.textContent).join('')).filter(Boolean)
    expect(texts).toEqual(['ノート 1 行目', 'ノート 2 行目', '• 箇条書き'])
  })

  it('--range は Slidev の range と同じ綴りで、枚数が絞られる。後処理は PPTX の番号で引く（図形名が付き直り、slideMap で元番号と対応する）', async () => {
    const out = join(tmp(), 'range.pptx')
    const r = await exportPptx({ entry: PLAIN, output: out, range: '1-2,4' })
    expect(r.slides).toBe(3)
    expect(r.slideMap).toEqual({ 1: 1, 2: 2, 3: 4 })
    const q = await openPptx(readFileSync(out))
    expect(q.list().filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f))).toHaveLength(3)
    // 3 枚目（元の 4 枚目）の図形名が PptxGenJS の既定（Text 0）のまま残っていない
    const doc = await q.xml('ppt/slides/slide3.xml')
    const names = els(doc, 'p', 'cNvPr').map((c) => c.getAttribute('name'))
    expect(names).not.toContain('Text 0')
    expect(names).toContain('Title')
  }, 180_000)

  it('check: false なら OPC 検査を飛ばし、report.check は空', async () => {
    const out = join(tmp(), 'nocheck.pptx')
    const r = await exportPptx({ entry: PLAIN, output: out, range: '1', check: false })
    expect(r.check).toEqual([])
  }, 180_000)
})

describe('exportPptx: dark.md', () => {
  it('colorSchema: dark のデッキは W-DARK が 1 回出て、測った色のまま出る', async () => {
    const out = join(tmp(), 'dark.pptx')
    const r = await exportPptx({ entry: DARK, output: out })
    expect(r.warnings.filter((w) => w.code === 'W-DARK')).toHaveLength(1)
    const q = await openPptx(readFileSync(out))
    const bg = els(els(await q.xml('ppt/slides/slide1.xml'), 'p', 'bg')[0], 'a', 'srgbClr')[0].getAttribute('val')!
    expect(bg).not.toBe('FFFFFF')
  }, 180_000)
})
