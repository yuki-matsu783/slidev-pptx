// 書き出しの全経路（native-export.md §1、§7、§8）
import { mkdtempSync, existsSync, readFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { exportPptx } from '../../packages/slidev-addon-pptx/src/export'
import { check } from '../../packages/slidev-addon-pptx/src/opc/check'
import type { Report } from '../../packages/slidev-addon-pptx/src/types'
import { openPptx, els } from '../helpers/pptx'
import { DARK, PLAIN } from './helpers'

const tmp = () => mkdtempSync(join(tmpdir(), 'slidev-pptx-'))

describe('exportPptx: plain.md', () => {
  let report: Report
  let output: string

  it('書き出せて、OPC 検査の error が 0', async () => {
    output = join(tmp(), 'plain.pptx')
    report = await exportPptx({ entry: PLAIN, output, lang: 'ja-JP' })
    expect(existsSync(output)).toBe(true)
    const results = await check(readFileSync(output))
    expect(results.filter((r) => r.level === 'error')).toEqual([])
    expect(report.check).toEqual(results)
  }, 180_000)

  it('Report の形（§8.1）', () => {
    expect(report).toMatchObject({ output, slides: 14 })
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

  it('画像への置き換えの一覧: math / mermaid / unknown-element / explicit', () => {
    const reasons = new Set(report.replacements.map((r) => r.reason))
    for (const r of ['math', 'mermaid', 'unknown-element', 'explicit']) expect(reasons.has(r as never), r).toBe(true)
    for (const r of report.replacements) {
      expect(r.width).toBeGreaterThan(0)
      expect(r.height).toBeGreaterThan(0)
      expect(r.selector).toMatch(/^[a-z]/)
      expect(r.name).toMatch(/^Replaced \d+$/)
    }
  })

  it('置き換え画像は 2 倍解像度（撮った px は枠の 2 倍）', async () => {
    const math = report.replacements.find((r) => r.reason === 'math')!
    const p = await openPptx(readFileSync(output))
    // 置き換え画像の枠（EMU）から px に戻して比べる
    const doc = await p.xml(`ppt/slides/slide${math.slide}.xml`)
    const pic = els(doc, 'p', 'pic').find((x) => els(x, 'p', 'cNvPr')[0].getAttribute('name') === math.name)!
    const cx = Number(els(pic, 'a', 'ext')[0].getAttribute('cx'))
    const px = cx / (12192000 / 980)
    expect(Math.round(math.width / px)).toBe(2)
  })

  it('警告コード（§8.2）: W-CSS / W-MATH-INLINE / W-LAYOUT / W-OFFSLIDE / W-HIDDEN / W-TRANSITION が出る', () => {
    const codes = new Set(report.warnings.map((w) => w.code))
    for (const c of ['W-CSS', 'W-MATH-INLINE', 'W-LAYOUT', 'W-OFFSLIDE', 'W-HIDDEN', 'W-TRANSITION']) expect(codes.has(c), c).toBe(true)
    expect(report.warnings.filter((w) => w.code === 'W-TRANSITION')).toHaveLength(1)
    for (const w of report.warnings) {
      expect(w.slide).toBeGreaterThanOrEqual(0)
      expect(typeof w.message).toBe('string')
    }
  })

  it('規則として捨てたものは警告ではなく dropped の件数', () => {
    expect(report.dropped['code-highlight']).toBeGreaterThanOrEqual(3)
    expect(report.dropped['blockquote-border']).toBe(1)
    expect(report.dropped['transition']).toBe(1)
    expect(report.dropped['background-crop']).toBe(1)
    expect(report.warnings.some((w) => /highlight|blockquote/i.test(w.message))).toBe(false)
  })

  it('zoom を使ったスライドが記録に出る', () => {
    expect(report.zoom[11]).toBeCloseTo(0.8, 5)
  })

  it('cover の背景画像: bg が blip で、Background dim が最初の図形', async () => {
    const p = await openPptx(readFileSync(output))
    const doc = await p.xml('ppt/slides/slide10.xml')
    expect(els(els(doc, 'p', 'bg')[0], 'a', 'blip')).toHaveLength(1)
    const first = els(els(doc, 'p', 'spTree')[0], 'p', 'cNvPr')[1]
    expect(first.getAttribute('name')).toBe('Background dim')
  })

  it('image-right の右半分は画像として入る（/bg.png をサーバから取得）', async () => {
    const p = await openPptx(readFileSync(output))
    const doc = await p.xml('ppt/slides/slide9.xml')
    expect(els(doc, 'p', 'pic').length).toBeGreaterThanOrEqual(1)
    expect(report.warnings.some((w) => w.code === 'W-IMAGE' && w.slide === 9)).toBe(false)
  })

  it('ノートが段落に割れて入る', async () => {
    const p = await openPptx(readFileSync(output))
    const doc = await p.xml('ppt/notesSlides/notesSlide1.xml')
    const texts = els(doc, 'a', 'p').map((x) => els(x, 'a', 't').map((t) => t.textContent).join('')).filter(Boolean)
    expect(texts).toEqual(['ノート 1 行目', 'ノート 2 行目', '• 箇条書き'])
  })

  it('--range は Slidev の range と同じ綴りで、枚数が絞られる', async () => {
    const out = join(tmp(), 'range.pptx')
    const r = await exportPptx({ entry: PLAIN, output: out, range: '1-2,4' })
    expect(r.slides).toBe(3)
    const p = await openPptx(readFileSync(out))
    expect(p.list().filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f))).toHaveLength(3)
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
    const p = await openPptx(readFileSync(out))
    const bg = els(els(await p.xml('ppt/slides/slide1.xml'), 'p', 'bg')[0], 'a', 'srgbClr')[0].getAttribute('val')!
    expect(bg).not.toBe('FFFFFF')
  }, 180_000)
})
