// PowerPoint の図形 187 種（PptShape の type）を並べたデッキ shapes.md。
// 収集: 全種が shape / line で出て、調整値・反転・矢じりが Capture に入る。書き出し: prstGeom と avLst と xfrm の反転。
import { mkdtempSync, readFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import type { Browser } from 'playwright-chromium'
import { collect } from '../../packages/slidev-addon-pptx/src/collect/index'
import { exportPptx } from '../../packages/slidev-addon-pptx/src/export'
import { check } from '../../packages/slidev-addon-pptx/src/opc/check'
import { PRESET_NAMES } from '../../packages/slidev-addon-pptx/src/shapes/geometry'
import type { Capture, Element, LineElement, ShapeElement } from '../../packages/slidev-addon-pptx/src/types'
import { els, openPptx, shapeNamed } from '../helpers/pptx'
import type { OpenedPptx } from '../helpers/pptx'
import { SHAPES, launch, openPrint, startSlidev } from './helpers'
import type { Running } from './helpers'

/** 格子 7 枚 + 見本 1 枚 */
const SLIDE_COUNT = 8
const SAMPLE = 8

describe('collect: 図形 187 種', () => {
  let server: Running
  let browser: Browser
  let capture: Capture

  beforeAll(async () => {
    server = await startSlidev(SHAPES)
    browser = await launch()
    const page = await openPrint(browser, server.base, { slides: SLIDE_COUNT })
    capture = await page.evaluate(collect)
  }, 120_000)
  afterAll(async () => {
    await browser?.close()
    await server?.close()
  })

  const all = (): Element[] => capture.slides.flatMap((s) => s.elements)
  const named = (name: string) => all().find((e) => e.name === name)
  const sample = (name: string) => capture.slides.find((s) => s.no === SAMPLE)!.elements.find((e) => e.name === name) as ShapeElement

  it(`スライドは ${SLIDE_COUNT} 枚`, () => {
    expect(capture.slides).toHaveLength(SLIDE_COUNT)
  })

  it('187 種すべてが出る: line は kind line、それ以外は kind shape で shape が type と同じ', () => {
    expect(PRESET_NAMES).toHaveLength(187)
    const missing: string[] = []
    for (const n of PRESET_NAMES) {
      const e = named(n)
      if (!e) {
        missing.push(n)
        continue
      }
      if (n === 'line') {
        expect(e.kind, n).toBe('line')
      } else {
        expect(e.kind, n).toBe('shape')
        expect((e as ShapeElement).shape, n).toBe(n)
        expect(e.boxSource, n).toBe('prop')
      }
    }
    expect(missing).toEqual([])
  })

  it('図形の中の文字は段落になり、SVG の中身を拾わない', () => {
    expect((named('rect') as ShapeElement).paragraphs?.[0].runs.map((r) => r.text).join('')).toBe('文字')
    expect((named('wedgeRectCallout') as ShapeElement).paragraphs?.[0].runs.map((r) => r.text).join('')).toBe('吹き出し')
    expect((named('text-triangle') as ShapeElement).paragraphs?.[0].runs.map((r) => r.text).join('')).toBe('三角の文字は下半分')
    // 文字の無い図形は段落を持たない（SVG の path や marker を中身として数えない）
    expect((named('actionButtonHelp') as ShapeElement).paragraphs).toBeUndefined()
    expect((named('conn-bent3') as ShapeElement).paragraphs).toBeUndefined()
  })

  it('図形の SVG で W-PPT-CONTENT / W-INLINE / W-NESTED-PPT / W-CSS / W-HIDDEN / W-RENDER が鳴らない', () => {
    const codes = capture.slides.flatMap((s) => s.warnings.map((w) => w.code))
    for (const c of ['W-PPT-CONTENT', 'W-INLINE', 'W-NESTED-PPT', 'W-CSS', 'W-HIDDEN', 'W-RENDER']) expect(codes.filter((x) => x === c), c).toEqual([])
  })

  it('adj は数値のまま Capture に入る', () => {
    expect(sample('adj-roundRect').adj).toEqual({ adj: 50000 })
    expect(sample('adj-rightArrow').adj).toEqual({ adj1: 80000, adj2: 25000 })
    expect(sample('arc-half').adj).toEqual({ adj1: 10800000, adj2: 0 })
    expect(sample('text-ellipse').adj).toBeUndefined()
  })

  it('flipH / flipV は true のときだけ入る', () => {
    expect(sample('flipH-rightArrow')).toMatchObject({ flipH: true })
    expect(sample('flipH-rightArrow').flipV).toBeUndefined()
    expect(sample('flipV-triangle')).toMatchObject({ flipV: true })
    expect(sample('flipV-triangle').flipH).toBeUndefined()
    expect(sample('flipHV-homePlate')).toMatchObject({ flipH: true, flipV: true })
    expect(sample('conn-curved3')).toMatchObject({ flipV: true })
  })

  it('line.head / tail は shape の arrow に入る（none や指定なしは付けない）。line の要素は今までどおり line に', () => {
    expect(sample('conn-bent3').arrow).toEqual({ head: 'oval', tail: 'triangle' })
    expect(sample('conn-curved3').arrow).toEqual({ tail: 'arrow' })
    expect(sample('conn-straight1').arrow).toEqual({ head: 'diamond', tail: 'arrow' })
    expect(sample('arc-half').arrow).toEqual({ tail: 'arrow' })
    expect(sample('brace').arrow).toBeUndefined()
    expect(sample('conn-bent3').frame.line).toMatchObject({ color: '#7c3aed', width: 2 })
    const line = named('line') as LineElement
    expect(line.line.head).toBeUndefined()
  })

  it('h なしの図形は box の h を実測する（文字の枠は枠全体）', () => {
    const e = sample('text-ellipse-auto')
    expect(e.boxSource).toBe('prop')
    expect(e.box.h).toBeGreaterThan(0)
  })
})

describe('exportPptx: 図形 187 種', () => {
  let p: OpenedPptx
  let buf: Buffer
  let slideXml: Document[]

  beforeAll(async () => {
    const output = join(mkdtempSync(join(tmpdir(), 'slidev-pptx-')), 'shapes.pptx')
    await exportPptx({ entry: SHAPES, output, lang: 'ja-JP' })
    buf = readFileSync(output)
    p = await openPptx(buf)
    slideXml = []
    for (let i = 1; i <= SLIDE_COUNT; i++) slideXml.push(await p.xml(`ppt/slides/slide${i}.xml`))
  }, 180_000)

  const shape = (name: string): Element => {
    for (const doc of slideXml) {
      const s = shapeNamed(doc, name)
      if (s) return s as unknown as Element
    }
    throw new Error(`no shape: ${name}`)
  }

  it('OPC 検査の error が 0', async () => {
    const results = await check(buf)
    expect(results.filter((r) => r.level === 'error')).toEqual([])
  })

  it('全スライドの <a:prstGeom prst> に 187 種がそろう', () => {
    const prsts = new Set(slideXml.flatMap((doc) => els(doc, 'a', 'prstGeom').map((g) => g.getAttribute('prst'))))
    expect(PRESET_NAMES.filter((n) => !prsts.has(n))).toEqual([])
  })

  it('図形名で引いた図形の prst が type と同じ（コネクタ 9 種と foldedCorner も）', () => {
    for (const n of ['bentConnector3', 'curvedConnector5', 'straightConnector1', 'foldedCorner', 'chartPlus', 'lineInv']) {
      const g = els(shape(n) as never, 'a', 'prstGeom')[0]
      expect(g?.getAttribute('prst'), n).toBe(n)
    }
  })

  it('adj を指定した図形の <a:avLst> に gd が入る', () => {
    const gds = (name: string) =>
      Object.fromEntries(els(els(shape(name) as never, 'a', 'avLst')[0] as never, 'a', 'gd').map((g) => [g.getAttribute('name'), g.getAttribute('fmla')]))
    expect(gds('adj-rightArrow')).toEqual({ adj1: 'val 80000', adj2: 'val 25000' })
    expect(gds('adj-star5')).toMatchObject({ adj: 'val 10000' })
    expect(gds('arc-half')).toMatchObject({ adj1: 'val 10800000', adj2: 'val 0' })
  })

  it('flipH / flipV が <a:xfrm> に出る', () => {
    const xfrm = (name: string) => els(shape(name) as never, 'a', 'xfrm')[0]
    expect(xfrm('flipH-rightArrow').getAttribute('flipH')).toBe('1')
    expect(xfrm('flipV-triangle').getAttribute('flipV')).toBe('1')
    expect(xfrm('flipHV-homePlate').getAttribute('flipH')).toBe('1')
    expect(xfrm('flipHV-homePlate').getAttribute('flipV')).toBe('1')
  })
})
