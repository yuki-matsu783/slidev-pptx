// PowerPoint の図形 187 種（PptShape の type）を並べたデッキ shapes.md。
// 収集: 全種が shape / line で出て、調整値・反転・矢じりが Capture に入る。描画: SVG の path と marker が壊れていない。
// 突き合わせ: DOM の path と evalPreset(normalizeAdjust)、PPTX の avLst と normalizeAdjust、矢じりと headEnd / tailEnd。
// 書き出し: prstGeom と avLst と xfrm の反転、W-SHAPE は境界の入力の枚だけ。
import { mkdtempSync, readFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import type { Browser, Page } from 'playwright-chromium'
import { collect } from '../../packages/slidev-addon-pptx/src/collect/index'
import { exportPptx } from '../../packages/slidev-addon-pptx/src/export'
import { check } from '../../packages/slidev-addon-pptx/src/opc/check'
import { normalizeAdjust, radiusToAdj } from '../../packages/slidev-addon-pptx/src/shapes/adjust'
import { PRESET_NAMES, evalPreset, isPreset } from '../../packages/slidev-addon-pptx/src/shapes/geometry'
import type { Capture, Element, LineElement, Report, ShapeElement } from '../../packages/slidev-addon-pptx/src/types'
import { els, openPptx, shapeNamed } from '../helpers/pptx'
import type { OpenedPptx } from '../helpers/pptx'
import { SHAPES, launch, openPrint, startSlidev } from './helpers'
import type { Running } from './helpers'

/** 格子 7 枚 + 見本 1 枚 + 境界の入力 1 枚 */
const SLIDE_COUNT = 9
const SAMPLE = 8
const EDGE = 9

/** 描画後の部品 1 つ分（突き合わせに使う）。line 以外は SVG の実寸と path の d、矢じりは marker の種類 */
interface DomShape {
  name: string
  type: string
  adj?: Record<string, unknown>
  radius?: number
  w: number
  h: number
  ds: string[]
  head: string | null
  tail: string | null
  /** 線を引く開いた path（または line 図形）があるか。矢じりはそこにだけ付く */
  open: boolean
}
let dom: DomShape[] = []
/**
 * 書き出しの突き合わせが使う描画の結果。collect の describe の beforeAll が集めるので、
 * `-t exportPptx` などで単独に回すと空になる。比べる図形が 0 件のまま通らないよう、件数が足りなければ落とす
 */
const domShapes = (): DomShape[] => {
  const shapes = dom.filter((d) => d.type !== 'line')
  if (shapes.length < 186) throw new Error(`描画の結果が ${shapes.length} 件しか無い（186 件以上が要る）。突き合わせは collect の describe と一緒に回す`)
  return dom
}

/** Slidev と PPTX が使うはずの図形名と調整値（未知の名前は rect、roundRect は radius の換算） */
const expected = (d: DomShape) => {
  const type = isPreset(d.type) ? d.type : 'rect'
  const { adj } = normalizeAdjust(type, d.adj)
  if (type === 'roundRect' && adj.adj === undefined) {
    const r = radiusToAdj(d.radius, d.w, d.h)
    if (r !== undefined) adj.adj = r
  }
  return { type, adj }
}

describe('collect: 図形 187 種', () => {
  let server: Running
  let browser: Browser
  let page: Page
  let capture: Capture

  beforeAll(async () => {
    server = await startSlidev(SHAPES)
    browser = await launch()
    page = await openPrint(browser, server.base, { slides: SLIDE_COUNT })
    capture = await page.evaluate(collect)
    dom = await page.evaluate(() =>
      Array.from(document.querySelectorAll<HTMLElement>('[data-ppt="shape"]')).map((el) => {
        const opts = JSON.parse(el.dataset.pptOpts ?? '{}')
        const svg = el.querySelector(':scope > svg')!
        const strokes = Array.from(svg.querySelectorAll('g path:not([stroke="none"]), line'))
        const kind = (attr: string) => {
          const ref = strokes.map((s) => s.getAttribute(attr)).find((v) => !!v)
          return ref ? (/-(\w+)\)$/.exec(ref)?.[1] ?? ref) : null
        }
        return {
          name: el.dataset.pptName ?? '',
          type: opts.type,
          adj: opts.adj,
          radius: opts.radius,
          w: Number(svg.getAttribute('width')),
          h: Number(svg.getAttribute('height')),
          ds: [...new Set(Array.from(svg.querySelectorAll('g path')).map((p) => p.getAttribute('d') ?? ''))],
          head: kind('marker-start'),
          tail: kind('marker-end'),
          open: strokes.some((s) => s.tagName.toLowerCase() === 'line' || !/Z\s*$/.test(s.getAttribute('d') ?? '')),
        }
      }),
    )
  }, 120_000)
  afterAll(async () => {
    await browser?.close()
    await server?.close()
  })

  const all = (): Element[] => capture.slides.flatMap((s) => s.elements)
  const named = (name: string) => all().find((e) => e.name === name)
  const onSlide = (no: number, name: string) => capture.slides.find((s) => s.no === no)!.elements.find((e) => e.name === name) as ShapeElement
  const sample = (name: string) => onSlide(SAMPLE, name)
  /** 描画後の部品のルート要素の offset*（回転・拡大の前の px） */
  const offsetOf = (name: string) =>
    page.evaluate((n) => {
      const el = document.querySelector<HTMLElement>(`[data-ppt-name="${n}"]`)!
      return { x: el.offsetLeft, y: el.offsetTop, w: el.offsetWidth, h: el.offsetHeight, minH: parseFloat(getComputedStyle(el).minHeight) || 0 }
    }, name)

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

  it('描画: 全図形の SVG の path の d が空でなく NaN を含まない。marker の id は重複せず、url(#…) の参照先がある', async () => {
    const r = await page.evaluate(() => {
      const bad: string[] = []
      const svgs = Array.from(document.querySelectorAll('svg.ppt-shape-svg'))
      for (const svg of svgs) {
        const name = svg.closest('[data-ppt]')?.getAttribute('data-ppt-name')
        const paths = Array.from(svg.querySelectorAll('g path'))
        if (!paths.length) bad.push(`${name}: path が無い`)
        for (const p of paths) {
          const d = p.getAttribute('d') ?? ''
          if (!d.trim()) bad.push(`${name}: d が空`)
          else if (/NaN|Infinity|undefined/.test(d)) bad.push(`${name}: ${d.slice(0, 60)}`)
        }
      }
      const ids = Array.from(document.querySelectorAll('marker')).map((m) => m.id)
      const dup = ids.filter((id, i) => ids.indexOf(id) !== i)
      const refs = Array.from(document.querySelectorAll('[marker-start], [marker-end]')).flatMap((e) => [e.getAttribute('marker-start'), e.getAttribute('marker-end')].filter((v): v is string => !!v))
      const unresolved = refs.filter((ref) => {
        const m = /^url\(#(.+)\)$/.exec(ref)
        return !m || !document.getElementById(m[1])
      })
      return { svgs: svgs.length, markers: ids.length, refs: refs.length, bad, dup, unresolved }
    })
    expect(r.svgs).toBeGreaterThanOrEqual(186) // 187 種のうち line 以外 + 見本
    expect(r.bad).toEqual([])
    expect(r.markers).toBeGreaterThan(0)
    expect(r.refs).toBeGreaterThan(0)
    expect(r.dup).toEqual([])
    expect(r.unresolved).toEqual([])
  })

  it('突き合わせ: DOM の SVG の path は evalPreset(図形名, 実寸, normalizeAdjust の結果) と一致する（PPTX に書く調整値で形を作る）', () => {
    const shapes = dom.filter((d) => d.type !== 'line')
    expect(shapes.length).toBeGreaterThanOrEqual(186)
    const bad: string[] = []
    for (const d of shapes) {
      const { type, adj } = expected(d)
      const want = new Set(evalPreset(type, d.w, d.h, adj).paths.map((p) => p.d))
      const miss = d.ds.filter((x) => !want.has(x))
      if (!d.ds.length || miss.length) bad.push(`${d.name}: ${miss.length}/${d.ds.length} が一致しない`)
    }
    expect(bad).toEqual([])
    // 小数の調整値をそのまま評価すると形が変わる入力で、検査が違いを見分けられること
    const arc = dom.find((d) => d.name === 'edge-blockarc')!
    expect(evalPreset('blockArc', arc.w, arc.h, arc.adj as Record<string, number>).paths.map((p) => p.d)).not.toEqual(evalPreset('blockArc', arc.w, arc.h, expected(arc).adj).paths.map((p) => p.d))
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

  it('adj は data-ppt-opts の値のまま Capture に入る（丸め・範囲・数でない値の判断は Node の変換がする）', () => {
    expect(sample('adj-roundRect').adj).toEqual({ adj: 50000 })
    expect(sample('adj-rightArrow').adj).toEqual({ adj1: 80000, adj2: 25000 })
    expect(sample('arc-half').adj).toEqual({ adj1: 10800000, adj2: 0 })
    expect(sample('text-ellipse').adj).toBeUndefined()
    expect(onSlide(EDGE, 'edge-big').adj).toEqual({ adj1: 3000000000 })
    expect(onSlide(EDGE, 'edge-frac').adj).toEqual({ adj1: 0.4, adj2: 80000.6 })
    expect(onSlide(EDGE, 'edge-str').adj).toEqual({ adj: '10000', hf: 50000 })
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
    expect(sample('arc-half').arrow).toEqual({ tail: 'stealth' })
    expect(sample('brace').arrow).toBeUndefined()
    expect(sample('conn-bent3').frame.line).toMatchObject({ color: '#7c3aed', width: 2 })
    const line = named('line') as LineElement
    expect(line.line.head).toBeUndefined()
  })

  it('知らない矢じりは Slidev でも arrow（開いた V 字）で描く', () => {
    expect(dom.find((d) => d.name === 'edge-bogus')).toMatchObject({ head: 'oval', tail: 'arrow' })
    expect(dom.find((d) => d.name === 'edge-line-bogus')).toMatchObject({ head: 'arrow', tail: 'stealth' })
  })

  it('h なしの図形は box の h を実測する（中身が 2 行で min-height の 2em より高い）', async () => {
    const e = sample('text-ellipse-auto')
    const d = await offsetOf('text-ellipse-auto')
    expect(e.boxSource).toBe('prop')
    // min-height で決まった高さでは実測の検査にならないので、中身がそれより高いことを先に確かめる
    expect(d.h).toBeGreaterThan(d.minH + 4)
    expect(Math.abs(e.box.h - d.h)).toBeLessThanOrEqual(1)
  })

  it('rotate と実測の幅: box は回転前の枠で、w は描画後の幅と一致する（onMounted の古い値のままにならない）', async () => {
    const e = sample('rot-wauto')
    const d = await offsetOf('rot-wauto')
    expect(e.rotate).toBe(30)
    expect(e.box).toMatchObject({ x: 760, y: 440, h: 60 })
    expect(Math.abs(e.box.w - d.w)).toBeLessThanOrEqual(0.5)
  })

  const H_ONLY = ['honly-star4', 'honly-pie', 'rot-wauto']

  it('h だけ指定した図形: 文字は枠全体に置き（margin なし）、幅は文字の幅程度に収まる', async () => {
    for (const n of H_ONLY) {
      const d = await offsetOf(n)
      expect(d.w, n).toBeGreaterThan(0)
      expect(d.w, n).toBeLessThan(200)
      expect(sample(n).box.w, n).toBeLessThan(200)
      const margin = await page.evaluate((name) => document.querySelector<HTMLElement>(`[data-ppt-name="${name}"] > .ppt-shape-text`)!.style.margin, n)
      expect(margin, n).toBe('')
    }
  })

  it('h だけ指定した図形: 1 秒おいて測っても幅と SVG の大きさが変わらない（測り直しが循環しない）', async () => {
    const read = () =>
      page.evaluate((names) => names.map((n) => {
        const el = document.querySelector<HTMLElement>(`[data-ppt-name="${n}"]`)!
        const svg = el.querySelector('svg.ppt-shape-svg')!
        return [n, el.offsetWidth, el.offsetHeight, svg.getAttribute('width'), svg.getAttribute('height'), el.dataset.pptBox]
      }), H_ONLY)
    const before = await read()
    await page.waitForTimeout(1000)
    expect(await read()).toEqual(before)
  })

  it('ブラウザで ResizeObserver loop の警告・エラーが出ない', async () => {
    // openPrint の中で作るページには読み込み前に聞き耳を立てられないので、別のページで開き直して数える
    const context = await browser.newContext({ viewport: { width: 980, height: 552 * SLIDE_COUNT } })
    const p = await context.newPage()
    const messages: string[] = []
    p.on('console', (m) => messages.push(m.text()))
    p.on('pageerror', (e) => messages.push(e.message))
    await p.addInitScript(() => {
      window.addEventListener('error', (e) => console.error(`window error: ${e.message}`))
    })
    try {
      await p.goto(`${server.base}/print?print=true`, { waitUntil: 'networkidle' })
      await p.waitForSelector('[data-ppt-name="honly-star4"]', { timeout: 60_000 })
      await p.waitForTimeout(3000)
    } finally {
      await context.close()
    }
    expect(messages.filter((m) => /ResizeObserver loop/.test(m))).toEqual([])
  })

  it('w と h の両方がある図形は、文字の div が定義の文字の枠（textRect）の中に入る', async () => {
    for (const n of ['text-ellipse', 'text-triangle', 'text-callout', 'adj-rightArrow']) {
      const e = sample(n)
      const rect = evalPreset(e.shape, e.box.w, e.box.h, normalizeAdjust(e.shape, e.adj).adj).textRect
      const div = await page.evaluate((name) => {
        const t = document.querySelector<HTMLElement>(`[data-ppt-name="${name}"] > .ppt-shape-text`)!
        return { l: t.offsetLeft, t: t.offsetTop, r: t.offsetLeft + t.offsetWidth, b: t.offsetTop + t.offsetHeight }
      }, n)
      // 左右は枠いっぱいに広がり、上下は中身の高さで枠の中に寄る（offset* は整数なので 1 px の誤差を許す）
      expect(Math.abs(div.l - rect.l), `${n} l`).toBeLessThanOrEqual(1)
      expect(Math.abs(div.r - rect.r), `${n} r`).toBeLessThanOrEqual(1)
      expect(div.t, `${n} t`).toBeGreaterThanOrEqual(rect.t - 1)
      expect(div.b, `${n} b`).toBeLessThanOrEqual(rect.b + 1)
    }
    // 三角の文字の枠は下半分なので、左右とも枠の端から離れる（枠全体に置いていないことの確認）
    const tri = evalPreset('triangle', 220, 150).textRect
    expect(tri.l).toBeGreaterThan(1)
    expect(tri.t).toBeGreaterThan(1)
  })
})

describe('exportPptx: 図形 187 種', () => {
  let report: Report
  let buf: Buffer
  let slideXml: Document[]

  beforeAll(async () => {
    const output = join(mkdtempSync(join(tmpdir(), 'slidev-pptx-')), 'shapes.pptx')
    report = await exportPptx({ entry: SHAPES, output, lang: 'ja-JP' })
    buf = readFileSync(output)
    const p: OpenedPptx = await openPptx(buf)
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

  it('W-SHAPE: 格子と見本（1〜8 枚目）は 0 件。境界の入力（9 枚目）は捨てた調整値・知らない矢じり・知らない図形の分だけ 1 件ずつ', () => {
    const ws = report.warnings.filter((w) => w.code === 'W-SHAPE')
    expect(ws.filter((w) => w.slide !== EDGE)).toEqual([])
    expect(ws.map((w) => w.name).sort()).toEqual(['edge-big', 'edge-bogus', 'edge-huge', 'edge-line-bogus', 'edge-rr-big', 'edge-str', 'edge-unknown'])
    // 数でない値も、範囲外と同じ変換の文面で出る
    expect(ws.find((w) => w.name === 'edge-str')?.message).toBe('図形 "star5" の調整値 adj は定義に無いか、数でないか、範囲外なので捨てた')
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

  it('突き合わせ: 全図形の prst と <a:avLst> の gd が、Slidev が形を作った図形名と normalizeAdjust の結果に一致する', () => {
    const bad: string[] = []
    for (const d of domShapes().filter((x) => x.type !== 'line')) {
      const want = expected(d)
      const sp = shape(d.name)
      const geom = els(sp as never, 'a', 'prstGeom')[0]
      const got = Object.fromEntries(els(geom as never, 'a', 'gd').map((g) => [g.getAttribute('name'), Number((g.getAttribute('fmla') ?? '').replace(/^val /, ''))]))
      if (geom?.getAttribute('prst') !== want.type) bad.push(`${d.name}: prst ${geom?.getAttribute('prst')} != ${want.type}`)
      if (JSON.stringify(Object.entries(got).sort()) !== JSON.stringify(Object.entries(want.adj).sort())) bad.push(`${d.name}: gd ${JSON.stringify(got)} != ${JSON.stringify(want.adj)}`)
    }
    expect(bad).toEqual([])
  })

  it('突き合わせ: 開いた path のある図形と線の矢じりの種類が <a:headEnd> / <a:tailEnd> と一致する', () => {
    const bad: string[] = []
    const opened = domShapes().filter((d) => d.open)
    expect(opened.some((d) => d.head || d.tail)).toBe(true)
    for (const d of opened) {
      const ln = els(shape(d.name) as never, 'a', 'ln')[0]
      const end = (local: string) => {
        const v = ln ? (els(ln as never, 'a', local)[0]?.getAttribute('type') ?? null) : null
        return v === 'none' ? null : v
      }
      if (end('headEnd') !== d.head || end('tailEnd') !== d.tail) bad.push(`${d.name}: pptx ${end('headEnd')}/${end('tailEnd')} dom ${d.head}/${d.tail}`)
    }
    expect(bad).toEqual([])
  })

  it('flipH / flipV が <a:xfrm> に出る', () => {
    const xfrm = (name: string) => els(shape(name) as never, 'a', 'xfrm')[0]
    expect(xfrm('flipH-rightArrow').getAttribute('flipH')).toBe('1')
    expect(xfrm('flipV-triangle').getAttribute('flipV')).toBe('1')
    expect(xfrm('flipHV-homePlate').getAttribute('flipH')).toBe('1')
    expect(xfrm('flipHV-homePlate').getAttribute('flipV')).toBe('1')
  })
})
