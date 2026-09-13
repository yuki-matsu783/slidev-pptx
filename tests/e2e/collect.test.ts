// 収集器（ppt-components.md §2、§3）を fixture のデッキ plain.md に当てる。
// collect は Vue にも Node にも依存しない自己完結の関数なので、page.evaluate にそのまま渡せる（native-export.md §9）。
import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import type { Browser, Page } from 'playwright-chromium'
import { collect } from '../../packages/slidev-addon-pptx/src/collect/index'
import type { Capture, SlideCapture, TextElement, ImageElement } from '../../packages/slidev-addon-pptx/src/types'
import { PLAIN, launch, openPrint, startSlidev } from './helpers'
import type { Running } from './helpers'

let server: Running
let browser: Browser
let page: Page
let capture: Capture

const slide = (no: number): SlideCapture => capture.slides.find((s) => s.no === no)!
const texts = (s: SlideCapture) => s.elements.filter((e): e is TextElement => e.kind === 'text')
const images = (s: SlideCapture) => s.elements.filter((e): e is ImageElement => e.kind === 'image')
const flat = (t: TextElement) => t.paragraphs.map((p) => p.runs.map((r) => r.text).join('')).join('\n')

beforeAll(async () => {
  server = await startSlidev(PLAIN)
  browser = await launch()
  page = await openPrint(browser, server.base)
  capture = await page.evaluate(collect)
}, 120_000)

afterAll(async () => {
  await browser?.close()
  await server?.close()
})

describe('collect: 入口（§2.1）', () => {
  it('キャンバスは .print-slide-container の実測（980×552）', () => {
    expect(capture.canvas).toEqual({ width: 980, height: 552 })
  })
  it('スライドは 14 枚、no は 1 始まり', () => {
    expect(capture.slides.map((s) => s.no)).toEqual(Array.from({ length: 14 }, (_, i) => i + 1))
  })
  it('backgroundColor は #rrggbb', () => {
    expect(slide(1).backgroundColor).toMatch(/^#[0-9a-f]{6}$/)
  })
})

describe('collect: 表紙（cover。見出しは div.my-auto 越し）', () => {
  it('最初の h1 が title 候補、続く p と裸のテキストは同じ枠で body 候補', () => {
    const [title, body] = texts(slide(1))
    expect(title.roleHint).toBe('title')
    expect(flat(title)).toBe('受入テスト')
    expect(title.paragraphs[0].kind).toBe('heading')
    expect(body.roleHint).toBe('body')
    expect(body.paragraphs).toHaveLength(2)
    expect(body.paragraphs[1].runs.map((r) => r.text).join('')).toContain('裸のテキスト')
  })
  it('opacity-70 の span の run は transparency 30', () => {
    const body = texts(slide(1))[1]
    expect(body.paragraphs[1].runs[0].transparency).toBe(30)
  })
  it('box は キャンバス px で、title は左 56 px から', () => {
    expect(texts(slide(1))[0].box.x).toBeCloseTo(56, 0)
    expect(texts(slide(1))[0].boxSource).toBe('measured')
  })
})

describe('collect: 箇条書きと run（§3.2、§3.4）', () => {
  let body: TextElement
  beforeAll(() => { body = texts(slide(2)).find((t) => t.roleHint === 'body')! })

  it('h2 → title 候補、ul → body 候補', () => {
    expect(texts(slide(2))[0].roleHint).toBe('title')
    expect(body.paragraphs[0].kind).toBe('bullet')
  })
  it('入れ子は level 1', () => {
    expect(body.paragraphs.find((p) => p.level === 1)?.kind).toBe('bullet')
  })
  it('strong / em / code / a / del / u / sub / sup / mark が run に写る', () => {
    const runs = body.paragraphs.flatMap((p) => p.runs)
    expect(runs.find((r) => r.text === '太字')?.bold).toBe(true)
    expect(runs.find((r) => r.text === '斜体')?.italic).toBe(true)
    const code = runs.find((r) => r.text === 'コード')!
    expect(code.code).toBe(true)
    expect(code.highlight).toMatch(/^#/)
    expect(runs.find((r) => r.text === '外部リンク')?.link).toEqual({ url: 'https://sli.dev' })
    expect(runs.find((r) => r.text === '3 枚目')?.link).toEqual({ slide: 3 })
    expect(runs.find((r) => r.text === '取り消し')?.strike).toBe(true)
    expect(runs.find((r) => r.text === '下線')?.underline).toBe(true)
    expect(runs.find((r) => r.text === '2')?.sub).toBe(true)
    expect(runs.find((r) => r.text === '2' && r.sup)?.sup).toBe(true)
    expect(runs.find((r) => r.text === 'マーク')?.highlight).toMatch(/^#/)
  })
  it('a の下線（border-bottom）は underline にしない', () => {
    const runs = body.paragraphs.flatMap((p) => p.runs)
    expect(runs.find((r) => r.text === '外部リンク')?.underline).toBe(false)
  })
  it('ol は number', () => {
    expect(body.paragraphs.filter((p) => p.kind === 'number')).toHaveLength(2)
  })
  it('v-click の div（装飾なし）の裸テキストは同じ枠の plain 段落', () => {
    expect(body.paragraphs.at(-1)?.runs.map((r) => r.text).join('')).toContain('クリックの文')
    expect(body.paragraphs.at(-1)?.kind).toBe('plain')
  })
})

describe('collect: two-cols（§2.1、§3.1、§3.8）', () => {
  it('左: h2 → title 候補、pre は自由配置の code 枠。body 候補は無い', () => {
    const ts = texts(slide(3))
    expect(ts[0].roleHint).toBe('title')
    const code = ts.find((t) => t.paragraphs[0].kind === 'code')!
    expect(code.roleHint).toBeUndefined()
    expect(code.paragraphs.map((p) => p.runs[0].text)).toEqual(['const a = 1', '  return a'])
    expect(ts.some((t) => t.roleHint === 'body')).toBe(false)
  })
  it('右: 最初の枠が body2 候補で、見出しは中の heading 段落', () => {
    const b2 = texts(slide(3)).find((t) => t.roleHint === 'body2')!
    expect(b2.paragraphs[0].kind).toBe('heading')
    expect(b2.paragraphs[1].kind).toBe('bullet')
  })
  it('引用は塗りのある別枠（区切りブロック）', () => {
    const quote = texts(slide(3)).find((t) => flat(t).includes('引用の段落'))!
    expect(quote.roleHint).toBeUndefined()
    expect(quote.frame.fill?.color).toMatch(/^#/)
  })
  it('コード枠の色付け・行強調は捨て、frame に背景色と padding', () => {
    const code = texts(slide(3)).find((t) => t.paragraphs[0].kind === 'code')!
    expect(new Set(code.paragraphs.flatMap((p) => p.runs.map((r) => r.color))).size).toBe(1)
    expect(code.frame.fill?.color).toMatch(/^#/)
    expect(code.frame.inset.some((v) => v > 0)).toBe(true)
  })
})

describe('collect: 表（§3.5）', () => {
  it('table 要素: headerRows 1、colW / rowH は実測、罫線は下辺だけ', () => {
    const tbl = slide(4).elements.find((e) => e.kind === 'table')
    expect(tbl?.kind).toBe('table')
    if (tbl?.kind !== 'table') return
    expect(tbl.headerRows).toBe(1)
    expect(tbl.colW).toHaveLength(2)
    expect(tbl.rowH).toHaveLength(3)
    const c = tbl.rows[0][0]
    expect(c.border[2].width).toBeGreaterThan(0)
    expect(c.border[0].width).toBe(0)
    expect(c.border[1].width).toBe(0)
    expect(c.border[3].width).toBe(0)
    expect(tbl.rows[1][1].align).toBe('center')
    expect(tbl.rows[1][0].paragraphs[0].runs[0].code).toBe(true)
  })
  it('表のあとの段落は別の枠（自由配置）', () => {
    const after = texts(slide(4)).find((t) => flat(t).includes('段落のあとの表'))!
    expect(after.roleHint).toBeUndefined()
  })
})

describe('collect: 数式と図（§2.2 の 5、§3.4）', () => {
  it('インライン数式は run に平坦化して W-MATH-INLINE', () => {
    const s = slide(5)
    expect(s.warnings.some((w) => w.code === 'W-MATH-INLINE')).toBe(true)
    const body = texts(s).find((t) => t.roleHint === 'body')!
    expect(flat(body)).toMatch(/E\s*=\s*mc/)
  })
  it('ブロック数式は p ごと画像（reason math）、Mermaid は画像（reason mermaid）', () => {
    const imgs = images(slide(5))
    expect(imgs.map((i) => i.reason).sort()).toEqual(['math', 'mermaid'])
    for (const i of imgs) expect(i.captureId).toBe(i.id)
  })
  it('置き換え要素には data-ppt-capture-id が付いている', async () => {
    const ids = images(slide(5)).map((i) => i.captureId!)
    for (const id of ids) expect(await page.locator(`[data-ppt-capture-id="${id}"]`).count()).toBe(1)
  })
})

describe('collect: 装飾つきの箱と未知の要素（§2.2）', () => {
  it('bg-blue の箱（子が装飾なしの div）は 13: 塗りのあるテキスト枠、2 段落', () => {
    const box = texts(slide(6)).find((t) => flat(t).includes('塗りのある箱'))!
    expect(box.frame.fill?.color).toMatch(/^#/)
    expect(box.frame.radius).toBeGreaterThan(0)
    expect(box.paragraphs).toHaveLength(2)
  })
  it('影つきの箱に区切りブロック（pre）があると 14: 装飾を捨てて中を歩き W-CSS', () => {
    const s = slide(6)
    expect(s.warnings.some((w) => w.code === 'W-CSS' && /shadow/i.test(w.message))).toBe(true)
    expect(texts(s).some((t) => flat(t).includes("const inside"))).toBe(true)
  })
  it('button は 15: 画像（unknown-element）', () => {
    expect(images(slide(6)).some((i) => i.reason === 'unknown-element')).toBe(true)
  })
  it('data-ppt-export="image" の囲みは 3: 画像（explicit）', () => {
    expect(images(slide(6)).some((i) => i.reason === 'explicit')).toBe(true)
  })
  it('手書きの data-ppt="text" は 2: PPT 部品として拾い、name と data-ppt-box の座標が使われる', () => {
    const hand = texts(slide(6)).find((t) => t.name === 'handwritten')!
    expect(hand.source).toBe('ppt')
    expect(hand.boxSource).toBe('prop')
    expect(hand.box).toMatchObject({ x: 600, y: 400, w: 300 })
    expect(hand.box.h).toBeGreaterThan(0)
  })
})

describe('collect: center / section / image-right / 背景画像つき cover', () => {
  it('center: h1 が title 候補、段落が body 候補、align center', () => {
    const ts = texts(slide(7))
    expect(ts[0].roleHint).toBe('title')
    expect(ts[0].paragraphs[0].align).toBe('center')
    expect(ts[1].roleHint).toBe('body')
  })
  it('section: 候補は付くが、対応表に無いことは Node が判断する（Capture はレイアウトを知らない）', () => {
    expect(texts(slide(8))[0].roleHint).toBe('title')
    expect((slide(8) as unknown as Record<string, unknown>).layout).toBeUndefined()
  })
  it('image-right: 右半分の background-image は画像要素（規則 4）、左の見出しは title 候補', () => {
    const s = slide(9)
    const img = images(s).find((i) => i.src && !i.captureId)!
    expect(img.src).toMatch(/\/bg\.png$/)
    expect(img.box.x).toBeGreaterThanOrEqual(490)
    expect(texts(s)[0].roleHint).toBe('title')
  })
  it('cover の背景画像は .slidev-layout 自身の style なので Capture には出ない（Node が frontmatter から取る）', () => {
    expect(images(slide(10))).toHaveLength(0)
  })
})

describe('collect: zoom と はみ出し（§2.3、§2.2 の 1）', () => {
  it('zoom 0.8: rect はそのままキャンバス px、computed の長さには zoom を掛ける', () => {
    const s = slide(11)
    expect(s.zoom).toBeCloseTo(0.8, 5)
    const title = texts(s)[0]
    expect(title.box.x).toBeCloseTo(56 * 0.8, 0)
    expect(title.paragraphs[0].runs[0].size).toBeCloseTo(30 * 0.8, 1)
  })
  it('はみ出した箱は測ったままの負の座標で写す（寄せるのは Node）。完全に外は W-HIDDEN、opacity 0 も W-HIDDEN', () => {
    const s = slide(12)
    const left = texts(s).find((t) => flat(t).includes('左にはみ出した'))!
    expect(left.box.x).toBeLessThan(0)
    expect(texts(s).some((t) => flat(t).includes('完全に外'))).toBe(false)
    expect(s.warnings.filter((w) => w.code === 'W-HIDDEN').length).toBeGreaterThanOrEqual(2)
  })
})

describe('collect: Slidev の UI は無警告で飛ばす（§2.2 の 0）', () => {
  it('コードのコピーボタンが画像にも W-HIDDEN にもならない', () => {
    const s = slide(3)
    expect(images(s)).toHaveLength(0)
    expect(s.warnings.some((w) => w.code === 'W-HIDDEN')).toBe(false)
  })
})
