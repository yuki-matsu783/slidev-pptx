// 収集器（ppt-components.md §2、§3）を fixture のデッキ plain.md に当てる。
// collect は Vue にも Node にも依存しない自己完結の関数なので、page.evaluate にそのまま渡せる（native-export.md §9）。
// Playwright は関数を文字列化して渡すので、collect はモジュール先頭の定数やヘルパを参照できない（1 関数の中に閉じる）。
import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import type { Browser, Page } from 'playwright-chromium'
import { collect } from '../../packages/slidev-addon-pptx/src/collect/index'
import type { Capture, SlideCapture, TextElement, ImageElement, LineElement } from '../../packages/slidev-addon-pptx/src/types'
import { PLAIN, PLAIN_READY, launch, openPrint, startSlidev } from './helpers'
import type { Running } from './helpers'
import { S, SLIDE_COUNT } from './slides'

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
  page = await openPrint(browser, server.base, { slides: SLIDE_COUNT, ready: PLAIN_READY })
  capture = await page.evaluate(collect)
}, 120_000)

afterAll(async () => {
  await browser?.close()
  await server?.close()
})

describe('collect: 入口（§2.1）', () => {
  it('キャンバスは .print-slide-container の実測（幅 980、高さは 551〜552。client は ceil で 552）', () => {
    expect(capture.canvas.width).toBe(980)
    expect(capture.canvas.height).toBeGreaterThanOrEqual(551)
    expect(capture.canvas.height).toBeLessThanOrEqual(552)
  })
  it(`スライドは ${SLIDE_COUNT} 枚、no は 1 始まり`, () => {
    expect(capture.slides.map((s) => s.no)).toEqual(Array.from({ length: SLIDE_COUNT }, (_, i) => i + 1))
  })
  it('backgroundColor は #rrggbb', () => {
    expect(slide(S.cover).backgroundColor).toMatch(/^#[0-9a-f]{6}$/)
  })
})

describe('collect: 表紙（cover。見出しは div.my-auto 越し）', () => {
  it('最初の h1 が title 候補、続く p と裸のテキストは同じ枠で body 候補', () => {
    const [title, body] = texts(slide(S.cover))
    expect(title.roleHint).toBe('title')
    expect(flat(title)).toBe('受入テスト')
    expect(title.paragraphs[0].kind).toBe('heading')
    expect(body.roleHint).toBe('body')
    expect(body.paragraphs).toHaveLength(2)
    expect(body.paragraphs[1].runs.map((r) => r.text).join('')).toContain('裸のテキスト')
  })
  it('opacity-70 の span の run は transparency 30（computed は "0.7" なので浮動小数の誤差を許す）', () => {
    const body = texts(slide(S.cover))[1]
    expect(body.paragraphs[1].runs[0].transparency).toBeCloseTo(30, 5)
  })
  it('box は キャンバス px で、title は左 56 px の少し左（h1 の -ml-[0.05em] で 3 px はみ出す）', () => {
    const x = texts(slide(S.cover))[0].box.x
    expect(x).toBeGreaterThan(50)
    expect(x).toBeLessThanOrEqual(56)
    expect(texts(slide(S.cover))[0].boxSource).toBe('measured')
  })
})

describe('collect: 箇条書きと run（§3.2、§3.4）', () => {
  let body: TextElement
  beforeAll(() => { body = texts(slide(S.bullets)).find((t) => t.roleHint === 'body')! })

  it('h2 → title 候補、ul → body 候補', () => {
    expect(texts(slide(S.bullets))[0].roleHint).toBe('title')
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
    // デッキは <a href="/3"> で書いてある。Markdown の [x](#3) は Slidev が href="##3" にするので、設計 §3.4 の
    // 「#N か /N」に当たらない（README の design-feedback 候補）
    expect(runs.find((r) => r.text === '3 枚目')?.link).toEqual({ slide: 3 })
    expect(runs.find((r) => r.text === '取り消し')?.strike).toBe(true)
    expect(runs.find((r) => r.text === '下線')?.underline).toBe(true)
    expect(runs.find((r) => r.text === '2' && r.sub)).toBeTruthy()
    expect(runs.find((r) => r.text === '2' && r.sup)).toBeTruthy()
    expect(runs.find((r) => r.text === 'マーク')?.highlight).toMatch(/^#/)
  })
  it('a の下線（border-bottom）は underline にしない', () => {
    const runs = body.paragraphs.flatMap((p) => p.runs)
    expect(runs.find((r) => r.text === '外部リンク')?.underline).toBe(false)
  })
  it('ol は number', () => {
    expect(body.paragraphs.filter((p) => p.kind === 'number')).toHaveLength(2)
  })
  it('箇条書きの中の表は無視して W-LI-BLOCK、枠は割れない', () => {
    expect(slide(S.bullets).warnings.some((w) => w.code === 'W-LI-BLOCK')).toBe(true)
    expect(slide(S.bullets).elements.some((e) => e.kind === 'table')).toBe(false)
  })
  it('v-click の div（装飾なし）の裸テキストは同じ枠の plain 段落', () => {
    expect(body.paragraphs.at(-1)?.runs.map((r) => r.text).join('')).toContain('クリックの文')
    expect(body.paragraphs.at(-1)?.kind).toBe('plain')
  })
})

describe('collect: img（§2.2 の 6）', () => {
  it('img: src は絶対 URL、alt。a > img は要素の link', () => {
    // デッキは裸の <img> で書いてある。Markdown の ![]() は <p><img></p> になり、規則 11 が先に当たって
    // 規則 6 に届かない（設計 §2.2 に「p の中身が img だけなら p ごと」の特例は無い。README の design-feedback 候補）。
    // <a> は 3 行に割って HTML ブロックにしてある（1 行に 2 タグ以上だと段落に落ちる）。規則 12 のタグ一覧に a が無い点も候補。
    // 置き場は「表と線」のスライド（箇条書きのスライドは 552 px を超え、下の画像がキャンバスの外に出る）
    const imgs = images(slide(S.table))
    expect(imgs).toHaveLength(2)
    expect(imgs[0].src).toMatch(/^http.*\/bg\.png$/)
    expect(imgs[0].alt).toBe('代替文字')
    expect(imgs[1].link).toEqual({ url: 'https://sli.dev' })
  })
})

describe('collect: two-cols（§2.1、§3.1、§3.8）', () => {
  it('左: h2 → title 候補、pre は自由配置の code 枠。body 候補は無い', () => {
    const ts = texts(slide(S.twoCols))
    expect(ts[0].roleHint).toBe('title')
    const code = ts.find((t) => t.paragraphs[0].kind === 'code')!
    expect(code.roleHint).toBeUndefined()
    expect(code.paragraphs.map((p) => p.runs[0].text)).toEqual(['const a = 1', '  return a'])
    expect(ts.some((t) => t.roleHint === 'body')).toBe(false)
  })
  it('右: 最初の枠が body2 候補で、見出しは中の heading 段落', () => {
    const b2 = texts(slide(S.twoCols)).find((t) => t.roleHint === 'body2')!
    expect(b2.paragraphs[0].kind).toBe('heading')
    expect(b2.paragraphs[1].kind).toBe('bullet')
  })
  it('引用は塗りのある別枠（区切りブロック）', () => {
    const quote = texts(slide(S.twoCols)).find((t) => flat(t).includes('引用の段落'))!
    expect(quote.roleHint).toBeUndefined()
    expect(quote.frame.fill?.color).toMatch(/^#/)
  })
  it('コード枠の色付け・行強調は捨て、frame に背景色と padding', () => {
    const code = texts(slide(S.twoCols)).find((t) => t.paragraphs[0].kind === 'code')!
    expect(new Set(code.paragraphs.flatMap((p) => p.runs.map((r) => r.color))).size).toBe(1)
    expect(code.frame.fill?.color).toMatch(/^#/)
    expect(code.frame.inset.some((v) => v > 0)).toBe(true)
  })
  it('行強調 {2} で強調されない行は .slidev-code-dishonored（opacity 0.3）だが、run に transparency を付けない', () => {
    // 設計 §3.4 の「祖先の opacity の積 → transparency」をそのまま当てると、強調していない行が薄くなる。§3.6 の例外（README の候補）
    const code = texts(slide(S.twoCols)).find((t) => t.paragraphs[0].kind === 'code')!
    for (const p of code.paragraphs) for (const r of p.runs) expect(r.transparency ?? 0).toBe(0)
  })
})

describe('collect: 表と線（§3.5、§2.2 の 10）', () => {
  it('table 要素: headerRows 1、colW / rowH は実測、罫線は下辺だけ', () => {
    const tbl = slide(S.table).elements.find((e) => e.kind === 'table')
    expect(tbl?.kind).toBe('table')
    if (tbl?.kind !== 'table') return
    expect(tbl.headerRows).toBe(1)
    expect(tbl.colW).toHaveLength(2)
    expect(tbl.rowH).toHaveLength(3)
    // Slidev の罫線は tr の border-b で、td/th の computed は 4 辺とも 0。セルが 0 なら tr → table へ遡る（README の候補）
    const c = tbl.rows[0][0]
    expect(c.border[2].width).toBeGreaterThan(0)
    expect(c.border[0].width).toBe(0)
    expect(c.border[1].width).toBe(0)
    expect(c.border[3].width).toBe(0)
    expect(tbl.rows[1][1].align).toBe('center')
    expect(tbl.rows[1][0].paragraphs[0].runs[0].code).toBe(true)
  })
  it('hr は line 要素（実測の上辺、水平）。高さ 0〜1 px なので、規則 1 の「rect が空」を w または h が 0 で判定すると届かないことに注意', () => {
    const line = slide(S.table).elements.find((e): e is LineElement => e.kind === 'line')!
    expect(line).toBeTruthy()
    expect(line.from.y).toBe(line.to.y)
    expect(line.to.x - line.from.x).toBeGreaterThan(800)
    expect(line.line.color).toMatch(/^#/)
  })
  it('表のあとの段落は別の枠（自由配置）', () => {
    const after = texts(slide(S.table)).find((t) => flat(t).includes('段落のあとの表'))!
    expect(after.roleHint).toBeUndefined()
  })
})

describe('collect: 数式と図（§2.2 の 5、§3.4）', () => {
  it('インライン数式は run に平坦化して W-MATH-INLINE', () => {
    const s = slide(S.math)
    expect(s.warnings.some((w) => w.code === 'W-MATH-INLINE')).toBe(true)
    const body = texts(s).find((t) => t.roleHint === 'body')!
    expect(flat(body)).toMatch(/E\s*=\s*mc/)
  })
  it('ブロック数式は p ごと画像（reason math）、Mermaid は画像（reason mermaid）', () => {
    const imgs = images(slide(S.math))
    expect(imgs.map((i) => i.reason).sort()).toEqual(['math', 'mermaid'])
    for (const i of imgs) expect(i.captureId).toBe(i.id)
  })
  it('置き換え要素には data-ppt-capture-id が付いている', async () => {
    const ids = images(slide(S.math)).map((i) => i.captureId!)
    for (const id of ids) expect(await page.locator(`[data-ppt-capture-id="${id}"]`).count()).toBe(1)
  })
})

describe('collect: 装飾つきの箱と未知の要素（§2.2）', () => {
  it('bg-blue の箱（子が装飾なしの div）は 13: 塗りのあるテキスト枠、2 段落', () => {
    const box = texts(slide(S.boxes)).find((t) => flat(t).includes('塗りのある箱'))!
    expect(box.frame.fill?.color).toMatch(/^#/)
    expect(box.frame.radius).toBeGreaterThan(0)
    expect(box.paragraphs).toHaveLength(2)
  })
  it('影つきの箱に区切りブロック（pre）があると 14: 装飾を捨てて中を歩き W-CSS', () => {
    const s = slide(S.boxes)
    expect(s.warnings.some((w) => w.code === 'W-CSS' && /shadow/i.test(w.message))).toBe(true)
    expect(texts(s).some((t) => flat(t).includes('const inside'))).toBe(true)
  })
  it('button は 15: 画像（unknown-element）', () => {
    expect(images(slide(S.boxes)).some((i) => i.reason === 'unknown-element')).toBe(true)
  })
  it('グラデーションだけの div は 4: 画像（gradient）', () => {
    expect(images(slide(S.boxes)).some((i) => i.reason === 'gradient')).toBe(true)
  })
  it('data-ppt-export="image" の囲みは 3: 画像（explicit）', () => {
    expect(images(slide(S.boxes)).some((i) => i.reason === 'explicit')).toBe(true)
  })
  it('取得できない画像は Capture では普通の image（W-IMAGE を出すのは Node）', () => {
    const missing = images(slide(S.boxes)).find((i) => i.src?.endsWith('/missing.png'))
    expect(missing).toBeTruthy()
  })
  it('手書きの data-ppt="text" は 2: PPT 部品として拾い、name と data-ppt-box の座標が使われる', () => {
    const hand = texts(slide(S.boxes)).find((t) => t.name === 'handwritten')!
    expect(hand.source).toBe('ppt')
    expect(hand.boxSource).toBe('prop')
    expect(hand.box).toMatchObject({ x: 600, y: 400, w: 300 })
    expect(hand.box.h).toBeGreaterThan(0)
  })
})

describe('collect: center / section / image-right / 背景画像つき cover', () => {
  it('center: h1 が title 候補、段落が body 候補、align center', () => {
    const ts = texts(slide(S.center))
    expect(ts[0].roleHint).toBe('title')
    expect(ts[0].paragraphs[0].align).toBe('center')
    expect(ts[1].roleHint).toBe('body')
  })
  it('section: 候補は付くが、対応表に無いことは Node が判断する（Capture はレイアウトを知らない）', () => {
    expect(texts(slide(S.section))[0].roleHint).toBe('title')
    expect((slide(S.section) as unknown as Record<string, unknown>).layout).toBeUndefined()
  })
  it('image-right: 右半分の background-image は画像要素（規則 4）、左の見出しは title 候補', () => {
    const s = slide(S.imageRight)
    const img = images(s).find((i) => i.src && !i.captureId)!
    expect(img.src).toMatch(/\/bg\.png$/)
    expect(img.box.x).toBeGreaterThanOrEqual(490)
    expect(texts(s)[0].roleHint).toBe('title')
  })
  it('cover の背景画像は .slidev-layout 自身の style なので Capture には出ない（Node が frontmatter から取る）', () => {
    expect(images(slide(S.coverBg))).toHaveLength(0)
  })
})

describe('collect: zoom と はみ出し（§2.3、§2.2 の 1）', () => {
  it('zoom 0.8: rect はそのままキャンバス px、computed の長さには zoom を掛ける', () => {
    const s = slide(S.zoom)
    expect(s.zoom).toBeCloseTo(0.8, 5)
    const title = texts(s)[0]
    expect(title.box.x).toBeCloseTo(56 * 0.8, 0)
    expect(title.paragraphs[0].runs[0].size).toBeCloseTo(30 * 0.8, 1)
  })
  it('はみ出した箱（装飾つき = 規則 13 の独立した枠）は測ったままの負の座標で写す（寄せるのは Node）。完全に外は W-HIDDEN、opacity 0 も W-HIDDEN', () => {
    const s = slide(S.offslide)
    const left = texts(s).find((t) => flat(t).includes('左にはみ出した'))!
    expect(left.box.x).toBeLessThan(0)
    expect(left.roleHint).toBeUndefined()
    expect(texts(s).some((t) => t.roleHint === 'body')).toBe(false)
    expect(texts(s).some((t) => flat(t).includes('完全に外'))).toBe(false)
    expect(s.warnings.filter((w) => w.code === 'W-HIDDEN').length).toBeGreaterThanOrEqual(2)
  })
  it('溢れる箱: fit.contentHeight > fit.boxHeight（縮小率は Node が決める）', () => {
    const over = texts(slide(S.offslide)).find((t) => flat(t).includes('溢れる本文'))!
    expect(over.fit.contentHeight).toBeGreaterThan(over.fit.boxHeight)
    expect(over.paragraphs[0].runs.filter((r) => r.breakAfter).length).toBe(5)
  })
})

describe('collect: Slidev の UI は無警告で飛ばす（§2.2 の 0）', () => {
  it('コードのコピーボタンが画像にも W-HIDDEN にもならない', () => {
    const s = slide(S.twoCols)
    expect(images(s)).toHaveLength(0)
    expect(s.warnings.some((w) => w.code === 'W-HIDDEN')).toBe(false)
  })
})
