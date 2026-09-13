// 座標の換算（native-export.md §4.1）、margin の並べ替え（§4.2、§4.3）、dash の変換（§4.3）
import { describe, expect, it } from 'vitest'
import {
  SLIDE_W_EMU, SLIDE_W_IN, emu, pt, inch, clampBox, textMargin, cellMargin, toDashType,
} from '../../packages/slidev-addon-pptx/src/build/units'

const canvas = { width: 980, height: 552 }

describe('build/units: 定数', () => {
  it('スライド幅は LAYOUT_WIDE の 12192000 EMU で、インチはそこから導く', () => {
    expect(SLIDE_W_EMU).toBe(12192000)
    expect(SLIDE_W_IN).toBeCloseTo(12192000 / 914400, 10)
  })
})

describe('build/units: emu(px)', () => {
  it('幅だけで換算する（980 px → 12440.816… EMU/px）', () => {
    expect(emu(980, canvas)).toBe(12192000)
    expect(emu(56, canvas)).toBe(Math.round(56 * 12192000 / 980))
    expect(emu(490, canvas)).toBe(6096000)
  })
  it('0 は 0 のまま', () => {
    expect(emu(0, canvas)).toBe(0)
  })
  it('1 以上 100 以下の EMU は 101 に切り上げる（PptxGenJS が 100 以下をインチ扱いするため）', () => {
    // 0.001 px → 12 EMU → 101
    expect(emu(0.001, canvas)).toBe(101)
    expect(emu(100 / (12192000 / 980), canvas)).toBe(101)
    expect(emu(101 / (12192000 / 980) + 1e-9, canvas)).toBeGreaterThanOrEqual(101)
  })
  it('キャンバス幅が違えば定数も変わる', () => {
    expect(emu(1024, { width: 1024, height: 576 })).toBe(12192000)
  })
})

describe('build/units: pt(px) と inch(px)', () => {
  it('pt は小数 1 桁（980 px なら px × 0.9796）', () => {
    expect(pt(36, canvas)).toBe(35.3)
    expect(pt(17.6, canvas)).toBe(17.2)
    expect(pt(60, canvas)).toBe(58.8)
  })
  it('inch は丸めない（rectRadius 用）', () => {
    expect(inch(73.5, canvas)).toBeCloseTo(1, 6)
    expect(inch(6, canvas)).toBeCloseTo(6 / 73.5, 6)
  })
})

describe('build/units: clampBox（スライドの外に掛かる要素を内側に寄せる）', () => {
  it('内側の箱はそのまま、移動量 0、除外しない', () => {
    const r = clampBox({ x: 10, y: 10, w: 100, h: 50 }, canvas)
    expect(r).toEqual({ box: { x: 10, y: 10, w: 100, h: 50 }, shift: 0, dropped: false })
  })
  it('x が負なら 0 に寄せ、右端は変えない', () => {
    const r = clampBox({ x: -30, y: 300, w: 200, h: 40 }, canvas)
    expect(r.box).toEqual({ x: 0, y: 300, w: 170, h: 40 })
    expect(r.shift).toBe(30)
    expect(r.dropped).toBe(false)
  })
  it('y が負でも同じ', () => {
    const r = clampBox({ x: 10, y: -8, w: 100, h: 50 }, canvas)
    expect(r.box).toEqual({ x: 10, y: 0, w: 100, h: 42 })
    expect(r.shift).toBe(8)
  })
  it('右端・下端のはみ出しは w / h を縮める（左端・上端は動かさない）', () => {
    const r = clampBox({ x: 900, y: 500, w: 200, h: 100 }, canvas)
    expect(r.box).toEqual({ x: 900, y: 500, w: 80, h: 52 })
    expect(r.shift).toBe(120)
  })
  it('寄せた結果 w か h が 0 以下なら除外する', () => {
    expect(clampBox({ x: -300, y: 400, w: 100, h: 40 }, canvas).dropped).toBe(true)
    expect(clampBox({ x: 10, y: 600, w: 100, h: 40 }, canvas).dropped).toBe(true)
  })
  it('shift は最大の移動量（警告のしきい値 5 px の判定に使う）', () => {
    expect(clampBox({ x: -2, y: -4, w: 100, h: 50 }, canvas).shift).toBe(4)
  })
})

describe('build/units: margin の並べ替え', () => {
  it('textMargin: Capture の [上, 右, 下, 左] px → PptxGenJS の [左, 右, 下, 上] pt', () => {
    expect(textMargin([12, 16, 12, 16], canvas)).toEqual([pt(16, canvas), pt(16, canvas), pt(12, canvas), pt(12, canvas)])
    expect(textMargin([1, 2, 3, 4], canvas)).toEqual([pt(4, canvas), pt(2, canvas), pt(3, canvas), pt(1, canvas)])
  })
  it('cellMargin: [上, 右, 下, 左] のまま pt に。上は 1 pt 未満なら 1 に切り上げる', () => {
    expect(cellMargin([12, 8, 12, 8], canvas)).toEqual([pt(12, canvas), pt(8, canvas), pt(12, canvas), pt(8, canvas)])
    expect(cellMargin([0, 8, 12, 8], canvas)).toEqual([1, pt(8, canvas), pt(12, canvas), pt(8, canvas)])
    expect(cellMargin([0.5, 0, 0, 0], canvas)[0]).toBe(1)
  })
})

describe('build/units: toDashType', () => {
  it("'dot' は PptxGenJS の 'sysDot' に、他はそのまま", () => {
    expect(toDashType('dot')).toBe('sysDot')
    expect(toDashType('dash')).toBe('dash')
    expect(toDashType('solid')).toBe('solid')
  })
})
