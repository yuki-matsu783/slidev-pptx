// 調整値の正規化（src/shapes/adjust.ts）。Slidev の描画と変換が同じ値を使うための共通の関数
import { describe, expect, it } from 'vitest'
import { ADJ_MAX, ADJ_MIN, normalizeAdjust, radiusToAdj } from '../../packages/slidev-addon-pptx/src/shapes/adjust'

describe('normalizeAdjust', () => {
  it('定義の avLst にある名前は整数に丸めて残す', () => {
    expect(normalizeAdjust('rightArrow', { adj1: 50000.4, adj2: 0.4 })).toEqual({ adj: { adj1: 50000, adj2: 0 }, dropped: [] })
    expect(normalizeAdjust('wedgeRectCallout', { adj1: 0.4 })).toEqual({ adj: { adj1: 0 }, dropped: [] })
    expect(normalizeAdjust('circularArrow', { adj5: 12500.6 })).toEqual({ adj: { adj5: 12501 }, dropped: [] })
  })

  it('-0.4 は -0 でなく 0 になる', () => {
    expect(Object.is(normalizeAdjust('roundRect', { adj: -0.4 }).adj.adj, 0)).toBe(true)
  })

  it('定義に無い名前、数でない値（文字列・NaN・Infinity）は捨てて、入力の順に dropped に並べる', () => {
    expect(normalizeAdjust('star5', { foo: 1, adj: '10000', hf: 50000, vf: NaN, adj2: Infinity })).toEqual({ adj: { hf: 50000 }, dropped: ['foo', 'adj', 'vf', 'adj2'] })
  })

  it('丸めた後に 32 bit 整数の範囲を外れる値は捨てる（境界は丸めてから判定）', () => {
    expect(ADJ_MIN).toBe(-2147483648)
    expect(ADJ_MAX).toBe(2147483647)
    expect(normalizeAdjust('rightArrow', { adj1: 3000000000 })).toEqual({ adj: {}, dropped: ['adj1'] })
    expect(normalizeAdjust('foldedCorner', { adj: 2147483647.4 }).adj).toEqual({ adj: 2147483647 })
    expect(normalizeAdjust('foldedCorner', { adj: 2147483647.5 }).dropped).toEqual(['adj'])
    expect(normalizeAdjust('foldedCorner', { adj: -2147483648.4 }).adj).toEqual({ adj: -2147483648 })
    expect(normalizeAdjust('foldedCorner', { adj: -2147483649 }).dropped).toEqual(['adj'])
    expect(normalizeAdjust('donut', { adj: 1e12 }).dropped).toEqual(['adj'])
  })

  it('未知の図形名（Object の既定のプロパティ名を含む）では全部捨てる。rect の avLst は空', () => {
    expect(normalizeAdjust('foo', { adj: 1 })).toEqual({ adj: {}, dropped: ['adj'] })
    expect(normalizeAdjust('constructor', { adj: 1 })).toEqual({ adj: {}, dropped: ['adj'] })
    expect(normalizeAdjust('rect', { adj: 1 })).toEqual({ adj: {}, dropped: ['adj'] })
  })

  it('調整値が無ければ空', () => {
    expect(normalizeAdjust('rightArrow', undefined)).toEqual({ adj: {}, dropped: [] })
    expect(normalizeAdjust('rightArrow', {})).toEqual({ adj: {}, dropped: [] })
  })
})

describe('radiusToAdj', () => {
  it('半径 / 短辺 × 100000 を整数に丸める', () => {
    expect(radiusToAdj(8, 100, 40)).toBe(20000)
    expect(radiusToAdj(10, 60, 80)).toBe(16667)
    expect(radiusToAdj(7.3, 97.5, 41.3)).toBe(17676)
  })
  it('定義の pin（0〜50000）に収める', () => {
    expect(radiusToAdj(30, 100, 40)).toBe(50000)
    expect(radiusToAdj(-5, 100, 40)).toBe(0)
    expect(Object.is(radiusToAdj(-0.001, 100, 40), 0)).toBe(true)
  })
  it('半径が有限の数でないか、短辺が 0 以下なら undefined', () => {
    expect(radiusToAdj(undefined, 100, 40)).toBeUndefined()
    expect(radiusToAdj(NaN, 100, 40)).toBeUndefined()
    expect(radiusToAdj(8, 0, 40)).toBeUndefined()
    expect(radiusToAdj(8, 100, -1)).toBeUndefined()
  })
})
