// 調整値（avLst の上書き）の正規化。Slidev の描画（components/PptShape.vue）と変換（build/convert.ts）が同じ値で形を作るため、
// 両方がここを通す。収集器（collect）は外の関数を import できないので、数だけを通して判断はしない。
import { PRESETS } from './presets.ts'

/** a:gd の fmla="val N" の N（ST_Coordinate32 の範囲。外は PowerPoint が修復に掛ける） */
export const ADJ_MIN = -2147483648
export const ADJ_MAX = 2147483647
/** roundRect の定義の `pin 0 adj 50000` */
const ROUND_RECT_ADJ_MAX = 50000

export interface NormalizedAdjust {
  /** 残った調整値（整数） */
  adj: Record<string, number>
  /** 捨てたキー（定義の avLst に無い名前、数でない値、丸めて 32 bit 整数の範囲外の値）。入力の順 */
  dropped: string[]
}

/**
 * 図形 `name` の調整値を、定義の avLst にある名前・数・丸めて 32 bit 整数の範囲に絞る。
 * 未知の図形名では全部捨てる（rect の avLst は空）。
 */
export function normalizeAdjust(name: string, adj: Record<string, unknown> | undefined): NormalizedAdjust {
  const known = new Set(Object.hasOwn(PRESETS, name) ? PRESETS[name].av.map(([n]) => n) : [])
  const out: Record<string, number> = {}
  const dropped: string[] = []
  for (const [k, v] of Object.entries(adj ?? {})) {
    const n = typeof v === 'number' ? Math.round(v) : NaN
    // -0.4 は丸めると -0 になるので 0 に揃える
    if (known.has(k) && n >= ADJ_MIN && n <= ADJ_MAX) out[k] = n === 0 ? 0 : n
    else dropped.push(k)
  }
  return { adj: out, dropped }
}

/**
 * roundRect の角の半径 px を adj に換算する（定義: 半径 = ss × adj / 100000）。定義の pin（0〜50000）に収め、整数に丸める。
 * 半径が有限の数でないか、短辺が 0 以下なら undefined
 */
export function radiusToAdj(radius: number | undefined, w: number, h: number): number | undefined {
  if (typeof radius !== 'number' || !Number.isFinite(radius)) return undefined
  const ss = Math.min(w, h)
  if (!(ss > 0)) return undefined
  return Math.max(0, Math.min(ROUND_RECT_ADJ_MAX, Math.round((radius / ss) * 100000)))
}
