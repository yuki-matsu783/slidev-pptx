// 座標の換算（native-export.md §4.1）、margin の並べ替え（§4.2、§4.3）、dash の写し（§4.3）
import type { Box, Canvas, Dash } from '../types.ts'

/** LAYOUT_WIDE の幅（EMU）。定数はこの 1 つから導く */
export const SLIDE_W_EMU = 12192000
export const SLIDE_H_EMU = 6858000
export const EMU_PER_INCH = 914400
export const SLIDE_W_IN = SLIDE_W_EMU / EMU_PER_INCH

/**
 * キャンバス px → EMU（整数）。幅だけで換算する。
 * PptxGenJS は `x y w h` を 100 未満ならインチ、`colW rowH` を 100 以下ならインチとみなすので、
 * 1 以上 100 以下は 101 に切り上げる（0.001 インチ未満。見えない）。0 は 0 のまま。
 */
export function emu(px: number, canvas: Canvas): number {
  const v = Math.round((px * SLIDE_W_EMU) / canvas.width)
  if (v > 0 && v <= 100) return 101
  return v
}

/** キャンバス px → pt（小数 1 桁） */
export function pt(px: number, canvas: Canvas): number {
  return Math.round(((px * 72 * SLIDE_W_IN) / canvas.width) * 10) / 10
}

/** キャンバス px → インチ（丸めない。rectRadius 用） */
export function inch(px: number, canvas: Canvas): number {
  return (px * SLIDE_W_IN) / canvas.width
}

export interface Clamped {
  box: Box
  /** 寄せた量と縮めた量の最大値（px）。5 px 以上なら W-OFFSLIDE */
  shift: number
  /** 寄せた結果 w か h が 0 以下（全体がスライドの外） */
  dropped: boolean
}

/** スライドの外に掛かる要素を内側に寄せる。左端・上端は 0 に、右端・下端は縮める */
export function clampBox(box: Box, canvas: Canvas): Clamped {
  let { x, y, w, h } = box
  let shift = 0
  if (x < 0) {
    shift = Math.max(shift, -x)
    w += x
    x = 0
  }
  if (y < 0) {
    shift = Math.max(shift, -y)
    h += y
    y = 0
  }
  if (x + w > canvas.width) {
    shift = Math.max(shift, x + w - canvas.width)
    w = canvas.width - x
  }
  if (y + h > canvas.height) {
    shift = Math.max(shift, y + h - canvas.height)
    h = canvas.height - y
  }
  return { box: { x, y, w, h }, shift, dropped: w <= 0 || h <= 0 }
}

/** テキスト枠の margin: Capture の [上, 右, 下, 左] px → PptxGenJS 4.0.1 の実装順 [左, 右, 下, 上] pt */
export function textMargin(inset: [number, number, number, number], canvas: Canvas): [number, number, number, number] {
  const [t, r, b, l] = inset
  return [pt(l, canvas), pt(r, canvas), pt(b, canvas), pt(t, canvas)]
}

/** 表セルの margin: [上, 右, 下, 左] のまま pt。上が 1 pt 未満だと 4 辺ともインチ扱いになるので 1 に切り上げる */
export function cellMargin(inset: [number, number, number, number], canvas: Canvas): [number, number, number, number] {
  const [t, r, b, l] = inset
  return [Math.max(1, pt(t, canvas)), pt(r, canvas), pt(b, canvas), pt(l, canvas)]
}

/** Capture の dash → PptxGenJS の dashType（'dot' は型に無いので 'sysDot'） */
export function toDashType(dash: Dash): 'solid' | 'dash' | 'sysDot' {
  return dash === 'dot' ? 'sysDot' : dash
}
