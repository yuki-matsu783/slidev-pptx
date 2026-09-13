// PowerPoint の図形の定義（presetShapeDefinitions.xml を scripts/gen-presets.mjs で写したもの）の型。
// 値は XML の文字列のまま。数値の定数も、ガイドの名前も、評価器（geometry.ts）が解く。

/** ガイド 1 つ: [名前, 数式]。数式は空白 1 つ区切り（`*/ w adj 100000`） */
export type Guide = [name: string, fmla: string]

/** path の命令。座標・半径・角度は数値の文字列か、ガイド・組み込みの名前 */
export type PathCmd =
  | ['M', x: string, y: string]
  | ['L', x: string, y: string]
  | ['Q', x1: string, y1: string, x: string, y: string]
  | ['C', x1: string, y1: string, x2: string, y2: string, x: string, y: string]
  | ['A', wR: string, hR: string, stAng: string, swAng: string]
  | ['Z']

export type PathFill = 'none' | 'darken' | 'darkenLess' | 'lighten' | 'lightenLess'

export interface PresetPath {
  /** path 自身の座標系の幅。あれば座標を枠の幅へ拡大する */
  w?: number
  h?: number
  /** 無ければ norm */
  fill?: PathFill
  /** 無ければ線を引く */
  stroke?: false
  cmds: PathCmd[]
}

export interface PresetShape {
  /** 調整値の既定（avLst） */
  av: Guide[]
  /** 数式（gdLst）。順に評価する */
  gd: Guide[]
  /** 文字の枠 [l, t, r, b]。無ければ枠全体 */
  rect?: [l: string, t: string, r: string, b: string]
  paths: PresetPath[]
}
