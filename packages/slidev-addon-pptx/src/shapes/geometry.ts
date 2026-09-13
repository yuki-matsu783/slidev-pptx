// PowerPoint の図形の定義（presets.ts）を枠の大きさで評価し、SVG の path と文字の枠を返す。DOM に依存しない純関数。
// 数式: ECMA-376 Part 1 §20.1.9.11（gd）。角度は 60000 分の 1 度。
// arcTo の角度の換算は Apache POI（ArcToCommand）と同じ。
import { PRESETS } from './presets.ts'
import type { PathCmd, PresetShape } from './types.ts'

export type PresetFill = 'norm' | 'none' | 'darken' | 'darkenLess' | 'lighten' | 'lightenLess'

export interface PresetPathResult {
  /** SVG の path の d（px） */
  d: string
  fill: PresetFill
  stroke: boolean
}

export interface PresetGeometry {
  paths: PresetPathResult[]
  /** 文字の枠（px、枠の左上が原点） */
  textRect: { l: number; t: number; r: number; b: number }
}

export const PRESET_NAMES: readonly string[] = Object.freeze(Object.keys(PRESETS))

export function isPreset(name: string): boolean {
  return Object.hasOwn(PRESETS, name)
}

/** 評価は px をこの倍率で EMU 相当に広げて行う（定義に絶対値の定数が混じっても崩れないように） */
const SCALE = 9525
/** 60000 分の 1 度 → ラジアン */
const ANG = Math.PI / 180 / 60000
/** 大きさ 0 の枠は割り算で崩れるので、見えない大きさに寄せる */
const MIN_SIZE = 0.01
const NUM = /^-?\d+(\.\d+)?$/

/**
 * 図形 `name` を幅 `w` 高さ `h`（px）で評価する。
 * `adj` は avLst の既定を名前ごとに上書きする（定義に無い名前は無視）。未知の名前は例外。
 */
export function evalPreset(name: string, w: number, h: number, adj?: Record<string, number>): PresetGeometry {
  if (!isPreset(name)) throw new Error(`未知の図形: ${name}`)
  const shape = PRESETS[name]
  const vars = guides(name, shape, Math.max(w, MIN_SIZE) * SCALE, Math.max(h, MIN_SIZE) * SCALE, adj)
  const get = (v: string) => value(name, vars, v)
  const W = vars.get('w')!
  const H = vars.get('h')!
  const paths = shape.paths.map((p): PresetPathResult => ({
    d: pathData(p.cmds, get, p.w ? W / p.w : 1, p.h ? H / p.h : 1),
    fill: p.fill ?? 'norm',
    stroke: p.stroke !== false,
  }))
  const [l, t, r, b] = shape.rect ? shape.rect.map((v) => get(v) / SCALE) : [0, 0, W / SCALE, H / SCALE]
  return { paths, textRect: { l, t, r, b } }
}

/** 組み込みの値と avLst・gdLst を順に評価した表 */
function guides(name: string, shape: PresetShape, w: number, h: number, adj?: Record<string, number>): Map<string, number> {
  const ss = Math.min(w, h)
  const vars = new Map<string, number>([
    ['w', w], ['h', h], ['l', 0], ['t', 0], ['r', w], ['b', h],
    ['hc', w / 2], ['vc', h / 2], ['ss', ss], ['ls', Math.max(w, h)],
  ])
  for (const [n, f] of shape.av) {
    vars.set(n, adj && Object.hasOwn(adj, n) && Number.isFinite(adj[n]) ? adj[n] : formula(name, vars, f))
  }
  for (const [n, f] of shape.gd) vars.set(n, formula(name, vars, f))
  return vars
}

/** 数値・ガイド・組み込み（cdN、3cd4 など、wdN hdN ssdN）の値 */
function value(name: string, vars: Map<string, number>, v: string): number {
  if (NUM.test(v)) return Number(v)
  const got = vars.get(v)
  if (got !== undefined) return got
  let m = /^(\d*)cd(\d+)$/.exec(v)
  if (m) return ((m[1] ? Number(m[1]) : 1) * 21600000) / Number(m[2])
  m = /^(wd|hd|ssd)(\d+)$/.exec(v)
  if (m) return vars.get(m[1] === 'wd' ? 'w' : m[1] === 'hd' ? 'h' : 'ss')! / Number(m[2])
  throw new Error(`${name}: 未知の名前 ${v}`)
}

function formula(name: string, vars: Map<string, number>, fmla: string): number {
  const [op, ...args] = fmla.split(' ')
  const [x = 0, y = 0, z = 0] = args.map((a) => value(name, vars, a))
  switch (op) {
    case '*/': return (x * y) / z
    case '+-': return x + y - z
    case '+/': return (x + y) / z
    case '?:': return x > 0 ? y : z
    case 'abs': return Math.abs(x)
    case 'at2': return Math.atan2(y, x) / ANG
    case 'cat2': return x * Math.cos(Math.atan2(z, y))
    case 'cos': return x * Math.cos(y * ANG)
    case 'max': return Math.max(x, y)
    case 'min': return Math.min(x, y)
    case 'mod': return Math.sqrt(x * x + y * y + z * z)
    case 'pin': return y < x ? x : y > z ? z : y
    case 'sat2': return x * Math.sin(Math.atan2(z, y))
    case 'sin': return x * Math.sin(y * ANG)
    case 'sqrt': return Math.sqrt(x)
    case 'tan': return x * Math.tan(y * ANG)
    case 'val': return x
    default: throw new Error(`${name}: 未知の演算子 ${op}`)
  }
}

/** 小数 3 桁まで。-0 は 0 */
function num(v: number): string {
  const r = Math.round(v * 1000) / 1000
  return String(r === 0 ? 0 : r)
}

/**
 * path の命令を SVG の d にする。座標は EMU 相当で進め、出力で px に戻す。
 * sx sy は path 自身の座標系（path の w h）から枠への倍率。
 */
function pathData(cmds: PathCmd[], get: (v: string) => number, sx: number, sy: number): string {
  const out: string[] = []
  const X = (v: string) => get(v) * sx
  const Y = (v: string) => get(v) * sy
  const px = (x: number, y: number) => `${num(x / SCALE)} ${num(y / SCALE)}`
  let cx = 0
  let cy = 0
  let startX = 0
  let startY = 0
  for (const c of cmds) {
    switch (c[0]) {
      case 'M':
        cx = startX = X(c[1])
        cy = startY = Y(c[2])
        out.push(`M${px(cx, cy)}`)
        break
      case 'L':
        cx = X(c[1])
        cy = Y(c[2])
        out.push(`L${px(cx, cy)}`)
        break
      case 'Q':
        cx = X(c[3])
        cy = Y(c[4])
        out.push(`Q${px(X(c[1]), Y(c[2]))} ${px(cx, cy)}`)
        break
      case 'C':
        cx = X(c[5])
        cy = Y(c[6])
        out.push(`C${px(X(c[1]), Y(c[2]))} ${px(X(c[3]), Y(c[4]))} ${px(cx, cy)}`)
        break
      case 'A': {
        const wR = X(c[1])
        const hR = Y(c[2])
        const st = get(c[3])
        const sw = get(c[4])
        if (sw === 0) break
        // 見た目の角度 θ → 楕円の媒介変数の角度
        const param = (a: number) => Math.atan2(wR * Math.sin(a * ANG), hR * Math.cos(a * ANG))
        const t0 = param(st)
        const ox = cx - wR * Math.cos(t0)
        const oy = cy - hR * Math.sin(t0)
        // 360 度以上は SVG の A 1 つでは描けない（始点と終点が重なる）ので割る
        const n = Math.abs(sw) >= 21600000 ? Math.floor(Math.abs(sw) / 21600000) + 1 : 1
        const step = sw / n
        for (let i = 1; i <= n; i++) {
          const t = param(st + step * i)
          cx = ox + wR * Math.cos(t)
          cy = oy + hR * Math.sin(t)
          const large = Math.abs(step) > 10800000 ? 1 : 0
          const sweep = step > 0 ? 1 : 0
          out.push(`A${num(wR / SCALE)} ${num(hR / SCALE)} 0 ${large} ${sweep} ${px(cx, cy)}`)
        }
        break
      }
      case 'Z':
        cx = startX
        cy = startY
        out.push('Z')
        break
    }
  }
  return out.join(' ')
}
