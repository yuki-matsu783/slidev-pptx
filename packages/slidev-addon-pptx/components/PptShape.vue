<script setup lang="ts">
// PPT 部品: 図形（wip/design/ppt-components.md §1.3）。
// line 以外は PowerPoint の図形の定義（src/shapes）を枠の実寸 px で評価し、inline SVG で描く。文字は上に重ねる。
// 文字の枠: h を props で決めたときだけ定義の rect に置く。h が内容で決まるときは枠全体（文字の枠が高さに依存して循環するため）。
// type="line" は対角線。矢じりの marker は部品ごとに一意な id を持つ（/print は全スライドを同時に描く）。
import { computed, onBeforeUnmount, onMounted, ref, useId } from 'vue'
import { evalPreset, isPreset } from '../src/shapes/geometry.ts'
import type { PresetFill, PresetGeometry } from '../src/shapes/geometry.ts'

type Dash = 'solid' | 'dash' | 'dot'
type Arrow = 'none' | 'arrow' | 'triangle' | 'oval' | 'diamond'
interface LineProps {
  color?: string
  width?: number
  dash?: Dash
  head?: Arrow
  tail?: Arrow
}

const props = withDefaults(
  defineProps<{
    x?: number
    y?: number
    w?: number
    h?: number
    export?: 'native' | 'image'
    name?: string
    /** PowerPoint の図形の名前（ECMA-376 の prst 名。187 種）。未知の名前は rect で描く */
    type?: string
    /** 調整値。キーは定義の avLst の名前（`adj` `adj1` …）、値は ECMA の単位（例 50000） */
    adj?: Record<string, number>
    /** 左右反転（文字は反転しない） */
    flipH?: boolean
    /** 上下反転（文字は 180 度回る） */
    flipV?: boolean
    fill?: string
    line?: string | LineProps
    radius?: number
    rotate?: number
    /** 図形の中の文字の内側余白 px。PPTX の inset にも反映する */
    padding?: number | [number, number, number, number]
    align?: 'left' | 'center' | 'right' | 'justify'
    valign?: 'top' | 'middle' | 'bottom'
    size?: number
    color?: string
  }>(),
  { export: 'native', name: '', type: 'rect', flipH: false, flipV: false, fill: '#ffffff', line: '#000000', radius: 8, rotate: 0, padding: () => [0, 8, 0, 8], align: 'center', valign: 'middle' },
)

const uid = useId()
const root = ref<HTMLElement | null>(null)
// rotate のときは回転前の枠（offset*）を測って data-ppt-box に載せる（回転後の外接矩形を使わない）
const measured = ref<Record<string, number>>({})
// SVG の実寸。w h が props に無い辺は ResizeObserver で測る（offset* は回転・拡大の前の px）
const size = ref({ w: 0, h: 0 })
let observer: ResizeObserver | undefined
const measure = () => {
  const el = root.value
  if (!el) return
  // 0.5 px 未満の変化は無視する（文字の枠の余白が幅に効くとき、測り直しが続かないように）
  if (Math.abs(el.offsetWidth - size.value.w) >= 0.5 || Math.abs(el.offsetHeight - size.value.h) >= 0.5) size.value = { w: el.offsetWidth, h: el.offsetHeight }
}
onMounted(() => {
  const el = root.value
  if (!el) return
  if (props.rotate) measured.value = { x: el.offsetLeft, y: el.offsetTop, w: el.offsetWidth, h: el.offsetHeight }
  if (props.type !== 'line') {
    measure()
    if (props.w === undefined || props.h === undefined) {
      observer = new ResizeObserver(measure)
      observer.observe(el)
    }
  }
  if (!isPreset(props.type)) console.warn(`[PptShape] 未知の type "${props.type}" は rect で描く`)
})
onBeforeUnmount(() => observer?.disconnect())

const box = computed(() => {
  const b: Record<string, number> = { ...measured.value }
  if (props.x !== undefined) b.x = props.x
  if (props.y !== undefined) b.y = props.y
  if (props.w !== undefined) b.w = props.w
  if (props.h !== undefined) b.h = props.h
  return b
})
const positioned = computed(() => props.x !== undefined || props.y !== undefined || props.w !== undefined || props.h !== undefined)
const lineProps = computed<LineProps | undefined>(() => {
  if (!props.line || props.line === 'none') return undefined
  return typeof props.line === 'string' ? { color: props.line, width: 1 } : { width: 1, ...props.line }
})
const fillColor = computed(() => (props.fill && props.fill !== 'none' ? props.fill : 'none'))
const dashArray = computed(() => {
  const d = lineProps.value?.dash
  const w = lineProps.value?.width ?? 1
  return d === 'dash' ? `${w * 4} ${w * 2}` : d === 'dot' ? `${w} ${w}` : undefined
})
const paddingCss = computed(() => (Array.isArray(props.padding) ? props.padding.map((v) => `${v}px`).join(' ') : `${props.padding}px`))

const W = computed(() => props.w ?? size.value.w)
const H = computed(() => props.h ?? size.value.h)

const geometry = computed((): PresetGeometry => {
  let adj = props.adj
  // roundRect の radius（px）は adj に換算する（定義: 角の半径 = ss × adj / 100000）。adj.adj があればそちらが勝つ
  if (props.type === 'roundRect' && !(adj && Object.hasOwn(adj, 'adj'))) {
    const ss = Math.min(W.value, H.value)
    if (ss > 0) adj = { ...adj, adj: (props.radius / ss) * 100000 }
  }
  // 未知の名前や評価の失敗で部品ごと落とさず、rect で描く
  try {
    if (isPreset(props.type)) return evalPreset(props.type, W.value, H.value, adj)
  } catch (e) {
    console.warn(`[PptShape] type "${props.type}" を評価できないので rect で描く`, e)
  }
  return evalPreset('rect', W.value, H.value)
})

/** path の fill の陰影: 塗りの上に黒か白を重ねる */
const SHADE: Partial<Record<PresetFill, { color: string; opacity: number }>> = {
  darken: { color: '#000000', opacity: 0.4 },
  darkenLess: { color: '#000000', opacity: 0.2 },
  lighten: { color: '#ffffff', opacity: 0.4 },
  lightenLess: { color: '#ffffff', opacity: 0.2 },
}

// ---------------------------------------------------------------- 矢じり
const MARKER_KINDS = ['triangle', 'oval', 'diamond'] as const
type MarkerKind = (typeof MARKER_KINDS)[number]
/** head / tail の値 → marker の形。arrow と triangle（と未知の値）は三角 */
const markerKind = (v: string | undefined): MarkerKind | undefined => (!v || v === 'none' ? undefined : v === 'oval' || v === 'diamond' ? v : 'triangle')
const markerId = (k: MarkerKind) => `ppt-arrow-${uid}-${k}`
const headKind = computed(() => markerKind(lineProps.value?.head))
const tailKind = computed(() => markerKind(lineProps.value?.tail))
const markers = computed(() => MARKER_KINDS.filter((k) => k === headKind.value || k === tailKind.value))
const markerStart = computed(() => (headKind.value ? `url(#${markerId(headKind.value)})` : undefined))
const markerEnd = computed(() => (tailKind.value ? `url(#${markerId(tailKind.value)})` : undefined))

const layers = computed(() =>
  geometry.value.paths.map((p, i) => ({
    key: i,
    d: p.d,
    fill: p.fill === 'none' ? 'none' : fillColor.value,
    shade: p.fill !== 'none' && fillColor.value !== 'none' ? SHADE[p.fill] : undefined,
    stroke: p.stroke && !!lineProps.value,
    // 矢じりは開いた path（Z で閉じない）の始点・終点に付ける
    open: !/Z\s*$/.test(p.d),
  })),
)
/** 反転は SVG の中で枠の中心に対して */
const flipTransform = computed(() => {
  if (!props.flipH && !props.flipV) return undefined
  return `translate(${props.flipH ? W.value : 0} ${props.flipV ? H.value : 0}) scale(${props.flipH ? -1 : 1} ${props.flipV ? -1 : 1})`
})

const style = computed(() => {
  // flipV は文字を 180 度回すので、上下の寄せも入れ替わって見える（PowerPoint と同じ）
  const valign = props.flipV ? (props.valign === 'top' ? 'bottom' : props.valign === 'bottom' ? 'top' : 'middle') : props.valign
  return {
    position: positioned.value ? 'absolute' : 'relative',
    left: props.x !== undefined ? `${props.x}px` : undefined,
    top: props.y !== undefined ? `${props.y}px` : undefined,
    width: props.w !== undefined ? `${props.w}px` : undefined,
    height: props.h !== undefined ? `${props.h}px` : props.type === 'line' ? '0px' : undefined,
    minHeight: props.h === undefined && props.type !== 'line' ? '2em' : undefined,
    transform: props.rotate ? `rotate(${props.rotate}deg)` : undefined,
    textAlign: props.align,
    fontSize: props.size !== undefined ? `${props.size}px` : undefined,
    color: props.color,
    display: 'flex',
    flexDirection: 'column',
    justifyContent: valign === 'middle' ? 'center' : valign === 'bottom' ? 'flex-end' : 'flex-start',
    boxSizing: 'border-box',
    overflow: 'visible',
  } as Record<string, string | undefined>
})

const textStyle = computed(() => {
  const s: Record<string, string | undefined> = { padding: paddingCss.value }
  if (props.flipV) s.transform = 'rotate(180deg)'
  if (props.h !== undefined) {
    // 文字の枠（反転した形の上での位置）を margin で空ける。上下の寄せは親の flex が枠の中で行う
    const { l, t, r, b } = geometry.value.textRect
    const left = props.flipH ? W.value - r : l
    const right = props.flipH ? l : W.value - r
    const top = props.flipV ? H.value - b : t
    const bottom = props.flipV ? t : H.value - b
    s.margin = [top, right, bottom, left].map((v) => `${Math.round(v * 100) / 100}px`).join(' ')
  }
  return s
})

const opts = computed(() => {
  const o: Record<string, unknown> = {
    type: props.type,
    fill: props.fill,
    line: props.line,
    rotate: props.rotate,
    padding: props.padding,
    valign: props.valign,
    align: props.align,
  }
  if (props.type === 'roundRect') o.radius = props.radius
  if (props.adj) o.adj = props.adj
  if (props.flipH) o.flipH = true
  if (props.flipV) o.flipV = true
  return JSON.stringify(o)
})
</script>

<template>
  <div
    ref="root"
    data-ppt="shape"
    :data-ppt-export="props.export"
    :data-ppt-name="props.name"
    :data-ppt-box="JSON.stringify(box)"
    :data-ppt-opts="opts"
    :style="style"
    class="ppt-shape"
  >
    <svg
      v-if="props.type !== 'line'"
      class="ppt-shape-svg"
      :width="W"
      :height="H"
      :viewBox="`0 0 ${W} ${H}`"
      aria-hidden="true"
    >
      <defs v-if="markers.length">
        <marker v-for="k in markers" :id="markerId(k)" :key="k" viewBox="0 0 10 10" :refX="k === 'triangle' ? 9 : 5" refY="5" markerWidth="8" markerHeight="8" orient="auto-start-reverse">
          <circle v-if="k === 'oval'" cx="5" cy="5" r="5" :fill="lineProps?.color ?? '#000000'" />
          <path v-else :d="k === 'diamond' ? 'M0,5 L5,0 L10,5 L5,10 z' : 'M0,0 L10,5 L0,10 z'" :fill="lineProps?.color ?? '#000000'" />
        </marker>
      </defs>
      <!-- 塗りを全部描いてから線を描く（chartPlus などは線の path が塗りの path より先に定義されている） -->
      <g :transform="flipTransform">
        <template v-for="p in layers" :key="`f${p.key}`">
          <path v-if="p.fill !== 'none'" :d="p.d" :fill="p.fill" stroke="none" />
          <path v-if="p.shade" :d="p.d" :fill="p.shade.color" :fill-opacity="p.shade.opacity" stroke="none" />
        </template>
        <template v-for="p in layers" :key="`s${p.key}`">
          <path
            v-if="p.stroke"
            :d="p.d"
            fill="none"
            :stroke="lineProps?.color ?? '#000000'"
            :stroke-width="lineProps?.width ?? 1"
            :stroke-dasharray="dashArray"
            stroke-linejoin="round"
            :marker-start="p.open ? markerStart : undefined"
            :marker-end="p.open ? markerEnd : undefined"
          />
        </template>
      </g>
    </svg>
    <svg v-else class="ppt-shape-line" aria-hidden="true">
      <defs v-if="markers.length">
        <marker v-for="k in markers" :id="markerId(k)" :key="k" viewBox="0 0 10 10" :refX="k === 'triangle' ? 9 : 5" refY="5" markerWidth="8" markerHeight="8" orient="auto-start-reverse">
          <circle v-if="k === 'oval'" cx="5" cy="5" r="5" :fill="lineProps?.color ?? '#000000'" />
          <path v-else :d="k === 'diamond' ? 'M0,5 L5,0 L10,5 L5,10 z' : 'M0,0 L10,5 L0,10 z'" :fill="lineProps?.color ?? '#000000'" />
        </marker>
      </defs>
      <line
        x1="0" y1="0" x2="100%" :y2="props.h ? '100%' : '0'"
        :stroke="lineProps?.color ?? '#000000'"
        :stroke-width="lineProps?.width ?? 1"
        :stroke-dasharray="dashArray"
        :marker-start="markerStart"
        :marker-end="markerEnd"
      />
    </svg>
    <div v-if="props.type !== 'line'" class="ppt-shape-text" :style="textStyle"><slot /></div>
  </div>
</template>

<style scoped>
.ppt-shape-svg {
  position: absolute;
  left: 0;
  top: 0;
  overflow: visible; /* 輪郭線の外半分と矢じりを切らない */
}
.ppt-shape-line {
  position: absolute;
  left: 0;
  top: 0;
  width: 100%;
  height: 100%;
  min-height: 2px;
  overflow: visible;
}
.ppt-shape-text {
  position: relative;
}
.ppt-shape-text :deep(> :first-child) {
  margin-top: 0;
}
.ppt-shape-text :deep(> :last-child) {
  margin-bottom: 0;
}
</style>
