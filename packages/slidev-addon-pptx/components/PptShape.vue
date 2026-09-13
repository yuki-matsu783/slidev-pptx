<script setup lang="ts">
// PPT 部品: 図形（wip/design/ppt-components.md §1.3）。
// rect / roundRect は div の CSS で、それ以外は inline SVG で描く（形は PowerPoint の preset の近似）。文字は上に重ねる。
// type="line" は対角線。矢じりの marker は部品ごとに一意な id を持つ（/print は全スライドを同時に描く）。
import { computed, onMounted, ref, useId } from 'vue'

type Dash = 'solid' | 'dash' | 'dot'
type Arrow = 'none' | 'arrow' | 'triangle' | 'oval' | 'diamond'
interface LineProps {
  color?: string
  width?: number
  dash?: Dash
  head?: Arrow
  tail?: Arrow
}
type ShapeType = 'rect' | 'roundRect' | 'ellipse' | 'line' | 'rightArrow' | 'leftArrow' | 'upArrow' | 'downArrow' | 'diamond' | 'triangle' | 'hexagon'

const props = withDefaults(
  defineProps<{
    x?: number
    y?: number
    w?: number
    h?: number
    export?: 'native' | 'image'
    name?: string
    type?: ShapeType
    fill?: string
    line?: string | LineProps
    radius?: number
    rotate?: number
    /** 図形の中の文字の内側余白 px。PPTX の inset にも写す */
    padding?: number | [number, number, number, number]
    align?: 'left' | 'center' | 'right' | 'justify'
    valign?: 'top' | 'middle' | 'bottom'
    size?: number
    color?: string
  }>(),
  { export: 'native', name: '', type: 'rect', fill: '#ffffff', line: '#000000', radius: 8, rotate: 0, padding: () => [0, 8, 0, 8], align: 'center', valign: 'middle' },
)

const uid = useId()
const root = ref<HTMLElement | null>(null)
// rotate のときは回転前の枠（offset*）を測って data-ppt-box に載せる（回転後の外接矩形を使わない）
const measured = ref<Record<string, number>>({})
onMounted(() => {
  if (props.rotate && root.value) {
    const el = root.value
    measured.value = { x: el.offsetLeft, y: el.offsetTop, w: el.offsetWidth, h: el.offsetHeight }
  }
})

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
const isCssShape = computed(() => props.type === 'rect' || props.type === 'roundRect')
const dashArray = computed(() => {
  const d = lineProps.value?.dash
  const w = lineProps.value?.width ?? 1
  return d === 'dash' ? `${w * 4} ${w * 2}` : d === 'dot' ? `${w} ${w}` : undefined
})
const paddingCss = computed(() => (Array.isArray(props.padding) ? props.padding.map((v) => `${v}px`).join(' ') : `${props.padding}px`))

const style = computed(() => {
  const s: Record<string, string | undefined> = {
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
    justifyContent: props.valign === 'middle' ? 'center' : props.valign === 'bottom' ? 'flex-end' : 'flex-start',
    boxSizing: 'border-box',
    overflow: 'visible',
  }
  if (isCssShape.value) {
    s.background = fillColor.value === 'none' ? undefined : fillColor.value
    if (lineProps.value) {
      const l = lineProps.value
      s.border = `${l.width ?? 1}px ${l.dash === 'dash' ? 'dashed' : l.dash === 'dot' ? 'dotted' : 'solid'} ${l.color ?? '#000000'}`
    }
    if (props.type === 'roundRect') s.borderRadius = `${props.radius}px`
  }
  return s
})

/** viewBox 0 0 100 100 の多角形。preserveAspectRatio="none" で枠に合わせる（PowerPoint の preset の既定 adj とは別物。近似） */
const polygon = computed(() => {
  switch (props.type) {
    case 'rightArrow':
      return '0,25 60,25 60,0 100,50 60,100 60,75 0,75'
    case 'leftArrow':
      return '100,25 40,25 40,0 0,50 40,100 40,75 100,75'
    case 'upArrow':
      return '25,100 25,40 0,40 50,0 100,40 75,40 75,100'
    case 'downArrow':
      return '25,0 25,60 0,60 50,100 100,60 75,60 75,0'
    case 'diamond':
      return '50,0 100,50 50,100 0,50'
    case 'triangle':
      return '50,0 100,100 0,100'
    case 'hexagon':
      return '25,0 75,0 100,50 75,100 25,100 0,50'
    default:
      return undefined
  }
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
      v-if="!isCssShape && props.type !== 'line'"
      class="ppt-shape-svg"
      viewBox="0 0 100 100"
      preserveAspectRatio="none"
      aria-hidden="true"
    >
      <ellipse v-if="props.type === 'ellipse'" cx="50" cy="50" rx="50" ry="50" :fill="fillColor" :stroke="lineProps?.color ?? 'none'" :stroke-width="lineProps?.width ?? 0" :stroke-dasharray="dashArray" vector-effect="non-scaling-stroke" />
      <polygon v-else :points="polygon" :fill="fillColor" :stroke="lineProps?.color ?? 'none'" :stroke-width="lineProps?.width ?? 0" :stroke-dasharray="dashArray" stroke-linejoin="round" vector-effect="non-scaling-stroke" />
    </svg>
    <svg v-else-if="props.type === 'line'" class="ppt-shape-line" aria-hidden="true">
      <defs>
        <marker :id="`ppt-arrow-${uid}`" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="8" markerHeight="8" orient="auto-start-reverse">
          <path d="M0,0 L10,5 L0,10 z" :fill="lineProps?.color ?? '#000000'" />
        </marker>
      </defs>
      <line
        x1="0" y1="0" x2="100%" :y2="props.h ? '100%' : '0'"
        :stroke="lineProps?.color ?? '#000000'"
        :stroke-width="lineProps?.width ?? 1"
        :stroke-dasharray="dashArray"
        :marker-start="lineProps?.head && lineProps.head !== 'none' ? `url(#ppt-arrow-${uid})` : undefined"
        :marker-end="lineProps?.tail && lineProps.tail !== 'none' ? `url(#ppt-arrow-${uid})` : undefined"
      />
    </svg>
    <div v-if="props.type !== 'line'" class="ppt-shape-text" :style="{ padding: paddingCss }"><slot /></div>
  </div>
</template>

<style scoped>
.ppt-shape-svg {
  position: absolute;
  inset: 0;
  width: 100%;
  height: 100%;
  overflow: visible; /* 輪郭線の外半分を切らない */
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
