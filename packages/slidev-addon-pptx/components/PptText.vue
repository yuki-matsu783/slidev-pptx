<script setup lang="ts">
// PPT 部品: テキスト枠（wip/design/ppt-components.md §1.2）。
// 指定した props を CSS にも反映し（Slidev の見た目 = PPTX の見た目）、収集器には data-ppt-* 属性で渡す。
import { computed, ref } from 'vue'
import { usePptDrag } from '../composables/usePptDrag'
import PptEditUi from '../internals/PptEditUi.vue'

type Dash = 'solid' | 'dash' | 'dot'
interface LineProps {
  color?: string
  width?: number
  dash?: Dash
}

const props = withDefaults(
  defineProps<{
    x?: number
    y?: number
    w?: number
    h?: number
    /** 試作: ドラッグで動かすときの id。位置はスライドの frontmatter の dragPos[drag] に保存され、x y w h より優先する */
    drag?: string
    export?: 'native' | 'image'
    name?: string
    align?: 'left' | 'center' | 'right' | 'justify'
    valign?: 'top' | 'middle' | 'bottom'
    size?: number
    color?: string
    bold?: boolean
    italic?: boolean
    fill?: string
    line?: string | LineProps
    radius?: number
    padding?: number | [number, number, number, number]
  }>(),
  { export: 'native', name: '', valign: 'top', radius: 0, padding: 0 },
)

const root = ref<HTMLElement | null>(null)
const dragged = usePptDrag(props, root, { rotatable: false })

const box = computed<Record<string, number>>(() => {
  if (dragged.box.value) return { ...dragged.box.value }
  const b: Record<string, number> = {}
  if (props.x !== undefined) b.x = props.x
  if (props.y !== undefined) b.y = props.y
  if (props.w !== undefined) b.w = props.w
  if (props.h !== undefined) b.h = props.h
  return b
})
const positioned = computed(() => Object.keys(box.value).length > 0)
const hasH = computed(() => box.value.h !== undefined)

const lineCss = computed(() => {
  if (!props.line || props.line === 'none') return undefined
  const l: LineProps = typeof props.line === 'string' ? { color: props.line } : props.line
  const style = l.dash === 'dash' ? 'dashed' : l.dash === 'dot' ? 'dotted' : 'solid'
  return `${l.width ?? 1}px ${style} ${l.color ?? '#000000'}`
})
const paddingCss = computed(() => (Array.isArray(props.padding) ? props.padding.map((v) => `${v}px`).join(' ') : `${props.padding}px`))

const style = computed(() => {
  const b = box.value
  const s: Record<string, string | undefined> = {
    position: positioned.value ? 'absolute' : undefined,
    left: b.x !== undefined ? `${b.x}px` : undefined,
    top: b.y !== undefined ? `${b.y}px` : undefined,
    width: b.w !== undefined ? `${b.w}px` : positioned.value ? 'max-content' : undefined,
    maxWidth: positioned.value && b.w === undefined ? `calc(100% - ${b.x ?? 0}px)` : undefined,
    height: hasH.value ? `${b.h}px` : undefined,
    textAlign: props.align,
    fontSize: props.size !== undefined ? `${props.size}px` : undefined,
    color: props.color,
    fontWeight: props.bold ? '600' : undefined,
    fontStyle: props.italic ? 'italic' : undefined,
    background: props.fill && props.fill !== 'none' ? props.fill : undefined,
    border: lineCss.value,
    borderRadius: props.radius ? `${props.radius}px` : undefined,
    padding: paddingCss.value,
    boxSizing: 'border-box',
    display: hasH.value ? 'flex' : undefined,
    flexDirection: hasH.value ? 'column' : undefined,
    justifyContent: hasH.value ? (props.valign === 'middle' ? 'center' : props.valign === 'bottom' ? 'flex-end' : 'flex-start') : undefined,
  }
  return s
})

const opts = computed(() =>
  JSON.stringify({
    // valign は h があるときだけ CSS に効く。効かないときは載せない（Slidev と PPTX の見た目を揃える）
    valign: hasH.value ? props.valign : undefined,
    align: props.align,
    fill: props.fill,
    line: props.line,
    radius: props.radius,
    padding: props.padding,
  }),
)
</script>

<template>
  <div
    ref="root"
    data-ppt="text"
    :data-ppt-export="props.export"
    :data-ppt-name="props.name"
    :data-ppt-box="JSON.stringify(box)"
    :data-ppt-opts="opts"
    :style="style"
    class="ppt-text"
    @dblclick="dragged.onDblclick"
  >
    <slot />
    <PptEditUi v-if="props.drag" tag="PptText" :drag="props.drag" :dragging="dragged.dragging.value" @open="dragged.stop" />
  </div>
</template>

<style scoped>
.ppt-text :deep(> :first-child) {
  margin-top: 0;
}
.ppt-text :deep(> :last-child) {
  margin-bottom: 0;
}
</style>
