<script setup lang="ts">
// PPT 部品: テキスト枠（wip/design/ppt-components.md §1.2）。
// 指定した props を CSS にも反映し（Slidev の見た目 = PPTX の見た目）、収集器には data-ppt-* 属性で渡す。
import { computed } from 'vue'

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

const box = computed(() => {
  const b: Record<string, number> = {}
  if (props.x !== undefined) b.x = props.x
  if (props.y !== undefined) b.y = props.y
  if (props.w !== undefined) b.w = props.w
  if (props.h !== undefined) b.h = props.h
  return b
})
const positioned = computed(() => Object.keys(box.value).length > 0)

const lineCss = computed(() => {
  if (!props.line || props.line === 'none') return undefined
  const l: LineProps = typeof props.line === 'string' ? { color: props.line } : props.line
  const style = l.dash === 'dash' ? 'dashed' : l.dash === 'dot' ? 'dotted' : 'solid'
  return `${l.width ?? 1}px ${style} ${l.color ?? '#000000'}`
})
const paddingCss = computed(() => (Array.isArray(props.padding) ? props.padding.map((v) => `${v}px`).join(' ') : `${props.padding}px`))

const style = computed(() => {
  const s: Record<string, string | undefined> = {
    position: positioned.value ? 'absolute' : undefined,
    left: props.x !== undefined ? `${props.x}px` : undefined,
    top: props.y !== undefined ? `${props.y}px` : undefined,
    width: props.w !== undefined ? `${props.w}px` : positioned.value ? 'max-content' : undefined,
    maxWidth: positioned.value && props.w === undefined ? `calc(100% - ${props.x ?? 0}px)` : undefined,
    height: props.h !== undefined ? `${props.h}px` : undefined,
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
    display: props.h !== undefined ? 'flex' : undefined,
    flexDirection: props.h !== undefined ? 'column' : undefined,
    justifyContent: props.h !== undefined ? (props.valign === 'middle' ? 'center' : props.valign === 'bottom' ? 'flex-end' : 'flex-start') : undefined,
  }
  return s
})

const opts = computed(() =>
  JSON.stringify({
    valign: props.valign,
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
    data-ppt="text"
    :data-ppt-export="props.export"
    :data-ppt-name="props.name"
    :data-ppt-box="JSON.stringify(box)"
    :data-ppt-opts="opts"
    :style="style"
    class="ppt-text"
  >
    <slot />
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
