<script setup lang="ts">
// PPT 部品: 表（wip/design/ppt-components.md §1.5）。slot の Markdown 表か、`rows` から描く。両方あれば slot を優先。
import { computed, useSlots } from 'vue'

interface BorderProps {
  color?: string
  width?: number
}

const props = withDefaults(
  defineProps<{
    x?: number
    y?: number
    w?: number
    h?: number
    export?: 'native' | 'image'
    name?: string
    rows?: string[][]
    header?: boolean
    colW?: number[]
    border?: string | BorderProps
    fill?: string
    size?: number
    align?: 'left' | 'center' | 'right' | 'justify'
    valign?: 'top' | 'middle' | 'bottom'
  }>(),
  { export: 'native', name: '', header: true },
)

const slots = useSlots()
const useRows = computed(() => !slots.default && !!props.rows?.length)

const box = computed(() => {
  const b: Record<string, number> = {}
  if (props.x !== undefined) b.x = props.x
  if (props.y !== undefined) b.y = props.y
  if (props.w !== undefined) b.w = props.w
  if (props.h !== undefined) b.h = props.h
  return b
})
const positioned = computed(() => Object.keys(box.value).length > 0)

const borderCss = computed(() => {
  if (!props.border) return undefined
  const b: BorderProps = typeof props.border === 'string' ? { color: props.border } : props.border
  return `${b.width ?? 1}px solid ${b.color ?? '#000000'}`
})

const style = computed(() => ({
  position: positioned.value ? 'absolute' : undefined,
  left: props.x !== undefined ? `${props.x}px` : undefined,
  top: props.y !== undefined ? `${props.y}px` : undefined,
  width: props.w !== undefined ? `${props.w}px` : positioned.value ? 'max-content' : undefined,
  height: props.h !== undefined ? `${props.h}px` : undefined,
  fontSize: props.size !== undefined ? `${props.size}px` : undefined,
  textAlign: props.align,
}))
const cellStyle = computed(() => ({
  border: borderCss.value,
  background: props.fill && props.fill !== 'none' ? props.fill : undefined,
  verticalAlign: props.valign,
}))

const opts = computed(() => JSON.stringify({ header: props.header, colW: props.colW, border: props.border, fill: props.fill, valign: props.valign, align: props.align }))
</script>

<template>
  <div
    data-ppt="table"
    :data-ppt-export="props.export"
    :data-ppt-name="props.name"
    :data-ppt-box="JSON.stringify(box)"
    :data-ppt-opts="opts"
    :style="style"
    class="ppt-table"
  >
    <table v-if="useRows" :style="{ tableLayout: props.colW ? 'fixed' : undefined }">
      <colgroup v-if="props.colW">
        <col v-for="(w, i) in props.colW" :key="i" :style="{ width: `${w}px` }">
      </colgroup>
      <thead v-if="props.header && props.rows!.length">
        <tr>
          <th v-for="(c, i) in props.rows![0]" :key="i" :style="cellStyle">{{ c }}</th>
        </tr>
      </thead>
      <tbody>
        <tr v-for="(r, ri) in props.rows!.slice(props.header ? 1 : 0)" :key="ri">
          <td v-for="(c, ci) in r" :key="ci" :style="cellStyle">{{ c }}</td>
        </tr>
      </tbody>
    </table>
    <slot v-else />
  </div>
</template>

<style scoped>
.ppt-table :deep(table) {
  width: 100%;
}
</style>
