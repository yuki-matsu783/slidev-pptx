<script setup lang="ts">
// PPT 部品: 表（wip/design/ppt-components.md §1.5）。slot の Markdown 表か、`rows` から描く。両方あれば slot を優先。
// slot の表にも colW / border / fill / valign を効かせる（colW は描画後に 1 行目のセルへ幅を当てる）。
import { computed, onMounted, ref, useSlots, watchEffect } from 'vue'
import { usePptDrag } from '../composables/usePptDrag'

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
    /** 試作: ドラッグで動かすときの id。位置はスライドの frontmatter の dragPos[drag] に保存され、x y w h より優先する */
    drag?: string
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

const borderCss = computed(() => {
  if (!props.border) return undefined
  const b: BorderProps = typeof props.border === 'string' ? { color: props.border } : props.border
  return `${b.width ?? 1}px solid ${b.color ?? '#000000'}`
})

const style = computed(() => ({
  position: positioned.value ? 'absolute' : undefined,
  left: box.value.x !== undefined ? `${box.value.x}px` : undefined,
  top: box.value.y !== undefined ? `${box.value.y}px` : undefined,
  width: box.value.w !== undefined ? `${box.value.w}px` : positioned.value ? 'max-content' : undefined,
  height: box.value.h !== undefined ? `${box.value.h}px` : undefined,
  fontSize: props.size !== undefined ? `${props.size}px` : undefined,
  textAlign: props.align,
  // slot の表のセルに :deep で当てる
  '--ppt-cell-border': borderCss.value,
  '--ppt-cell-bg': props.fill && props.fill !== 'none' ? props.fill : undefined,
  '--ppt-cell-valign': props.valign,
}))

// colW: slot の表には <colgroup> を差し込めないので、描画後に 1 行目のセルへ幅を当てる
const applyColW = () => {
  const tbl = root.value?.querySelector('table')
  if (!tbl || !props.colW) return
  tbl.style.tableLayout = 'fixed'
  const first = tbl.querySelector('tr')
  if (!first) return
  Array.from(first.children).forEach((c, i) => {
    if (props.colW![i] !== undefined) (c as HTMLElement).style.width = `${props.colW![i]}px`
  })
}
onMounted(() => watchEffect(applyColW))

const opts = computed(() => JSON.stringify({ header: props.header, colW: props.colW, border: props.border, fill: props.fill, valign: props.valign, align: props.align }))
</script>

<template>
  <div
    ref="root"
    data-ppt="table"
    :data-ppt-export="props.export"
    :data-ppt-name="props.name"
    :data-ppt-box="JSON.stringify(box)"
    :data-ppt-opts="opts"
    :style="style"
    class="ppt-table"
    @dblclick="dragged.onDblclick"
  >
    <table v-if="useRows">
      <thead v-if="props.header && props.rows!.length">
        <tr>
          <th v-for="(c, i) in props.rows![0]" :key="i" scope="col">{{ c }}</th>
        </tr>
      </thead>
      <tbody>
        <tr v-for="(r, ri) in props.rows!.slice(props.header ? 1 : 0)" :key="ri">
          <td v-for="(c, ci) in r" :key="ci">{{ c }}</td>
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
/* 変数が無いときは既定値へ（var() の解決失敗で初期値に倒れると vertical-align が baseline になる） */
.ppt-table :deep(th),
.ppt-table :deep(td) {
  border: var(--ppt-cell-border, initial);
  background: var(--ppt-cell-bg, transparent);
  vertical-align: var(--ppt-cell-valign, middle);
}
</style>
