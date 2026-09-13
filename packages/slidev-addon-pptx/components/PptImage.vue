<script setup lang="ts">
// PPT 部品: 画像（wip/design/ppt-components.md §1.4）。`fit` は CSS の object-fit にそのまま。PPTX 側の写しは Node が決める。
import { computed, ref } from 'vue'
import { usePptDrag } from '../composables/usePptDrag'

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
    src: string
    alt?: string
    fit?: 'contain' | 'cover' | 'fill'
  }>(),
  { export: 'native', name: '', alt: '', fit: 'contain' },
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

const style = computed(() => ({
  position: positioned.value ? 'absolute' : undefined,
  left: box.value.x !== undefined ? `${box.value.x}px` : undefined,
  top: box.value.y !== undefined ? `${box.value.y}px` : undefined,
  width: box.value.w !== undefined ? `${box.value.w}px` : undefined,
  height: box.value.h !== undefined ? `${box.value.h}px` : undefined,
  objectFit: props.fit,
  display: 'block',
  maxWidth: 'none',
}))

const opts = computed(() => JSON.stringify({ fit: props.fit, alt: props.alt }))
</script>

<template>
  <img
    ref="root"
    data-ppt="image"
    :data-ppt-export="props.export"
    :data-ppt-name="props.name"
    :data-ppt-box="JSON.stringify(box)"
    :data-ppt-opts="opts"
    :src="props.src"
    :alt="props.alt"
    :style="style"
    class="ppt-image"
    draggable="false"
    @dblclick="dragged.onDblclick"
  >
</template>
