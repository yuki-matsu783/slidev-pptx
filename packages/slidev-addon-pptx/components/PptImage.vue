<script setup lang="ts">
// PPT 部品: 画像（wip/design/ppt-components.md §1.4）。`fit` は CSS の object-fit にそのまま。PPTX 側の写しは Node が決める。
import { computed } from 'vue'

const props = withDefaults(
  defineProps<{
    x?: number
    y?: number
    w?: number
    h?: number
    export?: 'native' | 'image'
    name?: string
    src: string
    alt?: string
    fit?: 'contain' | 'cover' | 'fill'
  }>(),
  { export: 'native', name: '', alt: '', fit: 'contain' },
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

const style = computed(() => ({
  position: positioned.value ? 'absolute' : undefined,
  left: props.x !== undefined ? `${props.x}px` : undefined,
  top: props.y !== undefined ? `${props.y}px` : undefined,
  width: props.w !== undefined ? `${props.w}px` : undefined,
  height: props.h !== undefined ? `${props.h}px` : undefined,
  objectFit: props.fit,
  display: 'block',
  maxWidth: 'none',
}))

const opts = computed(() => JSON.stringify({ fit: props.fit, alt: props.alt }))
</script>

<template>
  <img
    data-ppt="image"
    :data-ppt-export="props.export"
    :data-ppt-name="props.name"
    :data-ppt-box="JSON.stringify(box)"
    :data-ppt-opts="opts"
    :src="props.src"
    :alt="props.alt"
    :style="style"
    class="ppt-image"
  >
</template>
