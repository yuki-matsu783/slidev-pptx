// 試作: PPT 部品を Slidev のドラッグ（ダブルクリックで掴む）で動かす。
// Slidev の useDragElement を frontmatter 方式で使い、位置はそのスライドの frontmatter の dragPos[drag] に保存される。
// 動かした枠と回転を部品自身の box / rotate に写すので、data-ppt-box と PPTX にもそのまま出る。
import type { Ref } from 'vue'
import { useDragElement } from '@slidev/client/composables/useDragElements.ts'
import { injectionFrontmatter, injectionSlideScale } from '@slidev/client/constants.ts'
import { computed, inject, onUnmounted, ref, unref, watch } from 'vue'

export interface PptDragProps {
  drag?: string
  x?: number
  y?: number
  w?: number
  h?: number
  rotate?: number
}

export interface PptBox {
  x: number
  y: number
  w: number
  h: number
}

export function usePptDrag(props: PptDragProps, root: Ref<HTMLElement | null>, opts: { rotatable: boolean }) {
  const id = props.drag
  if (!id) {
    return {
      box: computed<PptBox | undefined>(() => undefined),
      rotate: computed<number | undefined>(() => undefined),
      dragging: computed(() => false),
      stop: () => {},
      onDblclick: () => {},
    }
  }

  // dragPos が無く、x y w h が揃っていれば、props を初期位置として frontmatter（メモリ上だけ）に入れる。
  // ファイルに書かれるのは実際に動かしたときだけ（Slidev の書き戻しはサーバから取った frontmatter を使う）
  const frontmatter = inject(injectionFrontmatter, undefined) as Record<string, any> | undefined
  const complete = [props.x, props.y, props.w, props.h].every((v) => typeof v === 'number')
  if (frontmatter && complete && !frontmatter.dragPos?.[id]) {
    frontmatter.dragPos ||= {}
    frontmatter.dragPos[id] = [props.x, props.y, props.w, props.h, props.rotate ?? 0].join()
  }

  const scale = inject(injectionSlideScale, ref(1))
  const d = useDragElement(null, id)
  watch(root, (el) => { d.container.value = el ?? undefined }, { immediate: true })
  onUnmounted(d.unmounted)

  // 回転を PPTX に写せるのは PptShape だけ。ほかは回転の操作を 0 に戻す
  if (!opts.rotatable)
    watch(d.rotate, (r) => { if (r) d.rotate.value = 0 })

  const box = computed<PptBox | undefined>(() => {
    const w = d.width.value
    const h = d.height.value
    if (!Number.isFinite(d.x0.value) || !Number.isFinite(w) || !Number.isFinite(h)) return undefined
    return { x: Math.round(d.x0.value - w / 2), y: Math.round(d.y0.value - h / 2), w: Math.round(w), h: Math.round(h) }
  })
  const rotate = computed(() => (box.value && opts.rotatable ? Math.round(d.rotate.value) : undefined))

  const onDblclick = () => {
    // 位置を持っていない部品（x y を書いていない実測配置）は、掴んだ時点の実測を初期値にする
    if (!box.value && root.value) {
      const el = root.value
      const r = el.getBoundingClientRect()
      const page = el.closest('.slidev-page')?.getBoundingClientRect() ?? { left: 0, top: 0 }
      const s = unref(scale) || 1
      const w = r.width / s
      const h = r.height / s
      d.width.value = w
      d.height.value = h
      d.x0.value = (r.left - page.left) / s + w / 2
      d.y0.value = (r.top - page.top) / s + h / 2
    }
    d.startDragging()
  }

  return { box, rotate, dragging: d.dragging, stop: () => d.stopDragging(), onDblclick }
}
