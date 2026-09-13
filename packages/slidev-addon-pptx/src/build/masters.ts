// レイアウト対応表とマスターの定義（native-export.md §5.2、§5.3）。マスター定義はこのファイルだけ。
// 対応表は Slidev の px（canvasWidth 980 のとき）で持ち、defineMasters が canvasWidth で EMU に換算する。
import type PptxGenJS from 'pptxgenjs'
import { emu } from './units.ts'

export type LayoutName = 'blank' | 'cover' | 'default' | 'center' | 'two-cols'

export interface PlaceholderDef {
  name: 'title' | 'body' | 'body2'
  type: 'title' | 'body'
  x: number
  y: number
  w: number
  h: number
}

export interface LayoutDef {
  title: LayoutName
  background: { color: string }
  margin: number
  objects: { placeholder: { options: PlaceholderDef } }[]
  /** 付けない（DEFAULT レイアウトに番号 placeholder が増えるため）。型として「無い」ことを明示 */
  slideNumber?: undefined
}

const PX = 56 // .slidev-layout の px-14
const PY = 40 // py-10
const CW = 868 // 内容幅

const ph = (name: PlaceholderDef['name'], type: PlaceholderDef['type'], x: number, y: number, w: number, h: number) => ({
  placeholder: { options: { name, type, x, y, w, h } },
})

/** キーの順が slideLayout の番号を決める（1 は組み込み DEFAULT、2 から順） */
export const LAYOUTS: Record<LayoutName, LayoutDef> = {
  blank: { title: 'blank', background: { color: 'FFFFFF' }, margin: 0, objects: [] },
  cover: {
    title: 'cover',
    background: { color: 'FFFFFF' },
    margin: 0,
    objects: [ph('title', 'title', PX, 220, CW, 80), ph('body', 'body', PX, 308, CW, 110)],
  },
  default: {
    title: 'default',
    background: { color: 'FFFFFF' },
    margin: 0,
    objects: [ph('title', 'title', PX, PY, CW, 44), ph('body', 'body', PX, 96, CW, 416)],
  },
  center: {
    title: 'center',
    background: { color: 'FFFFFF' },
    margin: 0,
    objects: [ph('title', 'title', PX, 236, CW, 44), ph('body', 'body', PX, 292, CW, 110)],
  },
  'two-cols': {
    title: 'two-cols',
    background: { color: 'FFFFFF' },
    margin: 0,
    objects: [ph('title', 'title', PX, PY, 434, 44), ph('body', 'body', PX, 96, 434, 416), ph('body2', 'body', 490, PY, 434, 472)],
  },
}

/**
 * 表の順で defineSlideMaster を呼ぶ。PptxGenJS の createSlideMaster は渡した options を破壊するので、写しを渡す。
 * 座標は px → EMU（幅だけで換算。§4.1）
 */
export function defineMasters(pptx: PptxGenJS, canvasWidth: number): void {
  const canvas = { width: canvasWidth, height: Math.round((canvasWidth * 9) / 16) }
  for (const def of Object.values(LAYOUTS)) {
    pptx.defineSlideMaster({
      title: def.title,
      background: { ...def.background },
      margin: def.margin,
      objects: def.objects.map((o) => ({
        placeholder: {
          options: {
            name: o.placeholder.options.name,
            type: o.placeholder.options.type,
            x: emu(o.placeholder.options.x, canvas),
            y: emu(o.placeholder.options.y, canvas),
            w: emu(o.placeholder.options.w, canvas),
            h: emu(o.placeholder.options.h, canvas),
          },
        },
      })),
    })
  }
}

/** §5.1 の 2 段目: 対応表に無ければ blank */
export function masterFor(layout: string): LayoutName {
  return layout in LAYOUTS ? (layout as LayoutName) : 'blank'
}

/** そのレイアウトに placeholder があるか（§4.2 のガード） */
export function hasPlaceholder(layout: LayoutName, name: PlaceholderDef['name']): boolean {
  return LAYOUTS[layout].objects.some((o) => o.placeholder.options.name === name)
}
