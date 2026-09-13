// 後処理の層（native-export.md §3）。パス指定の XML 変換の列を jszip で開いた ZIP に順に当て、DEFLATE で再圧縮する。
import type { Capture, Report } from '../types.ts'
import { openZip, saveZip } from './zip.ts'
import type { ZipView } from './zip.ts'
import { renameShapes } from './patches/renameShapes.ts'
import { dropEmptyPlaceholders } from './patches/dropEmptyPlaceholders.ts'
import { dedupeParagraphProps } from './patches/dedupeParagraphProps.ts'
import { applyAutofitScale } from './patches/applyAutofitScale.ts'
import { splitNotesParagraphs } from './patches/splitNotesParagraphs.ts'
import { replaceMaster } from './patches/replaceMaster.ts'
import { rebuildContentTypes } from './patches/rebuildContentTypes.ts'

export type { ZipView } from './zip.ts'

export interface AutofitScale {
  fontScale: number
  lnSpcReduction: number
}

export interface PatchContext {
  capture: Capture
  /** スライド番号 → 生成時に add した順の図形名（§4.5） */
  shapeNames: Record<number, string[]>
  /** スライド番号 → 図形名 → 縮小率（§3.3） */
  autofit: Record<number, Record<string, AutofitScale>>
  report: Report
}

export interface Patch {
  name: string
  run(zip: ZipView, ctx: PatchContext): Promise<void> | void
}

/** パス指定の XML 変換を Patch にする道具 */
export function xmlPatch(name: string, pattern: RegExp, fn: (doc: Document, path: string, ctx: PatchContext) => void): Patch {
  return {
    name,
    async run(zip, ctx) {
      for (const path of zip.list().filter((p) => pattern.test(p))) {
        const doc = await zip.readXml(path)
        fn(doc, path, ctx)
        zip.writeXml(path, doc)
      }
    },
  }
}

/** 変換の列。順番は固定（§3.2） */
export const PATCHES: Patch[] = [
  renameShapes,
  dropEmptyPlaceholders,
  dedupeParagraphProps,
  applyAutofitScale,
  splitNotesParagraphs,
  replaceMaster,
  rebuildContentTypes,
]

export async function postProcess(buf: Buffer | Uint8Array, patches: Patch[], ctx: PatchContext): Promise<Buffer> {
  const { zip, view } = await openZip(buf)
  for (const patch of patches) await patch.run(view, ctx)
  return saveZip(zip)
}
