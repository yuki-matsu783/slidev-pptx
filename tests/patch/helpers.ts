// 後処理の検査で共通に使う PatchContext の見本と、1 変換だけを通す道具。
import type { Patch, PatchContext } from '../../packages/slidev-addon-pptx/src/patch/index'
import { postProcess } from '../../packages/slidev-addon-pptx/src/patch/index'
import type { Capture, Report } from '../../packages/slidev-addon-pptx/src/types'
import { readFixturePptx } from '../helpers/pptx'

export function emptyReport(): Report {
  return {
    output: 'x.pptx', generatedAt: '', slidev: '', pptxgenjs: '',
    slides: 0, native: 0, replaced: 0, replacements: [], warnings: [], dropped: {}, zoom: {}, slideMap: {}, check: [],
  }
}

export function contextFor(overrides: Partial<PatchContext> = {}): PatchContext {
  const capture: Capture = { canvas: { width: 980, height: 552 }, slides: [] }
  return {
    capture,
    // fixture の sample.pptx の add 順（tests/fixtures/gen-pptx.mjs）
    shapeNames: {
      1: ['Title', 'Text 2', 'Table 3', 'Image 4'],
      2: ['Text 1', 'Line 2'],
    },
    autofit: {},
    adjust: {},
    report: emptyReport(),
    ...overrides,
  }
}

/** fixture に 1 変換（または指定の列）だけを通す */
export async function runPatches(patches: Patch[], ctx: PatchContext = contextFor(), input: Buffer = readFixturePptx()): Promise<Buffer> {
  return postProcess(input, patches, ctx)
}
