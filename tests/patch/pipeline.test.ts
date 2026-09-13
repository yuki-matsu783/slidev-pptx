// 後処理の列全体（native-export.md §3.1、§3.2）: 順番、何もしない Patch の往復、再圧縮、全部通したあとの OPC 検査
import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { PATCHES, postProcess, xmlPatch } from '../../packages/slidev-addon-pptx/src/patch/index'
import type { Patch } from '../../packages/slidev-addon-pptx/src/patch/index'
import { check } from '../../packages/slidev-addon-pptx/src/opc/check'
import { openPptx, readFixturePptx } from '../helpers/pptx'
import { contextFor } from './helpers'

describe('patch/index: 列の定義', () => {
  it('順番は 1〜7 で固定', () => {
    expect(PATCHES.map((p) => p.name)).toEqual([
      'renameShapes',
      'dropEmptyPlaceholders',
      'dedupeParagraphProps',
      'applyAutofitScale',
      'splitNotesParagraphs',
      'replaceMaster',
      'rebuildContentTypes',
    ])
  })
})

describe('patch/index: postProcess', () => {
  it('何もしない Patch を通した結果は、全 XML が元と同値（§11 の 1: xmldom の往復）', async () => {
    const noop: Patch = xmlPatch('noop', /\.(xml|rels)$/, () => {})
    const before = await openPptx(readFixturePptx())
    const after = await openPptx(await postProcess(readFixturePptx(), [noop], contextFor()))
    expect(after.list().sort()).toEqual(before.list().sort())
    for (const path of before.list().filter((p) => /\.(xml|rels)$/.test(p))) {
      const a = normalize(await after.text(path))
      const b = normalize(await before.text(path))
      expect(a, path).toBe(b)
    }
  })

  it('バイナリ（media）は byte 単位で同じ', async () => {
    const before = await JSZip.loadAsync(readFixturePptx())
    const after = await JSZip.loadAsync(await postProcess(readFixturePptx(), [], contextFor()))
    const media = Object.keys(before.files).filter((f) => f.startsWith('ppt/media/'))
    expect(media.length).toBeGreaterThan(0)
    for (const m of media) {
      expect(Buffer.from(await after.file(m)!.async('uint8array'))).toEqual(Buffer.from(await before.file(m)!.async('uint8array')))
    }
  })

  it('再圧縮は DEFLATE（PptxGenJS の出力は STORE なので、何も変えなくても小さくなる）', async () => {
    const input = readFixturePptx()
    const out = await postProcess(input, [], contextFor())
    expect(out.length).toBeLessThan(input.length * 0.6)
    // ZIP のローカルヘッダの圧縮方式（offset 8 から 2 byte）: 0 = STORE, 8 = DEFLATE
    expect(input.readUInt16LE(8)).toBe(0)
    expect(out.readUInt16LE(8)).toBe(8)
  })

  it('全部通したあとは OPC 検査の error が 0', async () => {
    const out = await postProcess(readFixturePptx(), PATCHES, contextFor())
    const results = await check(out)
    expect(results.filter((r) => r.level === 'error')).toEqual([])
  })

  it('Patch は列の順に走る（ctx に記録される順で確かめる）', async () => {
    const order: string[] = []
    const a: Patch = { name: 'a', run: () => { order.push('a') } }
    const b: Patch = { name: 'b', run: () => { order.push('b') } }
    await postProcess(readFixturePptx(), [a, b], contextFor())
    expect(order).toEqual(['a', 'b'])
  })
})

/** 属性の順や空要素の書き方の違いを吸収せず、そのまま比べる（往復は同値であるべき）。改行だけ揃える */
function normalize(xml: string): string {
  return xml.replace(/\r\n/g, '\n').trim()
}
