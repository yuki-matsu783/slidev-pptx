// 後処理 7: rebuildContentTypes（native-export.md §3.2、C2 / C3）
import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { rebuildContentTypes } from '../../packages/slidev-addon-pptx/src/patch/patches/rebuildContentTypes'
import { openPptx, els, readFixturePptx } from '../helpers/pptx'
import { runPatches } from './helpers'

const overrides = (doc: Document) => els(doc, 'ct', 'Override').map((o) => o.getAttribute('PartName')!)
const defaults = (doc: Document) => Object.fromEntries(els(doc, 'ct', 'Default').map((d) => [d.getAttribute('Extension')!, d.getAttribute('ContentType')!]))

describe('patch/rebuildContentTypes', () => {
  it('fixture には実在しない slideMaster2 の Override がある（前提）', async () => {
    const p = await openPptx(readFixturePptx())
    expect(overrides(await p.xml('[Content_Types].xml'))).toContain('/ppt/slideMasters/slideMaster2.xml')
  })

  it('Override は ZIP に実在するパートだけになる', async () => {
    const p = await openPptx(await runPatches([rebuildContentTypes]))
    const ct = await p.xml('[Content_Types].xml')
    const parts = new Set(p.list().map((x) => '/' + x))
    for (const o of overrides(ct)) expect(parts.has(o)).toBe(true)
    expect(overrides(ct)).not.toContain('/ppt/slideMasters/slideMaster2.xml')
    expect(overrides(ct)).toContain('/ppt/slideMasters/slideMaster1.xml')
  })

  it('全パートが Default か Override で型を持つ', async () => {
    const p = await openPptx(await runPatches([rebuildContentTypes]))
    const ct = await p.xml('[Content_Types].xml')
    const ov = new Set(overrides(ct))
    const df = defaults(ct)
    for (const path of p.list().filter((x) => x !== '[Content_Types].xml')) {
      const ext = path.split('.').pop()!
      expect(ov.has('/' + path) || ext in df, path).toBe(true)
    }
  })

  it('Default は ZIP に実在する拡張子から作る（固定リストを持たない）', async () => {
    const zip = await JSZip.loadAsync(readFixturePptx())
    zip.file('ppt/media/extra.webp', new Uint8Array([0, 1]))
    zip.file('ppt/media/extra.jpg', new Uint8Array([0, 1]))
    const input = await zip.generateAsync({ type: 'nodebuffer' })
    const p = await openPptx(await runPatches([rebuildContentTypes], undefined, input))
    const df = defaults(await p.xml('[Content_Types].xml'))
    expect(df.webp).toBe('image/webp')
    expect(df.jpg).toBe('image/jpeg')
    expect(df.png).toBe('image/png')
    expect(df.rels).toBe('application/vnd.openxmlformats-package.relationships+xml')
    expect(df.xml).toBe('application/xml')
  })

  it('パートの種類ごとの ContentType が正しい', async () => {
    const p = await openPptx(await runPatches([rebuildContentTypes]))
    const ct = await p.xml('[Content_Types].xml')
    const byPart = Object.fromEntries(els(ct, 'ct', 'Override').map((o) => [o.getAttribute('PartName')!, o.getAttribute('ContentType')!]))
    expect(byPart['/ppt/presentation.xml']).toBe('application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml')
    expect(byPart['/ppt/slides/slide1.xml']).toBe('application/vnd.openxmlformats-officedocument.presentationml.slide+xml')
    expect(byPart['/ppt/slideLayouts/slideLayout1.xml']).toBe('application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml')
    expect(byPart['/ppt/notesSlides/notesSlide1.xml']).toBe('application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml')
    expect(byPart['/ppt/notesMasters/notesMaster1.xml']).toBe('application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml')
    expect(byPart['/ppt/theme/theme1.xml']).toBe('application/vnd.openxmlformats-officedocument.theme+xml')
    expect(byPart['/docProps/core.xml']).toBe('application/vnd.openxmlformats-package.core-properties+xml')
    expect(byPart['/docProps/app.xml']).toBe('application/vnd.openxmlformats-officedocument.extended-properties+xml')
  })
})
