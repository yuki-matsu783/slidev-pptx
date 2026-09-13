// 後処理 7: [Content_Types].xml を ZIP の実在パートから作り直す（native-export.md §3.2）。
// Default は ZIP に実在する拡張子から（固定リストを持たない）、Override はパートの種類ごと。最後に走る。
// 前提にする変換: 6（パートの追加・削除が終わっている）
import type { Patch } from '../index.ts'
import { NS } from '../zip.ts'

const PML = 'application/vnd.openxmlformats-officedocument.presentationml.'
const DEFAULTS_BY_EXT: Record<string, string> = {
  rels: 'application/vnd.openxmlformats-package.relationships+xml',
  xml: 'application/xml',
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  webp: 'image/webp',
  svg: 'image/svg+xml',
  tif: 'image/tiff',
  tiff: 'image/tiff',
  emf: 'image/x-emf',
  wmf: 'image/x-wmf',
  mp4: 'video/mp4',
  mp3: 'audio/mpeg',
  wav: 'audio/wav',
  bin: 'application/vnd.openxmlformats-officedocument.oleObject',
}

const OVERRIDES: [RegExp, string][] = [
  [/^ppt\/presentation\.xml$/, `${PML}presentation.main+xml`],
  [/^ppt\/slideMasters\/slideMaster\d+\.xml$/, `${PML}slideMaster+xml`],
  [/^ppt\/slideLayouts\/slideLayout\d+\.xml$/, `${PML}slideLayout+xml`],
  [/^ppt\/slides\/slide\d+\.xml$/, `${PML}slide+xml`],
  [/^ppt\/notesMasters\/notesMaster\d+\.xml$/, `${PML}notesMaster+xml`],
  [/^ppt\/notesSlides\/notesSlide\d+\.xml$/, `${PML}notesSlide+xml`],
  [/^ppt\/presProps\.xml$/, `${PML}presProps+xml`],
  [/^ppt\/viewProps\.xml$/, `${PML}viewProps+xml`],
  [/^ppt\/tableStyles\.xml$/, `${PML}tableStyles+xml`],
  [/^ppt\/theme\/theme\d+\.xml$/, 'application/vnd.openxmlformats-officedocument.theme+xml'],
  [/^ppt\/charts\/chart\d+\.xml$/, 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'],
  [/^docProps\/core\.xml$/, 'application/vnd.openxmlformats-package.core-properties+xml'],
  [/^docProps\/app\.xml$/, 'application/vnd.openxmlformats-officedocument.extended-properties+xml'],
]

export const rebuildContentTypes: Patch = {
  name: 'rebuildContentTypes',
  async run(zip) {
    const parts = zip.list().filter((p) => p !== '[Content_Types].xml')
    const exts = new Set<string>()
    const overrides: [string, string][] = []
    for (const part of parts) {
      const ov = OVERRIDES.find(([re]) => re.test(part))
      if (ov) {
        overrides.push([part, ov[1]])
        continue
      }
      const ext = part.includes('.') ? part.slice(part.lastIndexOf('.') + 1).toLowerCase() : ''
      if (ext) exts.add(ext)
    }
    exts.add('rels')
    exts.add('xml')
    const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;')
    const lines = [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      `<Types xmlns="${NS.ct}">`,
      ...Array.from(exts)
        .sort()
        .map((ext) => `<Default Extension="${esc(ext)}" ContentType="${esc(DEFAULTS_BY_EXT[ext] ?? 'application/octet-stream')}"/>`),
      ...overrides.map(([part, type]) => `<Override PartName="/${esc(part)}" ContentType="${esc(type)}"/>`),
      '</Types>',
    ]
    zip.writeText('[Content_Types].xml', lines.join(''))
  },
}
