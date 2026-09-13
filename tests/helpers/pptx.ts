// テストから PPTX（ZIP）と XML を読む共通の道具。実装（packages/）には依存しない。
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import { DOMParser } from '@xmldom/xmldom'

const __dirname = fileURLToPath(new URL('.', import.meta.url))

export const NS = {
  p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
  ct: 'http://schemas.openxmlformats.org/package/2006/content-types',
} as const

export const FIXTURE_PPTX = resolve(__dirname, '../fixtures/pptx/sample.pptx')

export function readFixturePptx(): Buffer {
  return readFileSync(FIXTURE_PPTX)
}

export interface OpenedPptx {
  zip: JSZip
  list(): string[]
  text(path: string): Promise<string>
  xml(path: string): Promise<Document>
}

export async function openPptx(buf: Buffer | Uint8Array): Promise<OpenedPptx> {
  const zip = await JSZip.loadAsync(buf)
  return {
    zip,
    list: () => Object.keys(zip.files).filter((f) => !zip.files[f].dir),
    text: async (path) => {
      const f = zip.file(path)
      if (!f) throw new Error(`no part: ${path}`)
      return f.async('string')
    },
    xml: async (path) => {
      const f = zip.file(path)
      if (!f) throw new Error(`no part: ${path}`)
      return new DOMParser().parseFromString(await f.async('string'), 'application/xml') as unknown as Document
    },
  }
}

export function els(node: Document | Element, ns: keyof typeof NS, local: string): Element[] {
  return Array.from(node.getElementsByTagNameNS(NS[ns], local)) as Element[]
}

/** <p:spTree> の直下の図形（sp / pic / graphicFrame / cxnSp / grpSp）を文書順に */
export function shapesOf(slide: Document): Element[] {
  const tree = els(slide, 'p', 'spTree')[0]
  return Array.from(tree.childNodes).filter(
    (n): n is Element => n.nodeType === 1 && ['sp', 'pic', 'graphicFrame', 'cxnSp', 'grpSp'].includes((n as Element).localName),
  )
}

export function cNvPrOf(shape: Element): Element {
  return els(shape, 'p', 'cNvPr')[0]
}

export function shapeNamed(slide: Document, name: string): Element | undefined {
  return shapesOf(slide).find((s) => cNvPrOf(s)?.getAttribute('name') === name)
}

export function textOf(el: Element): string {
  return els(el, 'a', 't')
    .map((t) => t.textContent ?? '')
    .join('')
}
