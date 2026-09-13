// jszip の薄い包み（native-export.md §3.1）。XML は @xmldom/xmldom の DOM で触る。
import JSZip from 'jszip'
import { DOMParser, XMLSerializer } from '@xmldom/xmldom'

export const NS = {
  p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
  ct: 'http://schemas.openxmlformats.org/package/2006/content-types',
} as const

export const REL_TYPE_BASE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'

export interface ZipView {
  list(): string[]
  has(path: string): boolean
  readXml(path: string): Promise<Document>
  writeXml(path: string, doc: Document): void
  readText(path: string): Promise<string>
  writeText(path: string, text: string): void
  readBinary(path: string): Promise<Uint8Array>
  writeBinary(path: string, data: Uint8Array): void
  remove(path: string): void
}

export function parseXml(xml: string): Document {
  return new DOMParser().parseFromString(xml, 'application/xml') as unknown as Document
}

export function serializeXml(doc: Document): string {
  return new XMLSerializer().serializeToString(doc as never)
}

export async function openZip(buf: Buffer | Uint8Array): Promise<{ zip: JSZip; view: ZipView }> {
  const zip = await JSZip.loadAsync(buf)
  const get = (path: string) => {
    const f = zip.file(path)
    if (!f) throw new Error(`no part: ${path}`)
    return f
  }
  const view: ZipView = {
    list: () => Object.keys(zip.files).filter((f) => !zip.files[f].dir),
    has: (path) => !!zip.file(path),
    readXml: async (path) => parseXml(await get(path).async('string')),
    writeXml: (path, doc) => {
      zip.file(path, serializeXml(doc))
    },
    readText: (path) => get(path).async('string'),
    writeText: (path, text) => {
      zip.file(path, text)
    },
    readBinary: (path) => get(path).async('uint8array'),
    writeBinary: (path, data) => {
      zip.file(path, data)
    },
    remove: (path) => {
      zip.remove(path)
    },
  }
  return { zip, view }
}

export async function saveZip(zip: JSZip): Promise<Buffer> {
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

/** 名前空間つきで子孫要素を集める */
export function elements(node: Document | Element, ns: keyof typeof NS, local: string): Element[] {
  return Array.from(node.getElementsByTagNameNS(NS[ns], local)) as Element[]
}

/** 直下の子要素だけ */
export function children(node: Element, local?: string): Element[] {
  return Array.from(node.childNodes).filter((n): n is Element => n.nodeType === 1 && (!local || (n as Element).localName === local))
}

/** rels を Type で引く（rId は決め打ちにしない）。短い名前（'slideLayout'）でも完全な URI でも可 */
export function findRelByType(relsDoc: Document, type: string): Element | undefined {
  const full = type.includes('/') ? type : REL_TYPE_BASE + type
  return elements(relsDoc, 'rel', 'Relationship').find((r) => r.getAttribute('Type') === full)
}

export function findRelsByType(relsDoc: Document, type: string): Element[] {
  const full = type.includes('/') ? type : REL_TYPE_BASE + type
  return elements(relsDoc, 'rel', 'Relationship').filter((r) => r.getAttribute('Type') === full)
}

/** パートのパス → その rels のパス（ppt/slides/slide1.xml → ppt/slides/_rels/slide1.xml.rels） */
export function relsPathOf(partPath: string): string {
  const i = partPath.lastIndexOf('/')
  const dir = i >= 0 ? partPath.slice(0, i) : ''
  const base = i >= 0 ? partPath.slice(i + 1) : partPath
  return `${dir ? dir + '/' : ''}_rels/${base}.rels`
}

/** rels の Target（相対）をパートの絶対パス（ZIP 内、先頭の / 無し）に解く */
export function resolveTarget(fromPart: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1)
  const dir = fromPart.includes('/') ? fromPart.slice(0, fromPart.lastIndexOf('/')).split('/') : []
  const parts = target.split('/')
  for (const p of parts) {
    if (p === '..') dir.pop()
    else if (p !== '.') dir.push(p)
  }
  return dir.join('/')
}
