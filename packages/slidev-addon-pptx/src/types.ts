// 収集結果（Capture）、書き出しの記録（Report）、OPC 整合チェッカの結果（CheckResult）の型。
// ブラウザ側（collect）と Node 側（build / patch / opc / export）の両方が import する唯一の共有物。
// 設計: wip/design/native-export.md §2、§3.4、§8.1

export interface Box {
  x: number
  y: number
  w: number
  h: number
}

export interface Canvas {
  width: number
  height: number
}

export interface Capture {
  canvas: Canvas
  slides: SlideCapture[]
}

export interface SlideCapture {
  no: number
  lang: string
  zoom: number
  backgroundColor: string
  elements: Element[]
  warnings: Warning[]
  /** 規則として捨てたものの件数（'code-highlight' 'blockquote-border' など）。Report.dropped に合算する */
  dropped?: Record<string, number>
}

export type Element = TextElement | ShapeElement | ImageElement | TableElement | LineElement

export type ElementKind = Element['kind']

export type RoleHint = 'title' | 'body' | 'body2'

export type Link = { url: string } | { slide: number }

export interface ElementBase {
  id: string
  name: string
  source: 'markdown' | 'ppt' | 'replaced'
  box: Box
  boxSource: 'prop' | 'measured'
  link?: Link
}

export interface TextElement extends ElementBase {
  kind: 'text'
  roleHint?: RoleHint
  paragraphs: Paragraph[]
  frame: FrameStyle
  fit: { contentHeight: number; boxHeight: number }
  valign: 'top' | 'middle' | 'bottom'
  rotate?: number
}

export type ParagraphKind = 'plain' | 'heading' | 'bullet' | 'number' | 'code'
export type Align = 'left' | 'center' | 'right' | 'justify'

export interface Paragraph {
  kind: ParagraphKind
  level: number
  numberStart?: number
  align: Align
  lineHeight: number
  spaceBefore: number
  spaceAfter: number
  runs: Run[]
}

export interface Run {
  text: string
  size: number
  color: string
  transparency?: number
  bold: boolean
  italic: boolean
  underline: boolean
  strike: boolean
  code: boolean
  highlight?: string
  sup?: boolean
  sub?: boolean
  link?: Link
  charSpacing?: number
  breakAfter?: boolean
}

export type Dash = 'solid' | 'dash' | 'dot'

export interface FrameStyle {
  fill?: { color: string; transparency?: number }
  line?: { color: string; width: number; dash: Dash }
  radius?: number
  inset: [number, number, number, number]
  transparency?: number
}

export interface ShapeElement extends ElementBase {
  kind: 'shape'
  shape: string
  frame: FrameStyle
  rotate?: number
  paragraphs?: Paragraph[]
  valign?: 'top' | 'middle' | 'bottom'
  /** 調整値。キーは定義の avLst の名前（`adj` `adj1` …）、値は ECMA の単位（例 50000） */
  adj?: Record<string, number>
  flipH?: boolean
  flipV?: boolean
  /** 開いた path の始点（head）・終点（tail）の矢じり。値は LineElement の head / tail と同じ */
  arrow?: { head?: string; tail?: string }
}

export interface LineElement extends ElementBase {
  kind: 'line'
  line: { color: string; width: number; dash: Dash; head?: string; tail?: string }
  from: { x: number; y: number }
  to: { x: number; y: number }
}

export type ReplaceReason = 'math' | 'mermaid' | 'svg' | 'unknown-element' | 'explicit' | 'gradient'

export interface ImageElement extends ElementBase {
  kind: 'image'
  src?: string
  captureId?: string
  reason?: ReplaceReason
  fit?: 'contain' | 'cover' | 'fill'
  alt?: string
}

export interface TableElement extends ElementBase {
  kind: 'table'
  colW: number[]
  rowH: number[]
  headerRows: number
  rows: Cell[][]
}

export interface BorderSide {
  color: string
  width: number
}

export interface Cell {
  paragraphs: Paragraph[]
  colspan?: number
  rowspan?: number
  fill?: string
  border: [BorderSide, BorderSide, BorderSide, BorderSide]
  inset: [number, number, number, number]
  align: Align
  valign: 'top' | 'middle' | 'bottom'
}

export interface Warning {
  code: string
  elementId?: string
  message: string
}

export interface CheckResult {
  rule: string
  part: string
  message: string
  level: 'error' | 'warn'
}

export interface Replacement {
  slide: number
  elementId: string
  name: string
  reason: ReplaceReason | undefined
  selector: string
  width: number
  height: number
}

export interface ReportWarning extends Warning {
  slide: number
  name?: string
}

export interface Report {
  output: string
  generatedAt: string
  slidev: string
  pptxgenjs: string
  slides: number
  native: number
  replaced: number
  replacements: Replacement[]
  warnings: ReportWarning[]
  dropped: Record<string, number>
  zoom: Record<number, number>
  /** PPTX の中の番号（slideN.xml）→ 元のスライド番号。--range で絞ると 2 つがずれる。warnings / replacements は元の番号 */
  slideMap: Record<number, number>
  check: CheckResult[]
}

/** Node 側が `options.data` から使う部分（Slidev の SlidevData の抜粋。必要な項目だけ） */
export interface DeckData {
  config?: { canvasWidth?: number; aspectRatio?: number; title?: string; author?: string; colorSchema?: string; transition?: string }
  slides: DeckSlide[]
  layouts?: string[]
}

export interface DeckSlide {
  index: number
  frontmatter: Record<string, unknown> & { layout?: string; background?: string; defaults?: { layout?: string }; transition?: string }
  note?: string
  title?: string
}

export function emptyReport(output = ''): Report {
  return {
    output,
    generatedAt: new Date().toISOString(),
    slidev: '',
    pptxgenjs: '',
    slides: 0,
    native: 0,
    replaced: 0,
    replacements: [],
    warnings: [],
    dropped: {},
    zoom: {},
    slideMap: {},
    check: [],
  }
}
