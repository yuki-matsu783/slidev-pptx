// 書き出しの流れ（native-export.md §1）。Slidev を export モードで立て、Playwright で /print を開き、
// 収集 → 生成 → 後処理 → OPC 検査 → 記録。
import { mkdirSync, writeFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { createRequire } from 'node:module'
import { createServer, resolveOptions } from '@slidev/cli'
import playwright from 'playwright-chromium'
import type { Browser, Page } from 'playwright-chromium'
import type { Capture, DeckData, Report, SlideCapture } from './types.ts'
import { collect } from './collect/index.ts'
import { build } from './build/convert.ts'
import type { Asset } from './build/convert.ts'
import { PATCHES, postProcess } from './patch/index.ts'
import { check } from './opc/check.ts'

export interface ExportOptions {
  entry: string
  output?: string
  range?: string
  theme?: string
  lang?: string
  wait?: number
  timeout?: number
  report?: string
  check?: boolean
  strict?: boolean
  keepServer?: boolean
  /** 標準出力に要約を書く（CLI が true にする） */
  log?: (line: string) => void
}

const require = createRequire(import.meta.url)

export async function exportPptx(o: ExportOptions): Promise<Report> {
  const entry = resolve(o.entry)
  const output = resolve(o.output ?? 'slides-export.pptx')
  const reportPath = resolve(o.report ?? `${output}.report.json`)
  const timeout = o.timeout ?? 30_000

  // NODE_ENV が test / production だと Vite 開発サーバで UnoCSS がデッキ由来のクラスを生成しない（実測）。
  // サーバを立てる間だけ development にする
  const prevEnv = process.env.NODE_ENV
  process.env.NODE_ENV = 'development'
  const options = await resolveOptions({ entry, theme: o.theme }, 'export')
  const server = await createServer(options, { server: { port: 0 } })
  await server.listen()
  if (prevEnv === undefined) delete process.env.NODE_ENV
  else process.env.NODE_ENV = prevEnv
  const addr = server.httpServer?.address()
  const port = typeof addr === 'object' && addr ? addr.port : Number(server.config.server.port)
  const base = `http://localhost:${port}`

  mkdirSync(dirname(output), { recursive: true })
  mkdirSync(dirname(reportPath), { recursive: true })
  let browser: Browser | undefined
  let failed = false
  try {
    browser = await playwright.chromium.launch()
    const data = toDeckData(options.data)
    const total = data.slides.length
    if (o.range && !isValidRange(o.range)) throw new Error(`--range の綴りが不正: ${o.range}（例: 1-3,5）`)
    const slideCount = o.range ? countRange(o.range, total) : total
    const canvasWidth = data.config?.canvasWidth ?? 980
    const canvasHeight = Math.round(canvasWidth / (data.config?.aspectRatio ?? 16 / 9))
    // viewport の高さは Slidev と同じ「高さ × 枚数」。Chromium の上限を超えないよう 16384 px で止める（rect はコンテナ基準なのでスクロールしても測れる）
    const context = await browser.newContext({ deviceScaleFactor: 2, colorScheme: 'light', viewport: { width: canvasWidth, height: Math.min(16384, canvasHeight * Math.max(1, slideCount)) } })
    const page = await context.newPage()
    const pageErrors: string[] = []
    page.on('pageerror', (e) => pageErrors.push(String(e.message ?? e)))
    page.on('console', (m) => {
      if (m.type() === 'error' && !/FloatingVue/.test(m.text())) pageErrors.push(m.text())
    })
    await openPrint(page, base, o.range, timeout, o.wait ?? 0, slideCount)

    // Vite の依存最適化（"optimized dependencies changed. reloading"）が収集の最中に来ると実行文脈が壊れる。
    // その場合は読み込み直しを待って 1 回だけやり直す
    let capture: Capture
    try {
      capture = await page.evaluate(collect)
    } catch (e) {
      if (!/Execution context was destroyed|navigation|Target closed/i.test(String(e))) throw e
      await page.waitForLoadState('load', { timeout })
      await openPrint(page, base, o.range, timeout, o.wait ?? 0, slideCount)
      capture = await page.evaluate(collect)
    }
    // Slidev の /print は useNav の初期化時に query.range を 1 度読むだけで、この経路では効かないことがある。
    // URL に渡したうえで、Node 側でも range で絞る（観測できる結果は同じ）
    if (o.range) {
      const keep = parseRange(o.range, data.slides.length)
      capture.slides = capture.slides.filter((s) => keep.has(s.no))
    }
    const assets = await collectAssets(page, capture, data, base)

    const layouts = Object.keys(await options.utils.getLayouts())
    const { pptx, ctx } = await build(capture, data, { assets, lang: o.lang, layouts, output, slidevVersion: slidevVersion() })
    for (const sc of capture.slides as (SlideCapture & { dropped?: Record<string, number> })[]) {
      for (const [k, v] of Object.entries(sc.dropped ?? {})) ctx.report.dropped[k] = (ctx.report.dropped[k] ?? 0) + v
    }
    for (const msg of pageErrors) ctx.report.warnings.push({ code: 'W-RENDER', slide: 0, message: `ブラウザでエラー: ${msg.slice(0, 200)}` })
    const raw = (await pptx.write({ outputType: 'nodebuffer' })) as Buffer
    const out = await postProcess(raw, PATCHES, ctx)
    writeFileSync(output, out)

    if (o.check !== false) ctx.report.check = await check(out)
    ctx.report.generatedAt = new Date().toISOString()
    writeFileSync(reportPath, JSON.stringify(ctx.report, null, 2))
    for (const line of summaryLines(ctx.report)) o.log?.(line)
    return ctx.report
  } catch (e) {
    failed = true
    throw e
  } finally {
    // --keep-server は失敗時だけサーバとブラウザを残す（調査用）
    if (!(o.keepServer && failed)) {
      await browser?.close().catch(() => {})
      await server.close().catch(() => {})
    }
  }
}

export function isValidRange(range: string): boolean {
  return range.split(',').every((part) => {
    const m = /^\s*(\d+)\s*(?:-\s*(\d+)\s*)?$/.exec(part)
    return !!m && Number(m[1]) >= 1 && (!m[2] || Number(m[2]) >= Number(m[1]))
  })
}

// ---------------------------------------------------------------- ブラウザ

/** 設計 §1.2 の待機列 + UnoCSS の遅延注入を待つ */
async function openPrint(page: Page, base: string, range: string | undefined, timeout: number, extraWait: number, expectedSlides: number): Promise<void> {
  const url = `${base}/print?print=true${range ? `&range=${encodeURIComponent(range)}` : ''}`
  if (page.url() !== url) await page.goto(url, { waitUntil: 'networkidle', timeout })
  else await page.waitForLoadState('networkidle', { timeout })
  await page.emulateMedia({ colorScheme: 'light', media: 'screen' })
  // Vite が初回に依存を最適化するとページを再読み込みする。枚数が揃うまで少し待って再試行する
  // （/print は range を無視することがあるので、期待は「全枚数以上」ではなく「1 枚以上」で十分）
  for (let i = 0; i < 10; i++) {
    await page.waitForSelector('[data-slidev-no]', { timeout })
    const n = await page.evaluate(() => document.querySelectorAll('.print-slide-container').length)
    if (n >= Math.min(1, expectedSlides) && (await page.evaluate(() => document.readyState)) === 'complete') break
    await page.waitForTimeout(1000)
  }
  await page.waitForSelector('[data-slidev-no]', { timeout })
  await page.waitForSelector('.slidev-slide-loading', { state: 'detached', timeout }).catch(() => {})
  await page.waitForFunction(() => Array.from(document.querySelectorAll('[data-waitfor]')).every((el) => el.querySelector(el.getAttribute('data-waitfor')!)), null, { timeout }).catch(() => {})
  await page.waitForFunction(() => Array.from(document.querySelectorAll('.mermaid')).every((el) => (el.shadowRoot ?? el).querySelector('svg')), null, { timeout }).catch(() => {})
  await page.waitForFunction(() => document.fonts.status === 'loaded', null, { timeout }).catch(() => {})
  await page.waitForLoadState('networkidle')
  // UnoCSS は dev では後から CSS を注入する。<style> の合計長が 2 回続けて（1 秒）動かないまで待つ
  let last = -1
  let stable = 0
  for (let i = 0; i < 40 && stable < 2; i++) {
    const len = await page.evaluate(() => Array.from(document.querySelectorAll('style')).reduce((n, s) => n + (s.textContent?.length ?? 0), 0))
    stable = len === last ? stable + 1 : 0
    last = len
    await page.waitForTimeout(500)
  }
  if (extraWait > 0) await page.waitForTimeout(extraWait)
}

/** 置き換え画像の撮影、URL 画像と背景画像の取得 */
async function collectAssets(page: Page, capture: Capture, data: DeckData, base: string): Promise<Record<string, string | Asset>> {
  const assets: Record<string, string | Asset> = {}
  const fetched = new Map<string, Promise<string | undefined>>()
  const fetchDataUrl = (url: string): Promise<string | undefined> => {
    if (!fetched.has(url)) {
      fetched.set(
        url,
        (async () => {
          try {
            const res = await fetch(url)
            if (!res.ok) return undefined
            const type = res.headers.get('content-type')?.split(';')[0] || 'application/octet-stream'
            // Vite の開発サーバは無い path にも index.html を 200 で返す。画像でなければ取得失敗とみなす
            if (!/^image\//.test(type)) return undefined
            const buf = Buffer.from(await res.arrayBuffer())
            if (/svg/.test(type) || /\.svg(\?|$)/i.test(url)) return 'svg'
            return `data:${type};base64,${buf.toString('base64')}`
          } catch {
            return undefined
          }
        })(),
      )
    }
    return fetched.get(url)!
  }
  const shoot = async (captureId: string): Promise<Asset | undefined> => {
    const loc = page.locator(`[data-ppt-capture-id="${captureId}"]`).first()
    if ((await loc.count()) === 0) return undefined
    try {
      const png = await loc.screenshot({ type: 'png', omitBackground: true, animations: 'disabled', timeout: 15_000 })
      const selector = await loc.evaluate((el) => {
        const cls = Array.from(el.classList).slice(0, 2).map((c) => '.' + c).join('')
        return el.tagName.toLowerCase() + cls
      })
      return { data: `data:image/png;base64,${png.toString('base64')}`, width: png.readUInt32BE(16), height: png.readUInt32BE(20), selector }
    } catch {
      return undefined
    }
  }

  for (const sc of capture.slides) {
    for (const e of sc.elements) {
      if (e.kind !== 'image') continue
      if (e.captureId) {
        const a = await shoot(e.captureId)
        if (a) assets[e.captureId] = a
        continue
      }
      if (!e.src || e.src.startsWith('data:')) continue
      const d = await fetchDataUrl(e.src)
      if (d === 'svg') {
        // PptxGenJS の SVG 経路は壊れた PPTX を出すので撮影に回す（§4.3）。収集器が data-ppt-capture-id を付けている。
        // 置き換えとして扱い、記録（reason: 'svg'）と名前（Replaced N）に出す
        e.captureId = e.id
        e.reason = 'svg'
        e.source = 'replaced'
        const a = await shoot(e.id)
        if (a) assets[e.id] = a
      } else if (d) assets[e.src] = d
    }
    const bg = data.slides[sc.no - 1]?.frontmatter?.background
    if (typeof bg === 'string' && bg && !/^(#|rgb|hsl|data:)/.test(bg)) {
      const url = /^https?:/.test(bg) ? bg : new URL(bg, base + '/').href
      const d = await fetchDataUrl(url)
      if (d && d !== 'svg') assets[`background:${sc.no}`] = d
    }
  }
  return assets
}

// ---------------------------------------------------------------- 補助

function toDeckData(data: { slides: unknown[]; config: object }): DeckData {
  const slides = (data.slides as { index: number; frontmatter: Record<string, unknown>; note?: string; title?: string }[]).map((s, i) => ({
    index: s.index ?? i,
    frontmatter: (s.frontmatter ?? {}) as DeckData['slides'][number]['frontmatter'],
    note: s.note,
    title: s.title,
  }))
  const c = data.config as Record<string, unknown>
  return {
    slides,
    config: {
      canvasWidth: c.canvasWidth as number | undefined,
      aspectRatio: c.aspectRatio as number | undefined,
      title: c.title as string | undefined,
      author: c.author as string | undefined,
      colorSchema: c.colorSchema as string | undefined,
      transition: (c.transition as string | undefined) ?? undefined,
    },
  }
}

/** Slidev の range と同じ綴り（'1-3,5'）。空や不正なら全部 */
export function parseRange(range: string, total: number): Set<number> {
  const set = new Set<number>()
  for (const part of range.split(',')) {
    const m = /^\s*(\d+)\s*(?:-\s*(\d+)\s*)?$/.exec(part)
    if (!m) continue
    const a = Number(m[1])
    const b = m[2] ? Number(m[2]) : a
    for (let i = Math.max(1, a); i <= Math.min(total, b); i++) set.add(i)
  }
  if (!set.size) for (let i = 1; i <= total; i++) set.add(i)
  return set
}

function countRange(range: string, total: number): number {
  return parseRange(range, total).size
}

function slidevVersion(): string {
  try {
    return (require('@slidev/cli/package.json') as { version: string }).version
  } catch {
    return ''
  }
}

export function summaryLines(r: Report): string[] {
  const dropped = Object.entries(r.dropped)
    .map(([k, v]) => `${k} ×${v}`)
    .join(', ')
  const lines = [`${r.output}: ${r.slides} slides, ${r.native} native, ${r.replaced} replaced, ${r.warnings.length} warnings${dropped ? `, dropped by rule: ${dropped}` : ''}`]
  for (const x of r.replacements) lines.push(`  slide ${x.slide}  replaced  ${(x.reason ?? '').padEnd(15)} ${x.selector} (${x.elementId} "${x.name}") → image ${x.width}×${x.height}`)
  for (const w of r.warnings) lines.push(`  slide ${w.slide}  warning   ${w.code.padEnd(15)} ${w.message}${w.elementId ? ` (${w.elementId}${w.name ? ` "${w.name}"` : ''})` : ''}`)
  for (const c of r.check) lines.push(`  check    ${c.level.padEnd(8)} ${c.rule.padEnd(4)} ${c.part}: ${c.message}`)
  return lines
}
