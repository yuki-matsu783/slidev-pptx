// e2e の共通: Slidev を export モードで立て、Playwright で /print を開く（native-export.md §1.2）
import { fileURLToPath } from 'node:url'
import { resolve } from 'node:path'
import { createServer, resolveOptions } from '@slidev/cli'
import { chromium } from 'playwright-chromium'
import type { Browser, Page } from 'playwright-chromium'

const here = fileURLToPath(new URL('.', import.meta.url))
export const DECK_DIR = resolve(here, '../fixtures/deck')
export const PLAIN = resolve(DECK_DIR, 'plain.md')
export const COMPONENTS = resolve(DECK_DIR, 'components.md')
export const DARK = resolve(DECK_DIR, 'dark.md')

export interface Running {
  port: number
  base: string
  close(): Promise<void>
}

export async function startSlidev(entry: string): Promise<Running> {
  const options = await resolveOptions({ entry }, 'export')
  const server = await createServer(options, { server: { port: 0 } })
  await server.listen()
  const addr = server.httpServer?.address()
  const port = typeof addr === 'object' && addr ? addr.port : Number(server.config.server.port)
  return { port, base: `http://localhost:${port}`, close: () => server.close() }
}

/**
 * 設計 §1.2 の待機列を写す: [data-slidev-no] の出現 → .slidev-slide-loading の消滅 → [data-waitfor] →
 * mermaid の svg → networkidle。viewport は Slidev の exportSlides と同じ「幅 × 高さ×枚数」。
 */
export async function openPrint(browser: Browser, base: string, opts: { range?: string; colorScheme?: 'light' | 'dark'; slides?: number } = {}): Promise<Page> {
  const context = await browser.newContext({
    deviceScaleFactor: 2,
    colorScheme: opts.colorScheme ?? 'light',
    viewport: { width: 980, height: 552 * (opts.slides ?? 1) },
  })
  const page = await context.newPage()
  const url = `${base}/print?print=true${opts.range ? `&range=${opts.range}` : ''}`
  await page.goto(url, { waitUntil: 'networkidle' })
  await page.waitForSelector('[data-slidev-no]')
  await page.waitForSelector('.slidev-slide-loading', { state: 'detached', timeout: 30_000 }).catch(() => {})
  await page.waitForFunction(() => Array.from(document.querySelectorAll('[data-waitfor]')).every((el) => el.querySelector(el.getAttribute('data-waitfor')!)), null, { timeout: 30_000 }).catch(() => {})
  await page.waitForFunction(() => Array.from(document.querySelectorAll('.mermaid')).every((el) => (el.shadowRoot ?? el).querySelector('svg')), null, { timeout: 30_000 }).catch(() => {})
  await page.waitForLoadState('networkidle')
  return page
}

export async function launch(): Promise<Browser> {
  return chromium.launch()
}
