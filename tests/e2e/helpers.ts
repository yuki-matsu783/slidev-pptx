// e2e の共通: Slidev を export モードで立て、Playwright で /print を開く（native-export.md §1.2）
import { resolve } from 'node:path'
import { createServer, resolveOptions } from '@slidev/cli'
import { chromium } from 'playwright-chromium'
import type { Browser, Page } from 'playwright-chromium'

export const DECK_DIR = resolve(__dirname, '../fixtures/deck')
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

export async function openPrint(browser: Browser, base: string, opts: { range?: string; colorScheme?: 'light' | 'dark' } = {}): Promise<Page> {
  const context = await browser.newContext({ deviceScaleFactor: 2, colorScheme: opts.colorScheme ?? 'light', viewport: { width: 980, height: 552 } })
  const page = await context.newPage()
  const url = `${base}/print?print=true${opts.range ? `&range=${opts.range}` : ''}`
  await page.goto(url, { waitUntil: 'networkidle' })
  await page.waitForSelector('[data-slidev-no]')
  await page.waitForSelector('.slidev-slide-loading', { state: 'detached' }).catch(() => {})
  return page
}

export async function launch(): Promise<Browser> {
  return chromium.launch()
}
