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
/** PowerPoint の図形 187 種を格子に並べたデッキ（最後の 1 枚は調整値・反転・矢じり・文字の枠の見本） */
export const SHAPES = resolve(DECK_DIR, 'shapes.md')

export interface Running {
  port: number
  base: string
  close(): Promise<void>
}

export async function startSlidev(entry: string): Promise<Running> {
  // vitest は NODE_ENV=test にする。その値だと Vite 開発サーバで UnoCSS がデッキ由来のクラスを生成しない
  // （実測: test だと .pt-12 は永遠に 0px、development なら読み込み直後に 48px）。サーバを立てる間は development にする
  const prev = process.env.NODE_ENV
  process.env.NODE_ENV = 'development'
  const options = await resolveOptions({ entry }, 'export')
  const server = await createServer(options, { server: { port: 0 } })
  await server.listen()
  if (prev === undefined) delete process.env.NODE_ENV
  else process.env.NODE_ENV = prev
  const addr = server.httpServer?.address()
  const port = typeof addr === 'object' && addr ? addr.port : Number(server.config.server.port)
  return { port, base: `http://localhost:${port}`, close: () => server.close() }
}

/**
 * 設計 §1.2 の待機列を再現する: [data-slidev-no] の出現 → .slidev-slide-loading の消滅 → [data-waitfor] →
 * mermaid の svg → networkidle。viewport は Slidev の exportSlides と同じ「幅 × 高さ×枚数」。
 */
export interface ReadyCheck {
  selector: string
  prop: string
  value: string
}

/** plain.md 用: デッキ由来のユーティリティ（.pt-12）が UnoCSS で生成されるまで待つ。Slidev 自身の規則（px-14）は
 *  transformer-directives がファイル変換時に展開するので、デッキのクラスが 1 つも生成されていなくても効いてしまい、生成の確認には使えない */
export const PLAIN_READY: ReadyCheck = { selector: '.pt-12', prop: 'paddingTop', value: '48px' }

export async function openPrint(browser: Browser, base: string, opts: { range?: string; colorScheme?: 'light' | 'dark'; slides?: number; ready?: ReadyCheck } = {}): Promise<Page> {
  const context = await browser.newContext({
    deviceScaleFactor: 2,
    colorScheme: opts.colorScheme ?? 'light',
    viewport: { width: 980, height: 552 * (opts.slides ?? 1) },
  })
  const page = await context.newPage()
  const url = `${base}/print?print=true${opts.range ? `&range=${opts.range}` : ''}`
  await page.goto(url, { waitUntil: 'networkidle' })
  // Vite が初回に依存を最適化するとページを再読み込みする（"optimized dependencies changed. reloading"）。
  // 枚数が揃うまで少し待って再試行する
  for (let i = 0; i < 10; i++) {
    await page.waitForSelector('[data-slidev-no]')
    const n = await page.evaluate(() => document.querySelectorAll('.print-slide-container').length)
    if (!opts.slides || n >= opts.slides) break
    await page.waitForTimeout(1000)
  }
  await page.waitForSelector('[data-slidev-no]')
  await page.waitForSelector('.slidev-slide-loading', { state: 'detached', timeout: 30_000 }).catch(() => {})
  await page.waitForFunction(() => Array.from(document.querySelectorAll('[data-waitfor]')).every((el) => el.querySelector(el.getAttribute('data-waitfor')!)), null, { timeout: 30_000 }).catch(() => {})
  await page.waitForFunction(() => Array.from(document.querySelectorAll('.mermaid')).every((el) => (el.shadowRoot ?? el).querySelector('svg')), null, { timeout: 30_000 }).catch(() => {})
  await page.waitForFunction(() => document.fonts.status === 'loaded', null, { timeout: 30_000 }).catch(() => {})
  await page.waitForLoadState('networkidle')
  // UnoCSS は dev では後から HMR で CSS を注入し、来ないこともある（同じデッキで run ごとに違う）。
  // デッキ由来のクラスが効いたことを確認してから測る。来なければ 30 秒で打ち切り、黙って誤った値を測らない
  //（設計 §1.2 の待機列には無い。README に記載）
  const ready = opts.ready
  if (ready) {
    // 最初の読み込みでは UnoCSS がデッキのクラスを生成しないことがある。再読み込みで効く（実測）ので、1 回だけやり直す
    const check = () =>
      page.waitForFunction(
        ({ selector, prop, value }) => {
          const el = document.querySelector(selector)
          return !!el && (getComputedStyle(el) as unknown as Record<string, string>)[prop] === value
        },
        ready,
        { timeout: 15_000 },
      )
    try {
      await check()
    } catch {
      await page.reload({ waitUntil: 'networkidle' })
      await page.waitForSelector('[data-slidev-no]')
      await page.waitForFunction(() => Array.from(document.querySelectorAll('.mermaid')).every((el) => (el.shadowRoot ?? el).querySelector('svg')), null, { timeout: 30_000 }).catch(() => {})
      await check()
    }
  }
  // さらに <style> の合計長が 2 回続けて（1 秒）動かないことを確かめる
  let last = -1
  let stable = 0
  for (let i = 0; i < 30 && stable < 2; i++) {
    const len = await page.evaluate(() => Array.from(document.querySelectorAll('style')).reduce((n, s) => n + (s.textContent?.length ?? 0), 0))
    stable = len === last ? stable + 1 : 0
    last = len
    await page.waitForTimeout(500)
  }
  return page
}

export async function launch(): Promise<Browser> {
  return chromium.launch()
}
