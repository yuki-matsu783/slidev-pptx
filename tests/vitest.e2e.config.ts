// ブラウザと Slidev サーバを使う検査。`vitest run -c tests/vitest.e2e.config.ts`（playwright-chromium が要る）
import { fileURLToPath } from 'node:url'
import { resolve } from 'node:path'
import { defineConfig } from 'vitest/config'

const here = fileURLToPath(new URL('.', import.meta.url))

export default defineConfig({
  root: resolve(here, '..'),
  test: {
    include: ['tests/e2e/**/*.test.ts', 'tests/cli/**/*.test.ts'],
    testTimeout: 180_000,
    hookTimeout: 120_000,
    fileParallelism: false,
  },
})
