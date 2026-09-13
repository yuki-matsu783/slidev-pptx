// ブラウザと Slidev サーバを使う検査。`vitest -c tests/vitest.e2e.config.ts`（playwright-chromium が要る）
import { resolve } from 'node:path'
import { defineConfig } from 'vitest/config'

export default defineConfig({
  root: resolve(__dirname, '..'),
  test: {
    include: ['tests/e2e/**/*.test.ts', 'tests/cli/**/*.test.ts'],
    testTimeout: 180_000,
    hookTimeout: 120_000,
    fileParallelism: false,
  },
})
