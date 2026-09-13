// 単体の検査（ブラウザ無し）。`vitest -c tests/vitest.config.ts`
// ルートの package.json / vitest.config.ts はこのフェーズの範囲外なので、設定は tests/ の中に置く。
import { resolve } from 'node:path'
import { defineConfig } from 'vitest/config'

export default defineConfig({
  root: resolve(__dirname, '..'),
  test: {
    include: ['tests/**/*.test.ts'],
    exclude: ['tests/e2e/**', 'tests/cli/**', '**/node_modules/**'],
    testTimeout: 30_000,
  },
})
