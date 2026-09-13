// 単体の検査（ブラウザ無し）。`vitest run -c tests/vitest.config.ts`
// ルートの package.json / vitest.config.ts はこのフェーズの範囲外なので、設定は tests/ の中に置く。
import { fileURLToPath } from 'node:url'
import { resolve } from 'node:path'
import { defineConfig } from 'vitest/config'

const here = fileURLToPath(new URL('.', import.meta.url))

export default defineConfig({
  root: resolve(here, '..'),
  test: {
    include: ['tests/**/*.test.ts'],
    exclude: ['tests/e2e/**', 'tests/cli/**', '**/node_modules/**'],
    testTimeout: 30_000,
    hookTimeout: 60_000, // convert.test の beforeAll は画像入りの PPTX を書く
    typecheck: {
      enabled: true,
      tsconfig: 'tests/tsconfig.json', // ルートに tsconfig.json は無い（置くのは範囲外）
      include: ['tests/**/*.test-d.ts'], // 型の主張（expectTypeOf）はここでだけ効く。typescript と @types/node が要る
    },
  },
})
