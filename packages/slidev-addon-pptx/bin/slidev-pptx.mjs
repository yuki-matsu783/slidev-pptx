#!/usr/bin/env node
// ビルド無しで src/cli.ts を読む。Node 22.18 以降の型ストリップ（既定で有効）に乗せる。
// それより古い Node では --experimental-strip-types を付けて起動し直す。
import { spawnSync } from 'node:child_process'
import { fileURLToPath } from 'node:url'

const [major, minor] = process.versions.node.split('.').map(Number)
const stripByDefault = major > 22 || (major === 22 && minor >= 18)
const cli = new URL('../src/cli.ts', import.meta.url)

if (!stripByDefault && !process.execArgv.includes('--experimental-strip-types')) {
  const r = spawnSync(process.execPath, ['--experimental-strip-types', '--no-warnings', fileURLToPath(import.meta.url), ...process.argv.slice(2)], { stdio: 'inherit' })
  process.exit(r.status ?? 1)
}

const { main } = await import(cli.href)
process.exit(await main(process.argv.slice(2)))
