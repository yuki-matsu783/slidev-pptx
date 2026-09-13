// CLI（native-export.md §7）: exit code、フラグ、package.json の scripts の綴り
import { spawnSync } from 'node:child_process'
import { existsSync, mkdtempSync, readFileSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

const here = fileURLToPath(new URL('.', import.meta.url))
import { PATCHES, postProcess } from '../../packages/slidev-addon-pptx/src/patch/index'
import { readFixturePptx } from '../helpers/pptx'
import { contextFor } from '../patch/helpers'
import { PLAIN } from '../e2e/helpers'

const ROOT = resolve(here, '../..')
// .bin のリンク（pnpm install で作られる）。起動には使わない: Windows では .bin の実体は sh スクリプトで、
// .CMD は Node 22 では shell: true 無しに spawn できない（CVE-2024-27980 の修正）。
// 起動は process.execPath + パッケージの bin エントリで、OS と .bin のリンクから独立させる
const BIN_LINK = resolve(ROOT, 'node_modules/.bin/slidev-pptx')
const BIN_ENTRY = resolve(ROOT, 'packages/slidev-addon-pptx/bin/slidev-pptx.mjs')
const tmp = () => mkdtempSync(join(tmpdir(), 'slidev-pptx-cli-'))

function run(args: string[], cwd = ROOT, timeout = 180_000) {
  const r = spawnSync(process.execPath, [BIN_ENTRY, ...args], { cwd, encoding: 'utf8', timeout })
  return { code: r.status, out: r.stdout, err: r.stderr }
}

describe('cli: 引数', () => {
  it('bin slidev-pptx が workspace から解決できる（pnpm install 後に .bin にリンクがある）', () => {
    expect(existsSync(BIN_LINK) || existsSync(BIN_LINK + '.CMD')).toBe(true)
    expect(existsSync(BIN_ENTRY)).toBe(true)
  })
  it('未知のサブコマンド / 未知のフラグは exit 2', () => {
    expect(run(['bogus']).code).toBe(2)
    expect(run(['export', PLAIN, '--bogus']).code).toBe(2)
  })
  it('entry が無ければ exit 1', () => {
    expect(run(['export', join(tmp(), 'nope.md'), '--no-check']).code).toBe(1)
  })
})

describe('cli: check', () => {
  it('後処理前の fixture は error があるので exit 1、規則名が出る', () => {
    const r = run(['check', resolve(ROOT, 'tests/fixtures/pptx/sample.pptx')])
    expect(r.code).toBe(1)
    expect(r.out + r.err).toMatch(/C6/)
  })
  it('後処理を通した PPTX は exit 0', async () => {
    const out = join(tmp(), 'ok.pptx')
    writeFileSync(out, await postProcess(readFixturePptx(), PATCHES, contextFor()))
    expect(run(['check', out]).code).toBe(0)
  })
})

describe('cli: export', () => {
  it('既定の出力名は ./slides-export.pptx、記録は <output>.report.json', () => {
    const dir = tmp()
    const r = run(['export', PLAIN, '--range', '1'], dir)
    expect(r.code).toBe(0)
    expect(existsSync(join(dir, 'slides-export.pptx'))).toBe(true)
    expect(existsSync(join(dir, 'slides-export.pptx.report.json'))).toBe(true)
  }, 200_000)

  it('-o と --report で置き場を変えられ、標準出力に要約が出る', () => {
    const dir = tmp()
    const r = run(['export', PLAIN, '-o', join(dir, 'a.pptx'), '--report', join(dir, 'r.json'), '--range', '1-2'])
    expect(r.code).toBe(0)
    expect(existsSync(join(dir, 'a.pptx'))).toBe(true)
    expect(JSON.parse(readFileSync(join(dir, 'r.json'), 'utf8')).slides).toBe(2)
    expect(r.out).toMatch(/2 slides/)
  }, 200_000)

  it('--strict は警告があれば exit 1（plain.md には W-TRANSITION がある）', () => {
    const dir = tmp()
    expect(run(['export', PLAIN, '-o', join(dir, 's.pptx'), '--range', '1', '--strict']).code).toBe(1)
  }, 200_000)

  it('--lang が run の既定 lang になる', async () => {
    const dir = tmp()
    const out = join(dir, 'l.pptx')
    expect(run(['export', PLAIN, '-o', out, '--range', '2', '--lang', 'en-US']).code).toBe(0)
    const { openPptx, els } = await import('../helpers/pptx')
    const p = await openPptx(readFileSync(out))
    const langs = new Set(els(await p.xml('ppt/slides/slide1.xml'), 'a', 'rPr').map((r) => r.getAttribute('lang')))
    expect(langs).toEqual(new Set(['en-US']))
  }, 200_000)
})

describe('package.json の scripts（§7）', () => {
  const pkg = JSON.parse(readFileSync(resolve(ROOT, 'package.json'), 'utf8'))
  it('export:pptx はネイティブ書き出し、export:pptx-image は Slidev の画像書き出し', () => {
    expect(pkg.scripts['export:pptx']).toBe('slidev-pptx export slides.md')
    expect(pkg.scripts['export:pptx-image']).toBe('slidev export --format pptx')
  })
})
