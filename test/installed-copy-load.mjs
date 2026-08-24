// Installed-copy load test: verifies the plugin as it is INSTALLED into a dsh
// profile (the pnpm-store copy under the profile's node_modules), not the
// source tree. Proves that package.json dependency resolution, peer
// dependencies, and every bundled skill resource work after installation.
//
// Usage:  node test/installed-copy-load.mjs [profileName]   (default: web)
//   - resolves $DSH_HOME/profiles/<profileName>
//   - $DSH_HOME defaults to ~/.dsh when unset
//   - the profile must already have longform-structured-doc installed:
//       dsh plugin --profile <profileName> add "file:<path-to-this-repo>"
//
// Zero extra dependencies: plain Node.js (>=18), run from the plugin root.

import { access } from 'node:fs/promises'
import { existsSync } from 'node:fs'
import { join } from 'node:path'
import os from 'node:os'
import { fileURLToPath, pathToFileURL } from 'node:url'
import { createRequire } from 'node:module'

const profileName = process.argv[2] ?? 'web'
const dshHome = process.env.DSH_HOME ?? join(os.homedir(), '.dsh')
const profileDir = join(dshHome, 'profiles', profileName)
if (!existsSync(join(profileDir, 'package.json'))) {
  console.error(`profile not found: ${profileDir}`)
  console.error('install the plugin into it first, e.g.:')
  console.error(`  dsh plugin --profile ${profileName} add "file:<path-to-this-repo>"`)
  process.exit(1)
}

console.log(`profile: ${profileName}  (${profileDir})`)
const require = createRequire(join(profileDir, 'package.json'))
const entry = require.resolve('longform-structured-doc')
console.log('resolved entry:', entry)
if (!entry.includes('node_modules')) throw new Error('entry is not under node_modules — not the installed copy: ' + entry)

const mod = await import(pathToFileURL(entry).href)
const { apply, name, inject } = mod
console.log('name  :', name)
console.log('inject:', inject.join(', '))

const providers = []
apply({ skills: { registerProvider(fn) { providers.push(fn()) } } })
if (providers.length !== 1) throw new Error('expected one provider, got ' + providers.length)

const list = await providers[0].list()
console.log('list():', JSON.stringify(list.map((s) => ({ name: s.name, source: s.source, rank: s.rank }))))
if (list.length !== 1 || list[0].name !== 'longform-structured-doc') throw new Error('bad skill list')

const skill = await providers[0].get({ name: 'longform-structured-doc' })
console.log('get(): description =', String(skill?.description).slice(0, 80) + '…')
if (!skill?.content?.includes('大规模结构化多章节文档')) throw new Error('skill content mismatch')

// resourceBase must point at real files inside the INSTALLED copy
const base = skill.resourceBase.path
console.log('resourceBase:', base)
if (!base.includes('node_modules')) throw new Error('resourceBase not under installed node_modules: ' + base)
const required = [
  'orchestrator.md',
  'templates/writer.md',
  'templates/similarity-auditor.md',
  'templates/rewriter.md',
  'templates/policy-checker.md',
  'templates/chapter-writer.md',
  'scripts/docs_to_md.py',
  'scripts/assemble_docx.py',
]
for (const rel of required) {
  await access(base + rel)
  console.log('OK', rel)
}

console.log('\nINSTALLED-COPY LOAD TEST PASSED')
