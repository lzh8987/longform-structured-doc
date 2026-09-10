// Smoke test for the longform-structured-doc plugin entry: loads the module
// with a fake ctx, verifies the skill provider registration, the published
// skill list/get, and that all referenced resources exist under the skill's
// resourceBase.
// Run: node test/check.mjs   (from the plugin root; zero-dependency, no node_modules needed)

import { access } from 'node:fs/promises'
import { fileURLToPath } from 'node:url'
import { apply, name, inject } from '../lib/index.js'

const providers = []
const tools = []
const ctx = {
  skills: { registerProvider(fn) { providers.push(fn()) } },
  tools: { register(tool) { tools.push(tool) } },
}

apply(ctx)

console.log('== plugin ==')
console.log('name  :', name)
console.log('inject:', inject.join(', '))
console.log('providers:', providers.map((p) => p.name).join(', ') || '(none)')
console.log('tools:', tools.map((t) => t.name).join(', ') || '(none)')

if (providers.length !== 1) throw new Error('expected exactly one skill provider')
if (tools.length !== 0) throw new Error('pure-skill plugin must not register tools')

console.log('\n== skill provider ==')
const list = await providers[0].list()
console.log('list():', JSON.stringify(list.map((s) => ({ name: s.name, provider: s.provider, rank: s.rank, source: s.source }))))
if (list.length !== 1 || list[0].name !== 'longform-structured-doc') throw new Error('unexpected skill list')

const skill = await providers[0].get({ name: 'longform-structured-doc' })
console.log('get():  description =', String(skill?.description).slice(0, 120) + '…')
if (!skill?.content || !skill.content.includes('大规模结构化多章节文档')) throw new Error('skill content mismatch')

console.log('\n== resources under resourceBase ==')
const base = fileURLToPath(new URL('../skills/longform-structured-doc/', import.meta.url))
const required = [
  'orchestrator.md',
  'templates/writer.md',
  'templates/similarity-auditor.md',
  'templates/rewriter.md',
  'templates/policy-checker.md',
  'templates/chapter-writer.md',
  'scripts/docs_to_md.py',
  'scripts/assemble_docx.py',
  'scripts/templates.json',
  'scripts/test_roundtrip.py',
]
for (const rel of required) {
  await access(base + rel)
  console.log('OK', rel)
}

console.log('\nSMOKE TEST PASSED')
