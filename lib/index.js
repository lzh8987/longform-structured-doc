/**
 * longform-structured-doc — 大规模长文档生成 skill 的 DeepSeek Harness 插件。
 *
 * 只做一件事：向 `ctx.skills` 注册一个 provider，发布打包在
 * `skills/longform-structured-doc/` 下的完整 skill 树：
 *
 *   - SKILL.md        方法与纪律参考（Mode A/B/C 路由、篇幅基线、反抄袭审计、
 *                      数值回源核查、空缺标注、去AI味、docx 组装等）
 *   - orchestrator.md 执行入口（六个阶段 + 闸门的照做清单）
 *   - templates/      五个角色 prompt（writer / similarity-auditor / rewriter /
 *                      policy-checker / chapter-writer）
 *   - scripts/        docs_to_md.py（源文档拆 md）与 assemble_docx.py（组装 docx）
 *
 * 模型按需通过 skill 的 resourceBase 解析相对路径读取上述资源；流水线本身
 * 由模型依 orchestrator 调用 python 脚本执行，插件无需封装任何工具。
 *
 * 纯 ESM、零构建、零依赖（不 import 任何 @deepseek-ai 包）：`dsh plugin add`
 * 拉的是源码，无 build 步骤，也没有需要解析的 peer 依赖。
 */

import { readFile } from 'node:fs/promises'
import { fileURLToPath } from 'node:url'

const PROVIDER_NAME = 'longform-structured-doc'
const SKILL_NAME = 'longform-structured-doc'
/** Rank of the `bundled` discovery source in dsh's own skill provider. */
const BUNDLED_SKILL_RANK = 600

const SKILL_DIR_URL = new URL('../skills/longform-structured-doc/', import.meta.url)
const SKILL_BODY_URL = new URL('SKILL.md', SKILL_DIR_URL)
const RESOURCE_BASE = { kind: 'directory', path: fileURLToPath(SKILL_DIR_URL) }
const INVOCATION = { modelInvocable: true, userInvocable: true }

/**
 * 读取 SKILL.md frontmatter 里的单行 `description`。SKILL.md 是 skill 路由的
 * 唯一事实来源，把它的描述复制到本文件就是第二份会悄悄腐烂的拷贝。
 *
 * @param {string} body - SKILL.md 原文。
 * @returns {string | undefined} description，缺失时返回 undefined。
 */
function readDescription(body) {
  const frontmatter = /^---\r?\n([\s\S]*?)\r?\n---/.exec(body)
  if (!frontmatter) return undefined
  const line = /^description:[ \t]*(.+)$/m.exec(frontmatter[1])
  if (!line) return undefined
  const raw = line[1].trim()
  if (raw.startsWith('"') && raw.endsWith('"') && raw.length > 1) {
    return raw.slice(1, -1).replace(/\\"/g, '"').replace(/\\\\/g, '\\')
  }
  if (raw.startsWith("'") && raw.endsWith("'") && raw.length > 1) {
    return raw.slice(1, -1).replace(/''/g, "'")
  }
  return raw
}

/**
 * 加载打包的 skill；读取失败时返回 undefined。
 *
 * @param {AbortSignal | undefined} signal - 查询取消信号。
 * @returns {Promise<{ description: string, content: string } | undefined>}
 */
async function loadSkill(signal) {
  const content = await readFile(SKILL_BODY_URL, { encoding: 'utf8', signal })
  const description = readDescription(content)
  if (description === undefined) return undefined
  return { description, content }
}

const provider = {
  name: PROVIDER_NAME,
  async list(options = {}) {
    const skill = await loadSkill(options.signal)
    if (skill === undefined) return []
    return [{
      name: SKILL_NAME,
      description: skill.description,
      invocation: INVOCATION,
      provider: PROVIDER_NAME,
      source: 'bundled',
      resourceBase: RESOURCE_BASE,
      rank: BUNDLED_SKILL_RANK,
      locator: SKILL_BODY_URL,
    }]
  },
  async get(_candidate, options = {}) {
    const skill = await loadSkill(options.signal)
    if (skill === undefined) return undefined
    return {
      name: SKILL_NAME,
      description: skill.description,
      invocation: INVOCATION,
      provider: PROVIDER_NAME,
      source: 'bundled',
      resourceBase: RESOURCE_BASE,
      content: skill.content,
    }
  },
}

/** Cordis 插件名。 */
export const name = 'longform-structured-doc'
/** 本插件注入的注册表。 */
export const inject = ['skills']

/**
 * 把打包的长文档 skill 注册到 `ctx.skills`。
 *
 * @param {import('@deepseek-ai/cordis').Context} ctx - 插件上下文。
 */
export function apply(ctx) {
  ctx.skills.registerProvider(() => provider)
}
