# longform-structured-doc

longform-structured-doc 的 DeepSeek Harness (DSH) 插件：把**大规模长文档生成 skill**（v2.4.2，Mode A/B/C 三模式流水线）打包成 DSH skill provider，让任意会话可直接加载使用。

## 它是什么

这个 skill 专为**一次性成稿数十至上百页、数万至十余万字**的技术/公文文档设计（方案、报告、标书、申报材料等，最终交付 docx），核心能力：

- **三模式路由**：Mode A 需求/规范驱动（目录逐条对齐招标评审条目）、Mode B 对标改写（基于现有方案、不照抄）、Mode C 调研原创（自行搜集、论断必有出处）
- **完整流水线**（执行入口 `orchestrator.md`，六个阶段 + 闸门）：素材建卡 → 结构生成 → 分章撰写（writer → 反抄袭审计 → rewriter → 政策检查 → 终稿）→ 组装校验 → 终检
- **硬纪律**：每章最低字数下限、重点章节加权深化、空缺符（【待补充】/【待确认】/【需核实】）不编造、数值回源核查、括号数据清洗、去AI味公文润色
- **配套资源**：5 个角色 prompt 模板（`templates/`）+ 2 个 python 工具（`scripts/`：`docs_to_md.py` 拆源文档、`assemble_docx.py` 组装合规 docx）

本插件只做一件事：向 `ctx.skills` 注册 provider 发布这套 skill 树。**纯 skill 形态，无工具、无 MCP、无需配置**——模型按 orchestrator 编排调用 python 脚本执行。

## 安装

### 方式 A：GitHub（已发布）

```bash
dsh plugin --profile web add github:lzh8987/longform-structured-doc
```

或在 profile 的 `package.json` 里加依赖与 bundle 行，然后 `pnpm install` 并重启：

```jsonc
{
  "dependencies": { "longform-structured-doc": "github:lzh8987/longform-structured-doc" },
  "dsh": { "profile": { "bundles": [ "...", "longform-structured-doc" ] } }
}
```

### 方式 B：本地目录

```bash
dsh plugin add "./longform-structured-doc"   # 指向插件包目录
```

依赖：peer 依赖 `@deepseek-ai/cordis` DSH 自带；python 侧按 SKILL.md 说明装 `python-docx` / `pypdf` / `xlrd` 等（缺库时脚本会给出明确提示）。

## 使用

安装后新开会话，skill 目录里会出现 `longform-structured-doc`（provider 同名）。模型按 skill 触发词自动路由（"写一份XX建设方案" / "按招标要求写技术标" / "参照XX方案但不能抄" / "调研一下XX写份报告"…），加载 `SKILL.md` + `orchestrator.md` 后按六阶段推进，各阶段产物落盘到工作目录的 `素材库/`、`章节草稿/`、`交付物/`。

## 结构

```
longform-structured-doc/
├── lib/index.js                    # 插件入口：skill provider（零依赖）
├── cordis.patch.yml                # bundle 补丁（向 profile 插入本插件行）
├── skills/longform-structured-doc/
│   ├── SKILL.md                    # 方法与纪律参考（v2.4.2）
│   ├── orchestrator.md             # 执行入口（六阶段 + 闸门）
│   ├── templates/                  # 五个角色 prompt
│   └── scripts/                    # docs_to_md.py / assemble_docx.py
├── test/check.mjs                  # 冒烟测试
├── test/installed-copy-load.mjs    # 安装副本加载测试（复验用）
└── package.json
```

## 复验（Verification）

拿到代码后，按以下步骤独立复验"安装链路"是否正常（零依赖，只需 Node.js ≥ 18）：

```bash
# 1. 冒烟测试：直接在源码目录验证插件入口与打包的 skill 资源树
node test/check.mjs
#    → 末尾输出 SMOKE TEST PASSED

# 2. 真实安装进一个 profile（以 web 为例；也可用独立测试 profile）
dsh plugin --profile web add "file:<本仓库绝对路径>"
#    → pnpm 输出 longform-structured-doc 0.1.0 已加入 dependencies

# 3. 验证 bundle 组合（patch 层应出现本插件行）
dsh --profile web --dump-config
#    → 输出含: # == longform-structured-doc / - id: longform-structured-doc

# 4. 从安装副本（profile 的 node_modules/.pnpm 存储）真实加载，验证
#    provider 注册、skill list/get、以及 resourceBase 下 8 个资源均可访问
node test/installed-copy-load.mjs web      # 默认就是 web，可省略参数
#    → 末尾输出 INSTALLED-COPY LOAD TEST PASSED
```

预期：四个步骤全部通过（exit 0）。其中第 4 步的 `resolved entry` 必须指向
`<profile>/node_modules/...` 路径——这证明加载的是安装副本而非源码目录。
Windows 上若 `dsh` 不在 PATH，用 `node <dsh安装>/node_modules/@deepseek-ai/dsh/lib/bin.js` 或完整路径调用。

## 更新 skill 内容

skill 本体（SKILL.md / orchestrator.md / templates / scripts）改完后，把改动同步到 `skills/longform-structured-doc/` 对应文件，重启宿主即生效（安装是 `file:` 链接，内容直接跟随）。

## License

MIT。
