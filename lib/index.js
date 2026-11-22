/** dsh-xlsx-diff —— Excel 对比 · 库入口（供测试与复用） */
export const NAME = 'dsh-xlsx-diff';

// ── DSH 插件接入（方案 A 薄壳：注册 skill 与 CLI 入口；业务逻辑仍在 lib/bin）──
import { createPlugin, pkgRootFrom } from '@deepwhale-cn/dsh-guard-core/plugin';

export const { name, inject, apply } = createPlugin({
  name: 'xlsx-diff',
  pkgRoot: pkgRootFrom(import.meta.url),
  skillFile: 'skill/xlsx-diff.md',
  binFile: 'bin/cli.mjs',
  description: "两份 Excel 按关键列对比，列出新增、删除、修改（逐列差异）；openpyxl 实现",
  whenToUse: "两份 Excel 按关键列对比找差异。用户会说：这两份表差在哪、新旧报价单帮我对比。",
});
