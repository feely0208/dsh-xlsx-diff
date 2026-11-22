#!/usr/bin/env node
/** dsh-xlsx-diff —— Excel 对比（openpyxl）。--a --b --key 列[,列] [--sheet] [--json] [--out 报告.md] */
import { spawnSync } from 'node:child_process';
import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { join, dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const args = process.argv.slice(2);
const has = (k) => args.includes('--' + k);
const val = (k, d = '') => {
  const i = args.indexOf('--' + k);
  return i >= 0 && args[i + 1] && !args[i + 1].startsWith('--') ? args[i + 1] : d;
};

if (has('help') || !has('a') || !has('b') || !has('key')) {
  console.log(`dsh-xlsx-diff —— Excel 对比

用法：dsh-xlsx-diff --a 旧.xlsx --b 新.xlsx --key 单号[,品名] [--sheet 表名] [--json] [--out 报告.md]
按关键列对比两份 Excel：新增 / 删除 / 修改（逐列列出差前→差后）。
重复键取第一行并在结果里提示；依赖 openpyxl（pip3 install openpyxl）。`);
  process.exit(2);
}

const py = join(dirname(fileURLToPath(import.meta.url)), '..', 'lib', 'diff.py');
const r = spawnSync('python3', [py, resolve(val('a')), resolve(val('b')), val('key'), val('sheet', '-'), val('sheet', '-')], { encoding: 'utf-8', timeout: 120_000, maxBuffer: 64 * 1024 * 1024 });
let j;
try { j = JSON.parse(r.stdout.trim().split('\n').filter((l) => l.startsWith('{')).pop()); }
catch { console.error('❌ python3/openpyxl 执行失败：' + (r.stderr || r.stdout || '').slice(0, 300)); process.exit(2); }
if (j.ok === false) { console.error('❌ ' + j.error); process.exit(2); }

if (has('json')) { process.stdout.write(JSON.stringify(j, null, 2) + '\n', () => process.exit(0));
} else {
const c = j.counts;
console.log(`\n  深鲸 · Excel 对比（关键列：${j.key.join(', ')}）\n  ${'─'.repeat(56)}`);
console.log(`  旧表 ${c.a_rows} 行 · 新表 ${c.b_rows} 行 · 新增 ${c.added} · 删除 ${c.removed} · 修改 ${c.changed}${c.dup_keys ? ` · 重复键跳过 ${c.dup_keys}` : ''}\n`);
for (const a of j.added) console.log(`  ➕ 新增：${JSON.stringify(a)}`);
for (const d of j.removed) console.log(`  ➖ 删除：${JSON.stringify(d)}`);
for (const ch of j.changed) {
  console.log(`  ✏️  修改：${JSON.stringify(ch.key)}`);
  for (const d of ch['差异']) console.log(`      ${d['列']}：${d['改前']} → ${d['改后']}`);
}
if (has('out')) {
  mkdirSync(dirname(resolve(val('out'))), { recursive: true });
  const L = ['# Excel 对比报告', '', `- 关键列：${j.key.join(', ')}；旧表 ${c.a_rows} 行 / 新表 ${c.b_rows} 行`, `- 新增 ${c.added} · 删除 ${c.removed} · 修改 ${c.changed}`, ''];
  for (const a of j.added) L.push(`- ➕ 新增：\`${JSON.stringify(a)}\``);
  for (const d of j.removed) L.push(`- ➖ 删除：\`${JSON.stringify(d)}\``);
  for (const ch of j.changed) L.push(`- ✏️  修改 \`${JSON.stringify(ch.key)}\`：` + ch['差异'].map((d) => `${d['列']} ${d['改前']}→${d['改后']}`).join('；'));
  writeFileSync(resolve(val('out')), L.join('\n') + '\n', 'utf-8');
  console.log(`\n报告已写入：${resolve(val('out'))}`);
}
process.exit((c.added + c.removed + c.changed) === 0 ? 0 : 1);
}
