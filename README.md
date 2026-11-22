# @deepwhale-cn/dsh-xlsx-diff

Excel 对比：两份 Excel 按关键列对比，列出新增、删除、修改（逐列 改前→改后）。openpyxl 实现；csv 版请用 dsh-csv-diff。

```bash
npx @deepwhale-cn/dsh-xlsx-diff --a 旧.xlsx --b 新.xlsx --key 单号 [--json] [--out 报告.md]
```

退出码：无差异 → 0（CI 可用）。
