# Excel 对比（dsh-xlsx-diff）

**什么时候用**：两份 Excel（账单/订单/名单）按关键列对比找差异。
用户会说：「这两份表差在哪」「帮我对比新旧两版报价单」。

## 怎么用

```bash
<cli 路径> --a 旧.xlsx --b 新.xlsx --key 单号 [--sheet 表名] [--json] [--out 报告.md]
```

## 输出

新增 / 删除（整行）+ 修改（逐列 改前→差后）。重复键取第一行并提示数量。
退出码：无差异 → 0；有差异 → 1。

## 边界

依赖 openpyxl（pip3 install openpyxl）；只读第一个工作表（--sheet 指定）；公式按计算值对比（data_only）。
