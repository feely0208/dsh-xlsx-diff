#!/usr/bin/env python3
"""两份 Excel 按关键列对比：新增 / 删除 / 修改（逐列差异）。输出 JSON。"""
import json, sys
from openpyxl import load_workbook

def read_rows(path, sheet, key_cols):
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet] if sheet else wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return [], []
    header = [str(c).strip() if c is not None else f"列{i+1}" for i, c in enumerate(rows[0])]
    key_idx = []
    for k in key_cols:
        if k not in header:
            print(json.dumps({"ok": False, "error": f"关键列 {k} 不在表头 {header}"}))
            sys.exit(2)
        key_idx.append(header.index(k))
    data, dup = {}, 0
    for r in rows[1:]:
        if all(c is None for c in r):
            continue
        key = tuple(str(r[i]) if i < len(r) and r[i] is not None else "" for i in key_idx)
        if key in data:
            dup += 1
            continue  # 重复键取第一行
        data[key] = {header[i]: (r[i] if i < len(r) else None) for i in range(len(header))}
    return data, dup

def fmt(v):
    return "" if v is None else str(v)

def main():
    a_path, b_path = sys.argv[1], sys.argv[2]
    key_cols = [k.strip() for k in sys.argv[3].split(",")]
    sheet_a = sys.argv[4] if len(sys.argv) > 4 and sys.argv[4] != "-" else None
    sheet_b = sys.argv[5] if len(sys.argv) > 5 and sys.argv[5] != "-" else sheet_a
    a, dup_a = read_rows(a_path, sheet_a, key_cols)
    b, dup_b = read_rows(b_path, sheet_b, key_cols)
    added, removed, changed = [], [], []
    for k in b:
        if k not in a:
            added.append({key_cols[i]: k[i] for i in range(len(k))})
    for k in a:
        if k not in b:
            removed.append({key_cols[i]: k[i] for i in range(len(k))})
    for k in (a.keys() & b.keys()):
        cols = []
        for col in a[k]:
            if fmt(a[k][col]) != fmt(b[k][col]):
                cols.append({"列": col, "改前": fmt(a[k][col]), "改后": fmt(b[k][col])})
        if cols:
            changed.append({"key": {key_cols[i]: k[i] for i in range(len(k))}, "差异": cols})
    print(json.dumps({"ok": True, "key": key_cols, "counts": {
        "a_rows": len(a), "b_rows": len(b), "added": len(added),
        "removed": len(removed), "changed": len(changed), "dup_keys": dup_a + dup_b},
        "added": added, "removed": removed, "changed": changed}, ensure_ascii=False))

main()
