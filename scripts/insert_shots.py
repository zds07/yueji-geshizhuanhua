#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
insert_shots.py — 在已有分镜 xlsx 的指定镜号之后插入新镜，并自动重排后续镜号。

适用场景：用户对已生成的分镜表做过手工修改（清空某些列、改地点、加备注），
此时**不要用 build_xlsx.py 整表重建**（会覆盖用户修改），改用本脚本原地插入。

用法：
  python insert_shots.py --file <xlsx> --after 16 --shots new_shots.json [--delta 3]

shots JSON 格式（数组）：
[
  {
    "location": "果园",
    "visual":   "画面内容",
    "camera":   "镜头设计",
    "dialogue": "台词/旁白",
    "duration": "",
    "notes":    "",
    "props":    "",
    "audio":    "",
    "ref":      ""
  }
]

行为：
1. 备份原文件为 <name>.bak.xlsx（确认无误后手动删）
2. 在镜号 after 所在行之后插入 N 行，复制该行样式（字体/边框/填充/换行/行高）
3. 重建合并单元格：min_row >= 插入点的区域整体下移 N 行；新行补 C:D E:F G:H I:K P:Q
4. 写入新镜内容
5. 后续所有数字镜号 +N
6. 更新信息头「镜号数」
7. 自动修正台词/备注里形如「配 19 镜」「按 19 镜」的旧镜号引用（+N）

注意：本脚本不含 openpyxl 依赖之外的东西，用隔离环境 python 跑。
"""
import argparse
import json
import os
import re
import shutil
import sys

import openpyxl
from copy import copy
from openpyxl.utils import range_boundaries, get_column_letter

# B=镜号 C=拍摄地点 E=画面内容 G=镜头设计 I=台词 L=时长 M=备注 N=道具 O=音效 P=参考画面
COL = {"no": 2, "loc": 3, "vis": 5, "cam": 7, "dlg": 9,
       "dur": 12, "note": 13, "prop": 14, "aud": 15, "ref": 16}

MERGE_TPL = ["C{r}:D{r}", "E{r}:F{r}", "G{r}:H{r}", "I{r}:K{r}", "P{r}:Q{r}"]


def shift_range(rg, delta, threshold):
    mn_c, mn_r, mx_c, mx_r = range_boundaries(rg)
    if mn_r >= threshold:
        mn_r += delta
        mx_r += delta
    return f"{get_column_letter(mn_c)}{mn_r}:{get_column_letter(mx_c)}{mx_r}"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--file", required=True)
    ap.add_argument("--after", required=True, type=int, help="插在哪个镜号之后")
    ap.add_argument("--shots", required=True, help="新镜 JSON 文件")
    ap.add_argument("--no-backup", action="store_true")
    args = ap.parse_args()

    path = os.path.abspath(args.file)
    shots = json.load(open(args.shots, encoding="utf-8"))
    n = len(shots)

    if not args.no_backup:
        shutil.copy(path, path.replace(".xlsx", ".bak.xlsx"))

    wb = openpyxl.load_workbook(path)
    ws = wb.active

    src_row = None
    for r in range(11, ws.max_row + 1):
        v = ws.cell(row=r, column=COL["no"]).value
        if v is not None and str(v).strip() == str(args.after):
            src_row = r
            break
    if src_row is None:
        print(f"[x] 找不到镜号 {args.after}", file=sys.stderr)
        sys.exit(1)

    ins_at = src_row + 1

    # 1) 拆掉所有合并（openpyxl 不会随 insert_rows 平移）
    old_ranges = [str(x) for x in ws.merged_cells.ranges]
    for rg in old_ranges:
        ws.unmerge_cells(rg)

    # 2) 插入行
    ws.insert_rows(ins_at, n)

    # 3) 复制源行样式
    for i in range(n):
        r = ins_at + i
        ws.row_dimensions[r].height = ws.row_dimensions[src_row].height
        for c in range(1, ws.max_column + 1):
            ws.cell(row=r, column=c)._style = copy(ws.cell(row=src_row, column=c)._style)

    # 4) 重建合并
    new_ranges = [shift_range(rg, n, ins_at) for rg in old_ranges]
    for i in range(n):
        r = ins_at + i
        new_ranges += [t.format(r=r) for t in MERGE_TPL]
    for rg in new_ranges:
        try:
            ws.merge_cells(rg)
        except Exception as e:
            print(f"[!] 合并失败 {rg}: {e}", file=sys.stderr)

    # 5) 写新镜（字段名兼容长短两套：location/loc、visual/vis、dialogue/dlg…）
    ALIAS = {"loc": "location", "vis": "visual", "cam": "camera",
             "dlg": "dialogue", "dur": "duration", "note": "notes",
             "prop": "props", "aud": "audio", "ref": "ref"}
    for i, s in enumerate(shots):
        r = ins_at + i
        ws.cell(row=r, column=COL["no"]).value = args.after + 1 + i
        for key, col in COL.items():
            if key == "no":
                continue
            v = s.get(key) or s.get(ALIAS[key]) or ""
            ws.cell(row=r, column=col).value = v or None

    # 6) 后续镜号 +n
    start = ins_at + n
    last = 0
    for r in range(start, ws.max_row + 1):
        v = ws.cell(row=r, column=COL["no"]).value
        if v is None:
            continue
        sv = str(v).strip()
        if sv.isdigit():
            nv = int(sv) + n
            ws.cell(row=r, column=COL["no"]).value = nv
            last = max(last, nv)

    # 7) 修正 B-roll 引用（配/按 XX 镜）与更新信息头镜号数
    def bump(m):
        return f"{m.group(1)} {int(m.group(2)) + n} 镜"

    fixed = 0
    for r in range(start, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(row=r, column=c)
            if isinstance(cell.value, str):
                nv = re.sub(r"(配|按)\s*(\d+)\s*镜", bump, cell.value)
                if nv != cell.value:
                    cell.value = nv
                    fixed += 1

    for r in range(1, 12):
        for c in range(1, ws.max_column + 1):
            if str(ws.cell(row=r, column=c).value).strip() == "镜号数":
                ws.cell(row=r, column=c + 2).value = str(last)

    wb.save(path)
    print(f"[OK] 在镜号 {args.after} 后插入 {n} 镜，总镜号 {last}，修正引用 {fixed} 处")


if __name__ == "__main__":
    main()
