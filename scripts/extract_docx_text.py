#!/usr/bin/env python3
"""
extract_docx_text.py — 零依赖读取 docx 全文（段落 + 表格）

为什么不用 parse_brief.py 的 python-docx 方案：
    用户默认 Python 环境通常没装 python-docx，装依赖会污染环境。
    本脚本只用标准库（zipfile + ElementTree），任何 Python 3.8+ 都能跑。

用法：
    python3 extract_docx_text.py <原稿.docx> [-o 输出.txt]

输出格式：
    [P] 段落文本
    [T] 表格行（单元格用 | 分隔）
"""

from __future__ import annotations

import argparse
import sys
import zipfile
from xml.etree import ElementTree as ET

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def para_text(p: ET.Element) -> str:
    """抽出单个段落的文本，保留 tab 与换行"""
    parts: list[str] = []
    for node in p.iter():
        tag = node.tag.split("}")[1] if "}" in node.tag else node.tag
        if tag == "t":
            parts.append(node.text or "")
        elif tag == "tab":
            parts.append("\t")
        elif tag == "br":
            parts.append("\n")
    return "".join(parts)


def extract(path: str) -> list[tuple[str, str]]:
    """返回 [(类型, 文本)]，类型为 'P'（段落）或 'T'（表格行）"""
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml").decode("utf-8")
    root = ET.fromstring(xml)
    body = root.find("w:body", NS)
    if body is None:
        return []

    out: list[tuple[str, str]] = []
    for child in body:
        tag = child.tag.split("}")[1]
        if tag == "p":
            out.append(("P", para_text(child)))
        elif tag == "tbl":
            for tr in child.findall("w:tr", NS):
                cells = []
                for tc in tr.findall("w:tc", NS):
                    text = " ".join(para_text(p) for p in tc.findall("w:p", NS))
                    cells.append(text.strip())
                out.append(("T", " | ".join(cells)))
    return out


def main() -> None:
    ap = argparse.ArgumentParser(description="零依赖读取 docx 全文")
    ap.add_argument("input", help="docx 路径")
    ap.add_argument("-o", "--output", help="输出文件，缺省打印到 stdout")
    args = ap.parse_args()

    items = extract(args.input)
    lines = [f"[{kind}] {text}" for kind, text in items if text.strip()]
    content = "\n".join(lines)

    if args.output:
        with open(args.output, "w", encoding="utf-8") as f:
            f.write(content + "\n")
        print(f"OK: {args.output} ({len(lines)} 行)")
    else:
        print(content)


if __name__ == "__main__":
    main()
