"""
build_xlsx.py — 生成绿白主题、紧凑列结构的分镜脚本 xlsx

列结构（删除空白缓冲列，表头跨列合并显示）：
    A: 左侧缓冲 (3.0)
    B: 镜号 (4.0)
    C/D 合并: 拍摄地点 (6.0+6.0)
    E/F 合并: 画面内容 (8.0+8.0)
    G/H 合并: 镜头设计 (12.0+12.0)
    I/J/K 合并: 台词/旁白 (8.0+8.0+8.0)
    L: 时长 (8.0)
    M: 备注 (10.0)
    N: 道具 (8.0)
    O: 音效 (8.0)
    P/Q 合并: 参考画面 (8.0+8.0)

行结构：
    row 1: 缓冲
    row 2: 大标题"短视频拍摄脚本"
    row 3-7: 信息头（参考 image#3 三列一组）
    row 8-9: 缓冲
    row 10: 表头（合并显示）
    row 11+: 镜头行（绿白交替）
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.worksheet import Worksheet


# 列宽（宽格结构：避免内容被压缩）
COL_WIDTHS = {
    "A": 4.0,    # 左侧缓冲
    "B": 5.0,    # 镜号
    "C": 8.0,    # 拍摄地点左
    "D": 8.0,    # 拍摄地点右
    "E": 10.0,   # 画面内容左
    "F": 10.0,   # 画面内容右
    "G": 15.0,   # 镜头设计左
    "H": 15.0,   # 镜头设计右
    "I": 10.0,   # 台词/旁白左
    "J": 10.0,   # 台词/旁白中
    "K": 10.0,   # 台词/旁白右
    "L": 9.0,    # 时长
    "M": 12.0,   # 备注
    "N": 10.0,   # 道具
    "O": 10.0,   # 音效
    "P": 10.0,   # 参考画面左
    "Q": 10.0,   # 参考画面右
}

# 实际写入列（左上角单元格）
COL_SHOT_NO = 2    # B
COL_LOCATION = 3   # C（合并 C:D）
COL_ACTION = 5     # E（合并 E:F）
COL_CAMERA = 7     # G（合并 G:H）
COL_DIALOGUE = 9   # I（合并 I:K）
COL_DURATION = 12  # L
COL_NOTES = 13     # M
COL_PROPS = 14     # N
COL_SFX = 15       # O
COL_REFERENCE = 16 # P（合并 P:Q）

LAST_COL = 17  # Q

# 行
TITLE_ROW = 2
META_START = 3
META_END = 7
HEADER_ROW = 10
SHOTS_START = 11

FONT_FAMILY = "微软雅黑"
TITLE_FONT_SIZE = 24
META_FONT_SIZE = 11
HEADER_FONT_SIZE = 12
THIN_BORDER = Border(
    left=Side(style="thin", color="FFCCCCCC"),
    right=Side(style="thin", color="FFCCCCCC"),
    top=Side(style="thin", color="FFCCCCCC"),
    bottom=Side(style="thin", color="FFCCCCCC"),
)

CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT_TOP = Alignment(horizontal="left", vertical="top", wrap_text=True)


def _color(theme_color: str | None) -> str:
    """把 #RRGGBB 或 RRGGBB 统一成 FFRRGGBB"""
    c = (theme_color or "383838").lstrip("#").upper()
    if len(c) == 6:
        return f"FF{c}"
    return c


def _fills(theme_color: str | None):
    """返回各区域填充：标题/信息头标签、表头、幕标题、交替行1、交替行2

    默认一比一复刻《示例脚本 A（金标准）.xlsx》配色：
    - 标题/信息头/表头：深灰 #383838（白字）
    - 幕标题：浅绿 #D9EAD3（黑字）
    - 镜头行：白色 / #F3F3F3 交替
    """
    main = PatternFill("solid", fgColor=_color(theme_color))
    header = PatternFill("solid", fgColor=_color(theme_color))
    act = PatternFill("solid", fgColor="FFD9EAD3")     # 幕标题浅绿
    alt1 = PatternFill("solid", fgColor="FFFFFFFF")    # 白
    alt2 = PatternFill("solid", fgColor="FFF3F3F3")    # 浅灰
    return main, header, act, alt1, alt2


def _apply_column_widths(ws: Worksheet) -> None:
    for letter, w in COL_WIDTHS.items():
        ws.column_dimensions[letter].width = w


def _fill_row_bg(ws: Worksheet, row: int, fill: PatternFill, start_col: int = 1, end_col: int = LAST_COL) -> None:
    """给一行所有单元格统一填背景和边框"""
    for col in range(start_col, end_col + 1):
        c = ws.cell(row=row, column=col)
        c.fill = fill
        c.border = THIN_BORDER


def _write_title(ws: Worksheet, title_fill: PatternFill) -> None:
    """row 2: 大标题"""
    ws.row_dimensions[1].height = 6
    ws.row_dimensions[TITLE_ROW].height = 45
    # 先给整行填色
    _fill_row_bg(ws, TITLE_ROW, title_fill)
    # 合并并写入标题
    ws.merge_cells(start_row=TITLE_ROW, start_column=1, end_row=TITLE_ROW, end_column=LAST_COL)
    cell = ws.cell(row=TITLE_ROW, column=1, value="短视频拍摄脚本")
    cell.font = Font(name=FONT_FAMILY, size=TITLE_FONT_SIZE, bold=True, color="FFFFFFFF")
    cell.fill = title_fill
    cell.alignment = CENTER


def _write_meta_single(ws: Worksheet, row: int, label: str, value: str, main_fill: PatternFill) -> None:
    """单行信息头（视频名称/视频主题）"""
    # 标签 B-C
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=3)
    label_cell = ws.cell(row=row, column=2, value=label)
    label_cell.font = Font(name=FONT_FAMILY, size=META_FONT_SIZE, bold=True, color="FFFFFFFF")
    label_cell.fill = main_fill
    label_cell.alignment = CENTER
    label_cell.border = THIN_BORDER
    # 值 D-Q
    ws.merge_cells(start_row=row, start_column=4, end_row=row, end_column=LAST_COL)
    value_cell = ws.cell(row=row, column=4, value=value)
    value_cell.font = Font(name=FONT_FAMILY, size=META_FONT_SIZE, color="FF000000")
    value_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    value_cell.border = THIN_BORDER
    ws.row_dimensions[row].height = 32


def _write_meta_triple(ws: Worksheet, row: int, items: list[tuple[str, str]], main_fill: PatternFill) -> None:
    """三列一组信息头"""
    # 三组：[B-C 标签, D-F 值], [G-H 标签, I-K 值], [L-M 标签, N-Q 值]
    groups = [
        (2, 3, 4, 6),    # B-C label, D-F value
        (7, 8, 9, 11),   # G-H label, I-K value
        (12, 13, 14, 17) # L-M label, N-Q value
    ]
    for (label, value), (lc_start, lc_end, vc_start, vc_end) in zip(items, groups):
        # 标签
        ws.merge_cells(start_row=row, start_column=lc_start, end_row=row, end_column=lc_end)
        label_cell = ws.cell(row=row, column=lc_start, value=label)
        label_cell.font = Font(name=FONT_FAMILY, size=META_FONT_SIZE, bold=True, color="FFFFFFFF")
        label_cell.fill = main_fill
        label_cell.alignment = CENTER
        label_cell.border = THIN_BORDER
        # 值
        ws.merge_cells(start_row=row, start_column=vc_start, end_row=row, end_column=vc_end)
        value_cell = ws.cell(row=row, column=vc_start, value=value)
        value_cell.font = Font(name=FONT_FAMILY, size=META_FONT_SIZE, color="FF000000")
        value_cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        value_cell.border = THIN_BORDER
    ws.row_dimensions[row].height = 28


def _write_meta(ws: Worksheet, meta: dict, main_fill: PatternFill) -> None:
    """row 3-7: 信息头"""
    _write_meta_single(ws, META_START, "视频名称", meta.get("video_name", ""), main_fill)
    _write_meta_single(ws, META_START + 1, "视频主题", meta.get("video_topic", ""), main_fill)
    _write_meta_triple(ws, META_START + 2, [
        ("拍摄时间", meta.get("shoot_date", "")),
        ("截止时间", meta.get("deadline", "")),
        ("镜号数", str(meta.get("shot_count", ""))),
    ], main_fill)
    _write_meta_triple(ws, META_START + 3, [
        ("拍摄道具", meta.get("props", "")),
        ("主要产品", meta.get("main_product", "")),
        ("", ""),  # 第三组留空，保持结构对称
    ], main_fill)
    _write_meta_triple(ws, META_START + 4, [
        ("时间进度", meta.get("time_progress", "")),
        ("完成镜号", str(meta.get("completed_shots", ""))),
        ("完成进度", meta.get("completion_rate", "")),
    ], main_fill)


def _write_header(ws: Worksheet, header_fill: PatternFill) -> None:
    """row 10: 表头（合并显示）"""
    headers = [
        (2, 2, "镜号"),          # B
        (3, 4, "拍摄地点"),       # C:D
        (5, 6, "画面内容"),       # E:F
        (7, 8, "镜头设计"),       # G:H
        (9, 11, "台词/旁白"),     # I:K
        (12, 12, "时长"),         # L
        (13, 13, "备注"),         # M
        (14, 14, "道具"),         # N
        (15, 15, "音效"),         # O
        (16, 17, "参考画面"),      # P:Q
    ]
    for start_col, end_col, text in headers:
        if start_col != end_col:
            ws.merge_cells(start_row=HEADER_ROW, start_column=start_col, end_row=HEADER_ROW, end_column=end_col)
        cell = ws.cell(row=HEADER_ROW, column=start_col, value=text)
        cell.font = Font(name=FONT_FAMILY, size=HEADER_FONT_SIZE, bold=True, color="FFFFFFFF")
        cell.fill = header_fill
        cell.alignment = CENTER
        cell.border = THIN_BORDER
        # 合并范围其他单元格也填色加框
        for col in range(start_col + 1, end_col + 1):
            c = ws.cell(row=HEADER_ROW, column=col)
            c.fill = header_fill
            c.border = THIN_BORDER
    ws.row_dimensions[HEADER_ROW].height = 33


def _estimate_row_height(text: str | None) -> float:
    if not text:
        return 50.0
    n_chars = len(str(text))
    n_lines = max(1, (n_chars // 30) + 1)
    return min(260.0, max(50.0, 30.0 * n_lines))


def _write_act(ws: Worksheet, row: int, act_name: str, act_fill: PatternFill) -> int:
    """幕标题行"""
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=LAST_COL)
    cell = ws.cell(row=row, column=2, value=act_name)
    cell.font = Font(name=FONT_FAMILY, size=META_FONT_SIZE, bold=True, color="FF000000")
    cell.alignment = CENTER
    cell.fill = act_fill
    _fill_row_bg(ws, row, act_fill, start_col=2)
    ws.row_dimensions[row].height = 35
    return row + 1


def _write_shot(ws: Worksheet, row: int, shot: dict, row_fill: PatternFill) -> int:
    """镜头行（合并显示单元格）"""
    def g(*keys):
        """字段别名兼容：action/visual、sfx/audio、reference/ref、location/loc 都认"""
        for k in keys:
            v = shot.get(k)
            if v:
                return v
        return ""

    fields = [
        (COL_SHOT_NO, COL_SHOT_NO, shot.get("shot_no", "")),
        (COL_LOCATION, 4, g("location", "loc")),
        (COL_ACTION, 6, g("action", "visual", "vis")),
        (COL_CAMERA, 8, g("camera", "cam")),
        (COL_DIALOGUE, 11, g("dialogue", "dlg")),
        (COL_DURATION, COL_DURATION, g("duration", "dur")),
        (COL_NOTES, COL_NOTES, g("notes", "note")),
        (COL_PROPS, COL_PROPS, g("props", "prop")),
        (COL_SFX, COL_SFX, g("sfx", "audio", "aud")),
        (COL_REFERENCE, 17, g("reference", "ref")),
    ]

    for start_col, end_col, value in fields:
        if start_col != end_col:
            ws.merge_cells(start_row=row, start_column=start_col, end_row=row, end_column=end_col)
        cell = ws.cell(row=row, column=start_col, value=value)
        cell.font = Font(name=FONT_FAMILY, size=11)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN_BORDER
        # 合并范围内统一填色
        for col in range(start_col, end_col + 1):
            c = ws.cell(row=row, column=col)
            c.fill = row_fill
            c.border = THIN_BORDER

    # 镜号单独居中加粗
    c = ws.cell(row=row, column=COL_SHOT_NO)
    c.font = Font(name=FONT_FAMILY, size=11, bold=True)
    c.alignment = CENTER

    longest = max(
        _len(g("action", "visual", "vis")),
        _len(g("dialogue", "dlg")),
        _len(g("notes", "note")),
        _len(g("camera", "cam")),
    )
    ws.row_dimensions[row].height = _estimate_row_height("x" * longest)
    return row + 1


def _len(text) -> int:
    return len(str(text or ""))


def _apply_page_setup(ws: Worksheet) -> None:
    """打印设置：横向 + 适合一页 + A3，避免预览时被截断"""
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True
    ws.page_margins = openpyxl.worksheet.page.PageMargins(
        left=0.4, right=0.4, top=0.5, bottom=0.5, header=0.3, footer=0.3
    )


def build(output_path: str, meta: dict, acts: list[dict]) -> None:
    """主入口"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    main_fill, header_fill, act_fill, alt1, alt2 = _fills(meta.get("theme_color"))

    _apply_column_widths(ws)
    _write_title(ws, main_fill)
    _write_meta(ws, meta, main_fill)
    _write_header(ws, header_fill)

    current_row = SHOTS_START
    shot_index = 0
    for act in acts:
        act_name = act.get("name")
        if act_name:
            current_row = _write_act(ws, current_row, act_name, act_fill)
        for shot in act.get("shots", []):
            row_fill = alt1 if shot_index % 2 == 0 else alt2
            current_row = _write_shot(ws, current_row, shot, row_fill)
            shot_index += 1

    # 不冻结任何窗格，方便用户自由滚动
    ws.freeze_panes = None

    # 打印设置：横向 + 适合一页
    _apply_page_setup(ws)

    out = Path(output_path).expanduser()
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out))
    n_shots = sum(len(a.get("shots", [])) for a in acts)
    print(f"OK: {out} ({out.stat().st_size} bytes, {n_shots} 镜头)")


def _cli():
    ap = argparse.ArgumentParser(description="生成绿白主题分镜脚本 xlsx")
    ap.add_argument("--input", required=True, help="结构化 JSON 输入文件路径")
    ap.add_argument("--output", required=True, help="输出 xlsx 路径")
    args = ap.parse_args()
    with open(args.input, encoding="utf-8") as f:
        data = json.load(f)
    build(args.output, data.get("meta", {}), data.get("acts", []))


if __name__ == "__main__":
    _cli()
