from __future__ import annotations

from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

from .utils import normalize_seat_no


TEACHER_INITIAL_ALIASES = ("代号 Initial", "Initial", "代号", "INITIAL")
TEACHER_NAME_ALIASES = ("教师姓名 Name", "教師姓名 Name", "姓名", "Name", "教师姓名")
TEACHER_SEAT_ALIASES = ("座号", "座號", "Seat", "Seat No", "Seat Number")
TEACHER_ID_ALIASES = ("教职员编号", "教職員編號", "Teacher ID", "ID")
TEACHER_GROUP_ALIASES = ("組別", "组别", "Group")
SEAT_BLOCK_ALIASES = ("区块", "區塊", "Block")
SEAT_NO_ALIASES = ("座位号", "座位號", "Seat", "Seat No", "Seat Number")


def import_config_from_excel(excel_path: Path) -> dict:
    workbook = load_workbook(excel_path, data_only=True)
    teacher_sheet = find_teacher_sheet(workbook)
    seat_sheet = find_seat_sheet(workbook)

    teacher_columns = map_headers(teacher_sheet[1], {
        "initial": TEACHER_INITIAL_ALIASES,
        "name": TEACHER_NAME_ALIASES,
        "seat_no": TEACHER_SEAT_ALIASES,
        "teacher_id": TEACHER_ID_ALIASES,
        "group": TEACHER_GROUP_ALIASES,
    })
    seat_columns = map_headers(seat_sheet[1], {
        "block": SEAT_BLOCK_ALIASES,
        "seat_no": SEAT_NO_ALIASES,
    })

    teachers = []
    for row in teacher_sheet.iter_rows(min_row=2, values_only=True):
        initial = normalize_text(row[teacher_columns["initial"]])
        if not initial:
            continue
        teachers.append({
            "teacher_id": normalize_text(row[teacher_columns["teacher_id"]]),
            "name": normalize_text(row[teacher_columns["name"]]),
            "initial": initial,
            "seat_no": normalize_seat_no(row[teacher_columns["seat_no"]]),
            "group": normalize_text(row[teacher_columns["group"]]),
        })

    seat_blocks = []
    for row in seat_sheet.iter_rows(min_row=2, values_only=True):
        seat_no = normalize_seat_no(row[seat_columns["seat_no"]])
        block = normalize_text(row[seat_columns["block"]]).upper()
        if not seat_no or not block:
            continue
        seat_blocks.append({
            "seat_no": seat_no,
            "block": block,
        })

    if not teachers:
        raise ValueError("Excel 中找不到任何老师名单资料。")
    if not seat_blocks:
        raise ValueError("Excel 中找不到任何座位区块资料。")

    return {
        "metadata": {
            "generated_from_excel": str(excel_path),
            "generated_at": datetime.now().isoformat(timespec="seconds"),
        },
        "sort_settings": {
            "group_order": ["L", "R", "C"],
        },
        "ocr_settings": {
            "zoom": 2.0,
            "confidence_threshold": 0.55,
            "text_only_mode": False,
        },
        "render_settings": {
            "label_anchor": "bottom_left",
            "label_x": 18.0,
            "label_y": 18.0,
            "label_margin": 18.0,
            "label_height": 28.0,
            "label_font_size": 14.0,
            "label_gap": 20.0,
            "star_scale": 4.95,
            "star_gap": 3.0,
            "star_y_offset": 0.0,
            "label_columns": [24.0, 120.0, 320.0],
            "highlight_padding": 3.0,
            "highlight_fill_rgb": [1.0, 0.95, 0.3],
            "highlight_border_rgb": [0.88, 0.2, 0.15],
            "label_font_path": "",
        },
        "teachers": teachers,
        "seat_blocks": seat_blocks,
    }


def find_teacher_sheet(workbook):
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        header_values = [normalize_text(cell.value) for cell in sheet[1]]
        if contains_alias(header_values, TEACHER_INITIAL_ALIASES) and contains_alias(header_values, TEACHER_NAME_ALIASES):
            return sheet
    raise ValueError("找不到包含老师名单欄位的工作表。")


def find_seat_sheet(workbook):
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        header_values = [normalize_text(cell.value) for cell in sheet[1]]
        if contains_alias(header_values, SEAT_BLOCK_ALIASES) and contains_alias(header_values, SEAT_NO_ALIASES):
            return sheet
    raise ValueError("找不到包含座位区块欄位的工作表。")


def map_headers(header_row, aliases: dict[str, tuple[str, ...]]) -> dict[str, int]:
    headers = [normalize_text(cell.value) for cell in header_row]
    mapped: dict[str, int] = {}
    for logical_name, options in aliases.items():
        for index, header in enumerate(headers):
            if any(normalize_text(option) == header for option in options):
                mapped[logical_name] = index
                break
        if logical_name not in mapped:
            raise ValueError(f"缺少必要栏位：{options[0]}")
    return mapped


def contains_alias(headers: list[str], aliases: tuple[str, ...]) -> bool:
    alias_set = {normalize_text(item) for item in aliases}
    return any(header in alias_set for header in headers)


def normalize_text(value: object) -> str:
    return str(value or "").strip()
