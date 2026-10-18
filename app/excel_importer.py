from __future__ import annotations

from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

from .utils import normalize_seat_no

# ======= 栏位别名定义 (简体中文优先) =======
# 简体中文别名置于最前，确保在 map_headers() 中优先匹配

# 老师代号: 简体中文版本 "代号" 放在最前
TEACHER_INITIAL_ALIASES = ("代号Initial", "代号 Initial", "代号", "Initial", "INITIAL")

# 教师姓名: 简体中文版本 "教师姓名" 和 "姓名" 放在最前
TEACHER_NAME_ALIASES = ("教师姓名 Name", "教師姓名 Name", "教师姓名Name", "教師姓名Name", "教师姓名", "姓名", "Name")

# 座号: 简体中文版本 "座号" 放在最前
TEACHER_SEAT_ALIASES = ("座号", "座位号", "Seat", "Seat No", "Seat Number", "座號", "座位號")

# 教职员编号: 简体中文版本 "教职员编号" 放在最前
TEACHER_ID_ALIASES = ("教职员编号", "Teacher ID", "ID", "教職員編號")

# 组别: 简体中文版本 "组别" 放在最前
TEACHER_GROUP_ALIASES = ("组别", "Group", "組別")

# 区块: 简体中文版本 "区块" 放在最前
SEAT_BLOCK_ALIASES = ("区块", "Block", "區塊")

# 座位号: 简体中文版本 "座位号" 放在最前
SEAT_NO_ALIASES = ("座位号", "座位", "Seat", "Seat No", "Seat Number", "座位號", "座號")

INVIGILATOR_CODE_WORKBOOK_NAMES = (
    "监考代号名单.xlsx",
    "監考代號名單.xlsx",
)


def import_config_from_excel(excel_path: Path, base_payload: dict | None = None) -> dict:
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
    seat_columns: dict[str, int] | None = None
    if seat_sheet is not None:
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

    merge_invigilator_code_workbook(excel_path, teachers)

    seat_blocks = []
    if seat_sheet is not None and seat_columns is not None:
        for row in seat_sheet.iter_rows(min_row=2, values_only=True):
            seat_no = normalize_seat_no(row[seat_columns["seat_no"]])
            block = normalize_text(row[seat_columns["block"]]).upper()
            if not seat_no or not block:
                continue
            seat_blocks.append({
                "seat_no": seat_no,
                "block": block,
            })
    elif base_payload:
        seat_blocks = list(base_payload.get("seat_blocks", []))

    if not teachers:
        raise ValueError("Excel 中找不到任何老师名单资料。")
    if not seat_blocks:
        raise ValueError("Excel 中找不到任何座位区块资料，且目前设定档也没有可沿用的座位区块资料。")

    sort_settings = dict(base_payload.get("sort_settings", {})) if base_payload else {"group_order": ["L", "R", "C"]}
    existing_ocr_settings = dict(base_payload.get("ocr_settings", {})) if base_payload else {}
    ocr_settings = {
        "zoom": max(float(existing_ocr_settings.get("zoom", 3.0)), 3.0),
        "confidence_threshold": min(float(existing_ocr_settings.get("confidence_threshold", 0.48)), 0.48),
        "text_only_mode": False,
    }
    render_settings = dict(base_payload.get("render_settings", {})) if base_payload else {
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
    }

    metadata = dict(base_payload.get("metadata", {})) if base_payload else {}
    invigilator_workbook_path = find_invigilator_code_workbook(excel_path)
    metadata.update({
        "generated_from_excel": str(excel_path),
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "seat_blocks_source": "excel" if seat_sheet is not None else "existing_config",
        "invigilator_codes_source": str(invigilator_workbook_path) if invigilator_workbook_path else "",
    })

    return {
        "metadata": metadata,
        "sort_settings": sort_settings,
        "ocr_settings": ocr_settings,
        "render_settings": render_settings,
        "teachers": teachers,
        "seat_blocks": seat_blocks,
    }


def find_teacher_sheet(workbook):
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        header_values = [normalize_header(cell.value) for cell in sheet[1]]
        if contains_alias(header_values, TEACHER_INITIAL_ALIASES) and contains_alias(header_values, TEACHER_NAME_ALIASES):
            return sheet
    raise ValueError("找不到包含老师名单欄位的工作表。")


def find_seat_sheet(workbook):
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        header_values = [normalize_header(cell.value) for cell in sheet[1]]
        if contains_alias(header_values, SEAT_BLOCK_ALIASES) and contains_alias(header_values, SEAT_NO_ALIASES):
            return sheet
    return None


def map_headers(header_row, aliases: dict[str, tuple[str, ...]]) -> dict[str, int]:
    headers = [normalize_header(cell.value) for cell in header_row]
    mapped: dict[str, int] = {}
    for logical_name, options in aliases.items():
        for index, header in enumerate(headers):
            if any(normalize_header(option) == header for option in options):
                mapped[logical_name] = index
                break
        if logical_name not in mapped:
            raise ValueError(f"缺少必要栏位：{options[0]}")
    return mapped


def contains_alias(headers: list[str], aliases: tuple[str, ...]) -> bool:
    alias_set = {normalize_header(item) for item in aliases}
    return any(header in alias_set for header in headers)


def normalize_header(value: object) -> str:
    text = str(value or "").strip().lower().replace("\u3000", " ")
    return "".join(text.split())


def normalize_text(value: object) -> str:
    return str(value or "").strip()


def normalize_person_name(value: object) -> str:
    return "".join(str(value or "").strip().split())


def find_invigilator_code_workbook(excel_path: Path) -> Path | None:
    for file_name in INVIGILATOR_CODE_WORKBOOK_NAMES:
        candidate = excel_path.parent / file_name
        if candidate.exists() and candidate != excel_path:
            return candidate
    return None


def merge_invigilator_code_workbook(excel_path: Path, teachers: list[dict]) -> None:
    workbook_path = find_invigilator_code_workbook(excel_path)
    if workbook_path is None:
        return

    workbook = load_workbook(workbook_path, data_only=True)
    sheet = workbook[workbook.sheetnames[0]]
    columns = map_headers(sheet[1], {
        "initial": TEACHER_INITIAL_ALIASES,
        "name": TEACHER_NAME_ALIASES,
        "teacher_id": TEACHER_ID_ALIASES,
        "group": TEACHER_GROUP_ALIASES,
    })

    teacher_by_id = {
        normalize_text(item.get("teacher_id", "")): item
        for item in teachers
        if normalize_text(item.get("teacher_id", ""))
    }
    teacher_by_name = {
        normalize_person_name(item.get("name", "")): item
        for item in teachers
        if normalize_person_name(item.get("name", ""))
    }

    for row in sheet.iter_rows(min_row=2, values_only=True):
        initial = normalize_text(row[columns["initial"]])
        name = normalize_text(row[columns["name"]])
        teacher_id = normalize_text(row[columns["teacher_id"]])
        group = normalize_text(row[columns["group"]])
        if not initial or not name:
            continue

        teacher_record = None
        if teacher_id:
            teacher_record = teacher_by_id.get(teacher_id)
        if teacher_record is None:
            teacher_record = teacher_by_name.get(normalize_person_name(name))

        if teacher_record is not None:
            teacher_record["initial"] = initial
            if name:
                teacher_record["name"] = name
            if group and not teacher_record.get("group"):
                teacher_record["group"] = group
            if teacher_id and not teacher_record.get("teacher_id"):
                teacher_record["teacher_id"] = teacher_id
            continue

        supplemental = {
            "teacher_id": teacher_id,
            "name": name,
            "initial": initial,
            "seat_no": "",
            "group": group,
        }
        teachers.append(supplemental)
        if teacher_id:
            teacher_by_id[teacher_id] = supplemental
        teacher_by_name[normalize_person_name(name)] = supplemental
