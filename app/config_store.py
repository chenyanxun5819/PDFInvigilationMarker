from __future__ import annotations

import json
import os
import sys
from pathlib import Path

from .models import AppConfig, SeatBlockRecord, TeacherRecord
from .utils import normalize_seat_no

APP_DIR_NAME = ".pdf_invigilation_marker"
USER_CONFIG_NAME = "config.json"


def get_resource_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys._MEIPASS)  # type: ignore[attr-defined]
    return Path(__file__).resolve().parent.parent


def get_default_config_path() -> Path:
    return get_resource_root() / "data" / "default_config.json"


def get_user_data_dir() -> Path:
    app_data = os.getenv("APPDATA")
    if app_data:
        return Path(app_data) / APP_DIR_NAME
    return Path.home() / APP_DIR_NAME


def get_user_config_path() -> Path:
    return get_user_data_dir() / USER_CONFIG_NAME


def ensure_user_config() -> Path:
    user_dir = get_user_data_dir()
    user_dir.mkdir(parents=True, exist_ok=True)
    config_path = get_user_config_path()
    if not config_path.exists():
        default_path = get_default_config_path()
        if default_path.exists():
            config_path.write_text(default_path.read_text(encoding="utf-8"), encoding="utf-8")
    return config_path


def load_config(config_path: Path | None = None) -> AppConfig:
    if config_path is None:
        config_path = ensure_user_config()
    if not config_path.exists():
        raise FileNotFoundError(
            "找不到设定档。请先匯入老师名单 Excel，或把 default_config.json 放到 data 目录。"
        )

    payload = json.loads(config_path.read_text(encoding="utf-8"))
    teachers = tuple(
        TeacherRecord(
            teacher_id=str(item.get("teacher_id", "")).strip(),
            name=str(item.get("name", "")).strip(),
            initial=str(item.get("initial", "")).strip(),
            seat_no=normalize_seat_no(item.get("seat_no", "")),
            group=str(item.get("group", "")).strip(),
        )
        for item in payload.get("teachers", [])
        if str(item.get("initial", "")).strip()
    )
    seat_blocks = tuple(
        SeatBlockRecord(
            seat_no=normalize_seat_no(item.get("seat_no", "")),
            block=str(item.get("block", "")).strip().upper(),
        )
        for item in payload.get("seat_blocks", [])
        if str(item.get("seat_no", "")).strip()
    )
    if not teachers:
        raise ValueError("设定档内没有老师名单资料。")
    if not seat_blocks:
        raise ValueError("设定档内没有座位区块资料。")

    sort_settings = payload.get("sort_settings", {})
    ocr_settings = payload.get("ocr_settings", {})
    render_settings = payload.get("render_settings", {})

    return AppConfig(
        teachers=teachers,
        seat_blocks=seat_blocks,
        group_order=tuple(sort_settings.get("group_order", ["L", "R", "C"])),
        text_only_mode=bool(ocr_settings.get("text_only_mode", False)),
        ocr_zoom=float(ocr_settings.get("zoom", 3.0)),
        confidence_threshold=float(ocr_settings.get("confidence_threshold", 0.48)),
        label_anchor=str(render_settings.get("label_anchor", "bottom_left")).strip() or "bottom_left",
        label_x=float(render_settings.get("label_x", render_settings.get("label_margin", 18.0))),
        label_y=float(render_settings.get("label_y", render_settings.get("label_margin", 18.0))),
        label_margin=float(render_settings.get("label_margin", 18.0)),
        label_height=float(render_settings.get("label_height", 28.0)),
        label_font_size=float(render_settings.get("label_font_size", 14.0)),
        label_gap=float(render_settings.get("label_gap", 20.0)),
        star_scale=float(render_settings.get("star_scale", 4.95)),
        star_gap=float(render_settings.get("star_gap", 3.0)),
        star_y_offset=float(render_settings.get("star_y_offset", 0.0)),
        label_columns=tuple(float(item) for item in render_settings.get("label_columns", [24.0, 120.0, 320.0])),
        highlight_padding=float(render_settings.get("highlight_padding", 3.0)),
        highlight_fill_rgb=tuple(float(item) for item in render_settings.get("highlight_fill_rgb", [1.0, 0.95, 0.3])),
        highlight_border_rgb=tuple(float(item) for item in render_settings.get("highlight_border_rgb", [0.88, 0.2, 0.15])),
        label_font_path=str(render_settings.get("label_font_path", "")).strip(),
        source_path=config_path,
    )


def save_config_payload(payload: dict, output_path: Path | None = None) -> Path:
    target_path = output_path or get_user_config_path()
    target_path.parent.mkdir(parents=True, exist_ok=True)
    target_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return target_path


def load_existing_payload() -> dict:
    user_path = get_user_config_path()
    if user_path.exists():
        return json.loads(user_path.read_text(encoding="utf-8"))

    default_path = get_default_config_path()
    if default_path.exists():
        return json.loads(default_path.read_text(encoding="utf-8"))

    return {}


def import_excel_to_user_config(excel_path: Path) -> Path:
    from .excel_importer import import_config_from_excel

    payload = import_config_from_excel(excel_path, base_payload=load_existing_payload())
    return save_config_payload(payload)


def restore_default_config() -> Path:
    default_path = get_default_config_path()
    if not default_path.exists():
        raise FileNotFoundError("找不到内建的 default_config.json。")
    return save_config_payload(json.loads(default_path.read_text(encoding="utf-8")))


def update_render_settings(render_settings: dict[str, float | str | list[float]]) -> Path:
    config_path = ensure_user_config()
    payload = json.loads(config_path.read_text(encoding="utf-8"))
    payload.setdefault("render_settings", {})
    payload["render_settings"].update(render_settings)
    return save_config_payload(payload, config_path)


def update_ocr_settings(ocr_settings: dict[str, float | bool]) -> Path:
    config_path = ensure_user_config()
    payload = json.loads(config_path.read_text(encoding="utf-8"))
    payload.setdefault("ocr_settings", {})
    payload["ocr_settings"].update(ocr_settings)
    return save_config_payload(payload, config_path)
