from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path


@dataclass(frozen=True)
class TeacherRecord:
    teacher_id: str
    name: str
    initial: str
    seat_no: str
    group: str = ""
    unresolved: bool = False


@dataclass(frozen=True)
class SeatBlockRecord:
    seat_no: str
    block: str


@dataclass(frozen=True)
class MatchOccurrence:
    page_index: int
    rect: tuple[float, float, float, float]
    recognized_text: str
    confidence: float
    detection_source: str = "text"


@dataclass(frozen=True)
class TeacherMatch:
    teacher: TeacherRecord
    seat_block: str
    occurrences: tuple[MatchOccurrence, ...]

    @property
    def page_indexes(self) -> tuple[int, ...]:
        return tuple(sorted({item.page_index for item in self.occurrences}))


@dataclass(frozen=True)
class ProcessResult:
    output_path: Path
    matched_teachers: tuple[TeacherMatch, ...]
    unmatched_teacher_codes: tuple[str, ...] = ()
    unresolved_teacher_codes: tuple[str, ...] = ()
    warnings: tuple[str, ...] = ()


@dataclass(frozen=True)
class AppConfig:
    teachers: tuple[TeacherRecord, ...]
    seat_blocks: tuple[SeatBlockRecord, ...]
    group_order: tuple[str, ...] = ("L", "R", "C")
    text_only_mode: bool = False
    ocr_zoom: float = 2.0
    confidence_threshold: float = 0.48
    label_anchor: str = "bottom_left"
    label_x: float = 18.0
    label_y: float = 18.0
    label_margin: float = 18.0
    label_height: float = 28.0
    label_font_size: float = 14.0
    label_gap: float = 20.0
    star_scale: float = 4.95
    star_gap: float = 3.0
    star_y_offset: float = 0.0
    label_columns: tuple[float, float, float] = (24.0, 120.0, 320.0)
    highlight_padding: float = 3.0
    highlight_fill_rgb: tuple[float, float, float] = (1.0, 0.95, 0.3)
    highlight_border_rgb: tuple[float, float, float] = (0.88, 0.2, 0.15)
    label_font_path: str = ""
    source_path: Path | None = field(default=None, compare=False)
