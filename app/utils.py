from __future__ import annotations


def normalize_seat_no(value: object) -> str:
    text = str(value or "").strip()
    digits = "".join(ch for ch in text if ch.isdigit())
    if digits:
        return digits.zfill(2)
    return text.upper()
