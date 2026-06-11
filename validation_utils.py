from __future__ import annotations

import os
import re
from datetime import datetime
from typing import Any

TRUE_VALUES = {"1", "true", "yes", "y"}
INACTIVE_VALUES = {"0", "false", "no", "n", "inactive"}
LOCATION_DATE_FORMATS = ("%d-%m-%Y, %H:%M", "%d-%m-%Y %H:%M", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%dT%H:%M:%S")


def clean_text(value: Any) -> str:
    return "" if value is None else str(value).replace("\u00a0", " ").strip()


def clean_digits(value: Any) -> str:
    return re.sub(r"\D", "", clean_text(value))


def first_comma_value(value: Any) -> str:
    return clean_text(value).split(",", 1)[0].strip()


def prod_branch_admin_name(value: Any) -> str:
    return " ".join(clean_text(value).split()[:2])


def normalize_key(value: str) -> str:
    return clean_text(value).lower()


def normalize_words(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", " ", clean_text(value).lower()).strip()


def parse_int_text(value: Any) -> int | None:
    digits = clean_digits(value)
    return int(digits) if digits else None


def env_flag(name: str, default: str = "0") -> bool:
    return clean_text(os.getenv(name, default)).lower() in TRUE_VALUES


def id_text(value: Any) -> str:
    return clean_text(value).replace(",", "")


def is_active_value(value: Any) -> bool:
    return clean_text(value).lower() not in INACTIVE_VALUES


def parse_location_datetime(value: Any) -> datetime | None:
    text = clean_text(value)
    for date_format in LOCATION_DATE_FORMATS if text else ():
        try:
            return datetime.strptime(text, date_format)
        except ValueError:
            pass
    return None
