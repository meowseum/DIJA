import re
from typing import Optional


FULL_PATTERN = re.compile(
    r"^(?P<level>.*?)(?P<location>[KLH])?(?P<month>\d{1,2})(?P<class_letter>[A-Z])(?P<year>\d{2})$"
)
CODE_PATTERN = re.compile(
    r"^(?P<location>[KLH])?(?P<month>\d{1,2})(?P<class_letter>[A-Z])(?P<year>\d{2})$"
)


def parse_sku(sku: str) -> Optional[dict]:
    value = sku.strip()
    match = FULL_PATTERN.match(value) or CODE_PATTERN.match(value)
    if not match:
        return None
    parts = match.groupdict()
    month_str = parts["month"]
    month = int(month_str)
    if month < 1 or month > 12:
        return None
    year_short = parts["year"]
    year_full = 2000 + int(year_short)
    location = parts.get("location") or ""
    level = parts.get("level") or ""
    code = f"{location}{month_str}{parts['class_letter']}{year_short}"
    return {
        "level": level,
        "location": location,
        "start_month": month,
        "class_letter": parts["class_letter"],
        "start_year": year_full,
        "code": code,
    }


def build_sku(level: str, location: str, start_month: int, class_letter: str, start_year: int) -> str:
    year_short = str(start_year)[-2:]
    return f"{level}{location}{start_month}{class_letter}{year_short}"
