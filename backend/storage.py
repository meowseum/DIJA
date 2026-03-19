import csv
import shutil
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Iterable, List, Type

from .config import data_file, get_data_dir
from .models import ClassRecord, HolidayRange, PostponeRecord, LessonOverride


CLASS_HEADERS = [
    "id",
    "sku",
    "level",
    "location",
    "start_month",
    "class_letter",
    "start_year",
    "classroom",
    "start_date",
    "weekday",
    "start_time",
    "teacher",
    "relay_teacher",
    "relay_date",
    "student_count",
    "lesson_total",
    "status",
    "doorplate_done",
    "questionnaire_done",
    "intro_done",
    "merged_into_id",
    "promoted_to_id",
    "notes",
]

HOLIDAY_HEADERS = ["id", "start_date", "end_date", "name"]
POSTPONE_HEADERS = ["id", "class_id", "original_date", "reason", "make_up_date"]
OVERRIDE_HEADERS = ["id", "class_id", "date", "action"]
SETTINGS_HEADERS = ["type", "value"]
APP_CONFIG_HEADERS = ["key", "value"]
STOCK_HISTORY_HEADERS = ["month", "textbook_name", "count", "timestamp"]

EPS_RECORD_HEADERS = [
    "date", "item_name", "item_price", "item_section",
    "qty_K", "qty_L", "qty_HK", "subtotal", "period",
]
EPS_AUDIT_HEADERS = [
    "date",
    "operator_1_before", "operator_2_before", "operator_3_before",
    "operator_1_after", "operator_2_after", "operator_3_after",
    "operators_sum_before", "operators_sum_after",
    "sheet_before", "sheet_after",
    "past_day_carry", "calculated_total", "status",
    "status_after", "status_audit",
]


def _ensure_file(path: Path, headers: List[str]) -> None:
    if path.exists():
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as file:
        writer = csv.DictWriter(file, fieldnames=headers)
        writer.writeheader()


def _load_records(path: Path, headers: List[str], model_cls: Type) -> list:
    _ensure_file(path, headers)
    with path.open("r", newline="", encoding="utf-8") as file:
        reader = csv.DictReader(file)
        return [model_cls.from_dict(row) for row in reader]


def _save_records(path: Path, headers: List[str], records: Iterable) -> None:
    _ensure_file(path, headers)
    tmp_fd, tmp_path_str = tempfile.mkstemp(dir=path.parent, suffix=".tmp")
    tmp_path = Path(tmp_path_str)
    try:
        with open(tmp_fd, "w", newline="", encoding="utf-8") as file:
            writer = csv.DictWriter(file, fieldnames=headers)
            writer.writeheader()
            for record in records:
                writer.writerow(record.to_dict())
        shutil.move(str(tmp_path), str(path))
    except Exception:
        tmp_path.unlink(missing_ok=True)
        raise


def backup_file(path: Path) -> None:
    """Copy path to data/backups/ with a timestamp suffix."""
    if not path.exists():
        return
    backup_dir = get_data_dir() / "backups"
    backup_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    dest = backup_dir / f"{path.stem}_{stamp}{path.suffix}"
    shutil.copy2(str(path), str(dest))


def load_classes() -> List[ClassRecord]:
    return _load_records(data_file("classes.csv"), CLASS_HEADERS, ClassRecord)


def save_classes(records: Iterable[ClassRecord]) -> None:
    _save_records(data_file("classes.csv"), CLASS_HEADERS, records)


def load_holidays() -> List[HolidayRange]:
    return _load_records(data_file("holidays.csv"), HOLIDAY_HEADERS, HolidayRange)


def save_holidays(records: Iterable[HolidayRange]) -> None:
    _save_records(data_file("holidays.csv"), HOLIDAY_HEADERS, records)


def load_postpones() -> List[PostponeRecord]:
    return _load_records(data_file("postpones.csv"), POSTPONE_HEADERS, PostponeRecord)


def save_postpones(records: Iterable[PostponeRecord]) -> None:
    _save_records(data_file("postpones.csv"), POSTPONE_HEADERS, records)


def load_overrides() -> List[LessonOverride]:
    return _load_records(data_file("overrides.csv"), OVERRIDE_HEADERS, LessonOverride)


def save_overrides(records: Iterable[LessonOverride]) -> None:
    _save_records(data_file("overrides.csv"), OVERRIDE_HEADERS, records)


def _parse_level_price(value: str):
    if not value:
        return None
    for sep in ("|", "=", ":"):
        if sep in value:
            level, price = value.split(sep, 1)
            level = level.strip()
            price = price.strip()
            if not level or not price:
                return None
            try:
                return level, max(0, int(float(price)))
            except ValueError:
                return None
    return None


def _parse_message_category(value: str):
    if not value:
        return None
    for sep in ("|", "=", ":"):
        if sep in value:
            name, category = value.split(sep, 1)
            name = name.strip()
            category = category.strip()
            if not name:
                return None
            return name, category
    return None


def load_settings() -> dict:
    path = data_file("settings.csv")
    _ensure_file(path, SETTINGS_HEADERS)
    settings = {
        "teacher": [],
        "room": [],
        "level": [],
        "time": [],
        "level_price": {},
        "message_category": {},
        "textbook": {},          # name -> price
        "level_textbook": {},    # level -> textbook_name
        "textbook_stock": {},    # textbook_name -> count
        "level_next": {},        # level -> next_level
    }
    with path.open("r", newline="", encoding="utf-8") as file:
        reader = csv.DictReader(file)
        for row in reader:
            entry_type = (row.get("type") or "").strip()
            value = (row.get("value") or "").strip()
            if entry_type == "level_price":
                parsed = _parse_level_price(value)
                if parsed:
                    level, price = parsed
                    settings["level_price"][level] = price
                continue
            if entry_type == "message_category":
                parsed = _parse_message_category(value)
                if parsed:
                    name, category = parsed
                    settings["message_category"][name] = category
                continue
            if entry_type == "textbook":
                parts = value.split("|", 1)
                if len(parts) == 2:
                    name, price_str = parts[0].strip(), parts[1].strip()
                    if name:
                        try:
                            settings["textbook"][name] = max(0, int(float(price_str)))
                        except ValueError:
                            settings["textbook"][name] = 0
                continue
            if entry_type == "level_textbook":
                parts = value.split("|", 1)
                if len(parts) == 2:
                    k = parts[0].strip()
                    tb_list = [n.strip() for n in parts[1].split(",") if n.strip()]
                    if k and tb_list:
                        settings["level_textbook"][k] = tb_list
                continue
            if entry_type == "level_next":
                parts = value.split("|", 1)
                if len(parts) == 2:
                    k, v = parts[0].strip(), parts[1].strip()
                    if k:
                        settings["level_next"][k] = v
                continue
            if entry_type == "textbook_stock":
                parts = value.split("|", 1)
                if len(parts) == 2:
                    name, count_str = parts[0].strip(), parts[1].strip()
                    if name:
                        try:
                            settings["textbook_stock"][name] = max(0, int(float(count_str)))
                        except ValueError:
                            settings["textbook_stock"][name] = 0
                continue
            if entry_type in settings and value and value not in settings[entry_type]:
                settings[entry_type].append(value)
    return settings


def save_settings(settings: dict) -> None:
    path = data_file("settings.csv")
    _ensure_file(path, SETTINGS_HEADERS)
    rows = []
    for entry_type in ("teacher", "room", "level", "time"):
        for value in settings.get(entry_type, []):
            rows.append({"type": entry_type, "value": value})
    level_prices = settings.get("level_price", {}) or {}
    for level, price in level_prices.items():
        rows.append({"type": "level_price", "value": f"{level}|{price}"})
    message_categories = settings.get("message_category", {}) or {}
    for name, category in message_categories.items():
        rows.append({"type": "message_category", "value": f"{name}|{category}"})
    textbooks = settings.get("textbook", {}) or {}
    for name, price in textbooks.items():
        rows.append({"type": "textbook", "value": f"{name}|{price}"})
    for k, v in (settings.get("level_textbook", {}) or {}).items():
        if isinstance(v, list) and v:
            rows.append({"type": "level_textbook", "value": f"{k}|{','.join(v)}"})
    for k, v in (settings.get("level_next", {}) or {}).items():
        rows.append({"type": "level_next", "value": f"{k}|{v}"})
    textbook_stock = settings.get("textbook_stock", {}) or {}
    for name, count in textbook_stock.items():
        rows.append({"type": "textbook_stock", "value": f"{name}|{count}"})
    with path.open("w", newline="", encoding="utf-8") as file:
        writer = csv.DictWriter(file, fieldnames=SETTINGS_HEADERS)
        writer.writeheader()
        writer.writerows(rows)


def load_app_config() -> dict:
    path = data_file("app_config.csv")
    _ensure_file(path, APP_CONFIG_HEADERS)
    config = {"location": ""}
    with path.open("r", newline="", encoding="utf-8") as file:
        reader = csv.DictReader(file)
        for row in reader:
            key = (row.get("key") or "").strip()
            value = (row.get("value") or "").strip()
            if key:
                config[key] = value
    return config


def save_app_config(config: dict) -> None:
    path = data_file("app_config.csv")
    _ensure_file(path, APP_CONFIG_HEADERS)
    rows = []
    for key, value in config.items():
        rows.append({"key": key, "value": value})
    with path.open("w", newline="", encoding="utf-8") as file:
        writer = csv.DictWriter(file, fieldnames=APP_CONFIG_HEADERS)
        writer.writeheader()
        writer.writerows(rows)


def load_stock_history() -> dict:
    """Return dict: { 'YYYY-MM': { textbook_name: count, ... }, ... }"""
    path = data_file("stock_history.csv")
    _ensure_file(path, STOCK_HISTORY_HEADERS)
    history: dict = {}
    with path.open("r", newline="", encoding="utf-8") as file:
        reader = csv.DictReader(file)
        for row in reader:
            month = (row.get("month") or "").strip()
            name = (row.get("textbook_name") or "").strip()
            count_str = (row.get("count") or "0").strip()
            if not month or not name:
                continue
            try:
                count = max(0, int(float(count_str)))
            except ValueError:
                count = 0
            if month not in history:
                history[month] = {}
            history[month][name] = count
    return history


def save_stock_snapshot(month: str, stock_data: dict) -> None:
    """Save/overwrite all stock entries for a given month. stock_data: { name: count }"""
    path = data_file("stock_history.csv")
    _ensure_file(path, STOCK_HISTORY_HEADERS)
    existing_rows = []
    with path.open("r", newline="", encoding="utf-8") as file:
        reader = csv.DictReader(file)
        for row in reader:
            if (row.get("month") or "").strip() != month:
                existing_rows.append(row)
    timestamp = datetime.now().isoformat()
    for name, count in stock_data.items():
        existing_rows.append({
            "month": month,
            "textbook_name": name,
            "count": str(count),
            "timestamp": timestamp,
        })
    _atomic_write_dicts(path, STOCK_HISTORY_HEADERS, existing_rows)


# ---------------------------------------------------------------------------
# EPS records / audit
# ---------------------------------------------------------------------------

def _atomic_write_dicts(path: Path, headers: List[str], rows: list) -> None:
    """Atomically write a list of dicts to a CSV file."""
    _ensure_file(path, headers)
    tmp_fd, tmp_path_str = tempfile.mkstemp(dir=path.parent, suffix=".tmp")
    tmp_path = Path(tmp_path_str)
    try:
        with open(tmp_fd, "w", newline="", encoding="utf-8") as file:
            writer = csv.DictWriter(file, fieldnames=headers)
            writer.writeheader()
            writer.writerows(rows)
        shutil.move(str(tmp_path), str(path))
    except Exception:
        tmp_path.unlink(missing_ok=True)
        raise


def load_eps_records(date_str: str) -> List[dict]:
    """Return all EPS record rows for a given date."""
    path = data_file("eps_records.csv")
    _ensure_file(path, EPS_RECORD_HEADERS)
    results = []
    with path.open("r", newline="", encoding="utf-8") as file:
        for row in csv.DictReader(file):
            if (row.get("date") or "").strip() == date_str:
                results.append(row)
    return results


def save_eps_records(date_str: str, new_rows: List[dict]) -> None:
    """Overwrite EPS record rows for *date_str*, keep other dates intact."""
    path = data_file("eps_records.csv")
    _ensure_file(path, EPS_RECORD_HEADERS)
    kept = []
    with path.open("r", newline="", encoding="utf-8") as file:
        for row in csv.DictReader(file):
            if (row.get("date") or "").strip() != date_str:
                kept.append(row)
    kept.extend(new_rows)
    _atomic_write_dicts(path, EPS_RECORD_HEADERS, kept)


def load_eps_audit(date_str: str) -> dict | None:
    """Return the audit row for a given date, or None."""
    path = data_file("eps_audit.csv")
    _ensure_file(path, EPS_AUDIT_HEADERS)
    with path.open("r", newline="", encoding="utf-8") as file:
        for row in csv.DictReader(file):
            if (row.get("date") or "").strip() == date_str:
                return row
    return None


def save_eps_audit(date_str: str, audit_row: dict) -> None:
    """Overwrite the audit row for *date_str*, keep other dates intact."""
    path = data_file("eps_audit.csv")
    _ensure_file(path, EPS_AUDIT_HEADERS)
    kept = []
    with path.open("r", newline="", encoding="utf-8") as file:
        for row in csv.DictReader(file):
            if (row.get("date") or "").strip() != date_str:
                kept.append(row)
    kept.append(audit_row)
    _atomic_write_dicts(path, EPS_AUDIT_HEADERS, kept)


def list_eps_dates() -> List[str]:
    """Return sorted distinct dates that have EPS records."""
    path = data_file("eps_records.csv")
    _ensure_file(path, EPS_RECORD_HEADERS)
    dates = set()
    with path.open("r", newline="", encoding="utf-8") as file:
        for row in csv.DictReader(file):
            d = (row.get("date") or "").strip()
            if d:
                dates.add(d)
    return sorted(dates)


def get_eps_after_total(date_str: str) -> int:
    """Sum of subtotal for rows where date=date_str and period='after'."""
    total = 0
    for row in load_eps_records(date_str):
        if (row.get("period") or "").strip() == "after":
            try:
                total += int(float(row.get("subtotal") or "0"))
            except ValueError:
                pass
    return total


def get_eps_after_items(date_str: str) -> dict:
    """Return {item_name: {qty_K, qty_L, qty_HK}} for after-1900 records."""
    result = {}
    for row in load_eps_records(date_str):
        if (row.get("period") or "").strip() == "after":
            name = (row.get("item_name") or "").strip()
            if name:
                result[name] = {
                    "qty_K": int(float(row.get("qty_K") or "0")),
                    "qty_L": int(float(row.get("qty_L") or "0")),
                    "qty_HK": int(float(row.get("qty_HK") or "0")),
                }
    return result


def has_eps_records_for_date(date_str: str, period: str = "before") -> bool:
    """Check if any saved records exist for a given date and period."""
    for row in load_eps_records(date_str):
        if (row.get("period") or "").strip() == period:
            return True
    return False
