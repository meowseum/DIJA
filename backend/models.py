from dataclasses import dataclass
from datetime import date
from typing import Optional


def _parse_int(value: str, default: int = 0) -> int:
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


def _parse_bool(value: str, default: bool = False) -> bool:
    if value is None:
        return default
    if isinstance(value, bool):
        return value
    text = str(value).strip().lower()
    if text in {"1", "true", "yes", "y"}:
        return True
    if text in {"0", "false", "no", "n"}:
        return False
    return default


@dataclass
class ClassRecord:
    id: str
    sku: str
    level: str
    location: str
    start_month: int
    class_letter: str
    start_year: int
    classroom: str
    start_date: str
    weekday: int
    start_time: str
    teacher: str
    relay_teacher: str
    relay_date: str
    student_count: int
    lesson_total: int
    status: str
    doorplate_done: bool = False
    questionnaire_done: bool = False
    intro_done: bool = False
    merged_into_id: str = ""
    promoted_to_id: str = ""
    notes: str = ""

    @classmethod
    def from_dict(cls, row: dict) -> "ClassRecord":
        return cls(
            id=row.get("id", ""),
            sku=row.get("sku", ""),
            level=row.get("level", ""),
            location=row.get("location", ""),
            start_month=_parse_int(row.get("start_month", ""), 0),
            class_letter=row.get("class_letter", ""),
            start_year=_parse_int(row.get("start_year", ""), 0),
            classroom=row.get("classroom", ""),
            start_date=row.get("start_date", ""),
            weekday=_parse_int(row.get("weekday", ""), 0),
            start_time=row.get("start_time", ""),
            teacher=row.get("teacher", ""),
            relay_teacher=row.get("relay_teacher", ""),
            relay_date=row.get("relay_date", ""),
            student_count=_parse_int(row.get("student_count", ""), 0),
            lesson_total=_parse_int(row.get("lesson_total", ""), 0),
            status=row.get("status", "active"),
            doorplate_done=_parse_bool(row.get("doorplate_done", ""), False),
            questionnaire_done=_parse_bool(row.get("questionnaire_done", ""), False),
            intro_done=_parse_bool(row.get("intro_done", ""), False),
            merged_into_id=row.get("merged_into_id", ""),
            promoted_to_id=row.get("promoted_to_id", ""),
            notes=row.get("notes", ""),
        )

    def to_dict(self) -> dict:
        return {
            "id": self.id,
            "sku": self.sku,
            "level": self.level,
            "location": self.location,
            "start_month": str(self.start_month),
            "class_letter": self.class_letter,
            "start_year": str(self.start_year),
            "classroom": self.classroom,
            "start_date": self.start_date,
            "weekday": str(self.weekday),
            "start_time": self.start_time,
            "teacher": self.teacher,
            "relay_teacher": self.relay_teacher,
            "relay_date": self.relay_date,
            "student_count": str(self.student_count),
            "lesson_total": str(self.lesson_total),
            "status": self.status,
            "doorplate_done": "1" if self.doorplate_done else "0",
            "questionnaire_done": "1" if self.questionnaire_done else "0",
            "intro_done": "1" if self.intro_done else "0",
            "merged_into_id": self.merged_into_id,
            "promoted_to_id": self.promoted_to_id,
            "notes": self.notes,
        }


@dataclass
class HolidayRange:
    id: str
    start_date: str
    end_date: str
    name: str = ""

    @classmethod
    def from_dict(cls, row: dict) -> "HolidayRange":
        return cls(
            id=row.get("id", ""),
            start_date=row.get("start_date", ""),
            end_date=row.get("end_date", ""),
            name=row.get("name", ""),
        )

    def to_dict(self) -> dict:
        return {
            "id": self.id,
            "start_date": self.start_date,
            "end_date": self.end_date,
            "name": self.name,
        }


@dataclass
class PostponeRecord:
    id: str
    class_id: str
    original_date: str
    reason: str
    make_up_date: str

    @classmethod
    def from_dict(cls, row: dict) -> "PostponeRecord":
        return cls(
            id=row.get("id", ""),
            class_id=row.get("class_id", ""),
            original_date=row.get("original_date", ""),
            reason=row.get("reason", ""),
            make_up_date=row.get("make_up_date", ""),
        )

    def to_dict(self) -> dict:
        return {
            "id": self.id,
            "class_id": self.class_id,
            "original_date": self.original_date,
            "reason": self.reason,
            "make_up_date": self.make_up_date,
        }


@dataclass
class LessonOverride:
    id: str
    class_id: str
    date: str
    action: str

    @classmethod
    def from_dict(cls, row: dict) -> "LessonOverride":
        return cls(
            id=row.get("id", ""),
            class_id=row.get("class_id", ""),
            date=row.get("date", ""),
            action=row.get("action", ""),
        )

    def to_dict(self) -> dict:
        return {
            "id": self.id,
            "class_id": self.class_id,
            "date": self.date,
            "action": self.action,
        }


def parse_date(value: str) -> Optional[date]:
    if not value:
        return None
    try:
        parts = [int(part) for part in value.split("-")]
        if len(parts) != 3:
            return None
        return date(parts[0], parts[1], parts[2])
    except ValueError:
        return None
