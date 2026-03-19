from datetime import date, timedelta
from typing import Iterable, List, Optional

from .models import HolidayRange, PostponeRecord, LessonOverride, parse_date


def holiday_set(holiday_ranges: Iterable[HolidayRange]) -> set[date]:
    holidays = set()
    for holiday in holiday_ranges:
        start = parse_date(holiday.start_date)
        end = parse_date(holiday.end_date)
        if not start or not end:
            continue
        if end < start:
            start, end = end, start
        current = start
        while current <= end:
            holidays.add(current)
            current += timedelta(days=1)
    return holidays


def _next_weekday(after_date: date, weekday: int) -> date:
    days_ahead = (weekday - after_date.weekday()) % 7
    if days_ahead == 0:
        days_ahead = 7
    return after_date + timedelta(days=days_ahead)


def _find_next_available_weekly(
    after_date: date, weekday: int, scheduled: set[date], holidays: set[date]
) -> date:
    candidate = _next_weekday(after_date, weekday)
    while candidate in scheduled or candidate in holidays:
        candidate = _next_weekday(candidate, weekday)
    return candidate


def generate_weekly_schedule(
    start_date: date, weekday: int, lesson_total: int, holiday_ranges: Iterable[HolidayRange]
) -> List[date]:
    holidays = holiday_set(holiday_ranges)
    schedule = []
    current = start_date
    if current.weekday() != weekday:
        current = _next_weekday(current - timedelta(days=1), weekday)

    while len(schedule) < lesson_total:
        if current not in holidays:
            schedule.append(current)
        current = current + timedelta(days=7)

    return schedule


def apply_postpones(
    schedule: List[date],
    weekday: int,
    holiday_ranges: Iterable[HolidayRange],
    postpones: Iterable[PostponeRecord],
) -> List[date]:
    holidays = holiday_set(holiday_ranges)
    updated = list(schedule)
    scheduled = set(updated)

    for postpone in postpones:
        original = parse_date(postpone.original_date)
        if not original or original not in scheduled:
            continue
        scheduled.remove(original)
        updated.remove(original)

        target = parse_date(postpone.make_up_date)
        if not target or target in holidays or target in scheduled or target <= original:
            target = _find_next_available_weekly(original, weekday, scheduled, holidays)
        scheduled.add(target)
        updated.append(target)

    updated.sort()
    return updated


def apply_overrides(schedule: List[date], holiday_ranges: Iterable[HolidayRange], overrides: Iterable[LessonOverride]) -> List[date]:
    holidays = holiday_set(holiday_ranges)
    updated = list(schedule)
    for override in overrides:
        target = parse_date(override.date)
        if not target:
            continue
        if override.action == "remove":
            if target in updated:
                updated.remove(target)
        elif override.action == "add":
            if target not in updated and target not in holidays:
                updated.append(target)
    updated.sort()
    return updated


def calculate_progress(
    start_date_str: str,
    weekday: int,
    lesson_total: int,
    holiday_ranges: Iterable[HolidayRange],
    postpones: Iterable[PostponeRecord],
    overrides: Iterable[LessonOverride],
) -> dict:
    start_date = parse_date(start_date_str)
    if not start_date:
        return {
            "lessons_elapsed": 0,
            "lessons_remaining": lesson_total,
            "end_date": "",
            "next_lesson_date": "",
        }

    schedule = generate_weekly_schedule(start_date, weekday, lesson_total, holiday_ranges)
    schedule = apply_postpones(schedule, weekday, holiday_ranges, postpones)
    schedule = apply_overrides(schedule, holiday_ranges, overrides)

    today = date.today()
    elapsed = sum(1 for lesson in schedule if lesson <= today)
    remaining = max(lesson_total - elapsed, 0)
    end_date = schedule[-1].isoformat() if schedule else ""
    next_lesson = next((lesson for lesson in schedule if lesson >= today), None)
    next_lesson_date = next_lesson.isoformat() if next_lesson else ""

    return {
        "lessons_elapsed": elapsed,
        "lessons_remaining": remaining,
        "end_date": end_date,
        "next_lesson_date": next_lesson_date,
    }
