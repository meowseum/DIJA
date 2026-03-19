use chrono::{Datelike, Duration, NaiveDate};
use std::collections::HashSet;

use crate::models::{parse_date, HolidayRange, LessonOverride, PostponeRecord};

pub fn holiday_set(holiday_ranges: &[HolidayRange]) -> HashSet<NaiveDate> {
    let mut holidays = HashSet::new();
    for holiday in holiday_ranges {
        let start = match parse_date(&holiday.start_date) {
            Some(d) => d,
            None => continue,
        };
        let end = match parse_date(&holiday.end_date) {
            Some(d) => d,
            None => continue,
        };
        let (start, end) = if end < start { (end, start) } else { (start, end) };
        let mut current = start;
        while current <= end {
            holidays.insert(current);
            current += Duration::days(1);
        }
    }
    holidays
}

fn next_weekday(after_date: NaiveDate, weekday: u32) -> NaiveDate {
    let days_ahead = ((weekday as i64 - after_date.weekday().num_days_from_monday() as i64) % 7 + 7) % 7;
    let days_ahead = if days_ahead == 0 { 7 } else { days_ahead };
    after_date + Duration::days(days_ahead)
}

pub fn find_next_available_weekly(
    after_date: NaiveDate,
    weekday: u32,
    scheduled: &HashSet<NaiveDate>,
    holidays: &HashSet<NaiveDate>,
) -> NaiveDate {
    let mut candidate = next_weekday(after_date, weekday);
    while scheduled.contains(&candidate) || holidays.contains(&candidate) {
        candidate = next_weekday(candidate, weekday);
    }
    candidate
}

pub fn generate_weekly_schedule(
    start_date: NaiveDate,
    weekday: u32,
    lesson_total: i64,
    holiday_ranges: &[HolidayRange],
) -> Vec<NaiveDate> {
    let holidays = holiday_set(holiday_ranges);
    let mut schedule = Vec::new();
    let mut current = start_date;
    if current.weekday().num_days_from_monday() != weekday {
        current = next_weekday(current - Duration::days(1), weekday);
    }

    while (schedule.len() as i64) < lesson_total {
        if !holidays.contains(&current) {
            schedule.push(current);
        }
        current += Duration::days(7);
    }

    schedule
}

pub fn apply_postpones(
    schedule: &[NaiveDate],
    weekday: u32,
    holiday_ranges: &[HolidayRange],
    postpones: &[PostponeRecord],
) -> Vec<NaiveDate> {
    let holidays = holiday_set(holiday_ranges);
    let mut updated: Vec<NaiveDate> = schedule.to_vec();
    let mut scheduled: HashSet<NaiveDate> = updated.iter().cloned().collect();

    for postpone in postpones {
        let original = match parse_date(&postpone.original_date) {
            Some(d) => d,
            None => continue,
        };
        if !scheduled.contains(&original) {
            continue;
        }
        scheduled.remove(&original);
        updated.retain(|d| *d != original);

        let mut target = parse_date(&postpone.make_up_date);
        if let Some(t) = target {
            if holidays.contains(&t) || scheduled.contains(&t) || t <= original {
                target = None;
            }
        }
        let target = target.unwrap_or_else(|| {
            find_next_available_weekly(original, weekday, &scheduled, &holidays)
        });

        scheduled.insert(target);
        updated.push(target);
    }

    updated.sort();
    updated
}

pub fn apply_overrides(
    schedule: &[NaiveDate],
    holiday_ranges: &[HolidayRange],
    overrides: &[LessonOverride],
) -> Vec<NaiveDate> {
    let holidays = holiday_set(holiday_ranges);
    let mut updated: Vec<NaiveDate> = schedule.to_vec();
    for ov in overrides {
        let target = match parse_date(&ov.date) {
            Some(d) => d,
            None => continue,
        };
        if ov.action == "remove" {
            updated.retain(|d| *d != target);
        } else if ov.action == "add" {
            if !updated.contains(&target) && !holidays.contains(&target) {
                updated.push(target);
            }
        }
    }
    updated.sort();
    updated
}

pub fn calculate_progress(
    start_date_str: &str,
    weekday: u32,
    lesson_total: i64,
    holiday_ranges: &[HolidayRange],
    postpones: &[PostponeRecord],
    overrides: &[LessonOverride],
) -> serde_json::Value {
    let start_date = match parse_date(start_date_str) {
        Some(d) => d,
        None => {
            return serde_json::json!({
                "lessons_elapsed": 0,
                "lessons_remaining": lesson_total,
                "end_date": "",
                "next_lesson_date": "",
            });
        }
    };

    let schedule = generate_weekly_schedule(start_date, weekday, lesson_total, holiday_ranges);
    let schedule = apply_postpones(&schedule, weekday, holiday_ranges, postpones);
    let schedule = apply_overrides(&schedule, holiday_ranges, overrides);

    let today = chrono::Local::now().date_naive();
    let elapsed = schedule.iter().filter(|&&d| d <= today).count() as i64;
    let remaining = (lesson_total - elapsed).max(0);
    let end_date = schedule.last().map(|d| d.format("%Y-%m-%d").to_string()).unwrap_or_default();
    let next_lesson = schedule.iter().find(|&&d| d >= today);
    let next_lesson_date = next_lesson.map(|d| d.format("%Y-%m-%d").to_string()).unwrap_or_default();

    serde_json::json!({
        "lessons_elapsed": elapsed,
        "lessons_remaining": remaining,
        "end_date": end_date,
        "next_lesson_date": next_lesson_date,
    })
}
