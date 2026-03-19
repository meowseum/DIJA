use regex::Regex;
use serde_json::{json, Value};
use std::sync::LazyLock;

static FULL_PATTERN: LazyLock<Regex> = LazyLock::new(|| {
    Regex::new(r"^(?P<level>.*?)(?P<location>[KLH])?(?P<month>\d{1,2})(?P<class_letter>[A-Z])(?P<year>\d{2})$").unwrap()
});

static CODE_PATTERN: LazyLock<Regex> = LazyLock::new(|| {
    Regex::new(r"^(?P<location>[KLH])?(?P<month>\d{1,2})(?P<class_letter>[A-Z])(?P<year>\d{2})$").unwrap()
});

pub fn parse_sku(sku: &str) -> Option<Value> {
    let value = sku.trim();
    let caps = FULL_PATTERN.captures(value).or_else(|| CODE_PATTERN.captures(value))?;

    let month_str = caps.name("month").map(|m| m.as_str()).unwrap_or("");
    let month: i64 = month_str.parse().ok()?;
    if !(1..=12).contains(&month) {
        return None;
    }

    let year_short = caps.name("year").map(|m| m.as_str()).unwrap_or("");
    let year_full = 2000 + year_short.parse::<i64>().ok()?;
    let location = caps.name("location").map(|m| m.as_str()).unwrap_or("");
    let level = caps.name("level").map(|m| m.as_str()).unwrap_or("");
    let class_letter = caps.name("class_letter").map(|m| m.as_str()).unwrap_or("");
    let code = format!("{}{}{}{}", location, month_str, class_letter, year_short);

    Some(json!({
        "level": level,
        "location": location,
        "start_month": month,
        "class_letter": class_letter,
        "start_year": year_full,
        "code": code,
    }))
}

pub fn build_sku(level: &str, location: &str, start_month: i64, class_letter: &str, start_year: i64) -> String {
    let year_short = format!("{:02}", start_year % 100);
    format!("{}{}{}{}{}", level, location, start_month, class_letter, year_short)
}
