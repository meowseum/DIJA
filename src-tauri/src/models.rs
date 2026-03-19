use chrono::NaiveDate;
use serde::{Deserialize, Deserializer, Serialize, Serializer};

// ---------------------------------------------------------------------------
// Helper parsers (mirror Python's _parse_int, _parse_bool, parse_date)
// ---------------------------------------------------------------------------

pub fn parse_int(value: &str, default: i64) -> i64 {
    value.trim().parse::<f64>().map(|v| v as i64).unwrap_or(default)
}

pub fn parse_bool_loose(value: &str, default: bool) -> bool {
    let text = value.trim().to_lowercase();
    match text.as_str() {
        "1" | "true" | "yes" | "y" => true,
        "0" | "false" | "no" | "n" => false,
        _ => default,
    }
}

pub fn parse_date(value: &str) -> Option<NaiveDate> {
    if value.is_empty() {
        return None;
    }
    NaiveDate::parse_from_str(value, "%Y-%m-%d").ok()
}

// ---------------------------------------------------------------------------
// Custom serde helpers for CSV bool fields ("1"/"0")
// ---------------------------------------------------------------------------

fn serialize_bool_csv<S: Serializer>(val: &bool, s: S) -> Result<S::Ok, S::Error> {
    s.serialize_str(if *val { "1" } else { "0" })
}

fn deserialize_bool_csv<'de, D: Deserializer<'de>>(d: D) -> Result<bool, D::Error> {
    let raw = String::deserialize(d)?;
    Ok(parse_bool_loose(&raw, false))
}

fn serialize_i64_csv<S: Serializer>(val: &i64, s: S) -> Result<S::Ok, S::Error> {
    s.serialize_str(&val.to_string())
}

fn deserialize_i64_csv<'de, D: Deserializer<'de>>(d: D) -> Result<i64, D::Error> {
    let raw = String::deserialize(d)?;
    Ok(parse_int(&raw, 0))
}

// ---------------------------------------------------------------------------
// ClassRecord  (22 fields)
// ---------------------------------------------------------------------------

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ClassRecord {
    #[serde(default)]
    pub id: String,
    #[serde(default)]
    pub sku: String,
    #[serde(default)]
    pub level: String,
    #[serde(default)]
    pub location: String,
    #[serde(default, serialize_with = "serialize_i64_csv", deserialize_with = "deserialize_i64_csv")]
    pub start_month: i64,
    #[serde(default)]
    pub class_letter: String,
    #[serde(default, serialize_with = "serialize_i64_csv", deserialize_with = "deserialize_i64_csv")]
    pub start_year: i64,
    #[serde(default)]
    pub classroom: String,
    #[serde(default)]
    pub start_date: String,
    #[serde(default, serialize_with = "serialize_i64_csv", deserialize_with = "deserialize_i64_csv")]
    pub weekday: i64,
    #[serde(default)]
    pub start_time: String,
    #[serde(default)]
    pub teacher: String,
    #[serde(default)]
    pub relay_teacher: String,
    #[serde(default)]
    pub relay_date: String,
    #[serde(default, serialize_with = "serialize_i64_csv", deserialize_with = "deserialize_i64_csv")]
    pub student_count: i64,
    #[serde(default, serialize_with = "serialize_i64_csv", deserialize_with = "deserialize_i64_csv")]
    pub lesson_total: i64,
    #[serde(default = "default_status")]
    pub status: String,
    #[serde(default, serialize_with = "serialize_bool_csv", deserialize_with = "deserialize_bool_csv")]
    pub doorplate_done: bool,
    #[serde(default, serialize_with = "serialize_bool_csv", deserialize_with = "deserialize_bool_csv")]
    pub questionnaire_done: bool,
    #[serde(default, serialize_with = "serialize_bool_csv", deserialize_with = "deserialize_bool_csv")]
    pub intro_done: bool,
    #[serde(default)]
    pub merged_into_id: String,
    #[serde(default)]
    pub promoted_to_id: String,
    #[serde(default)]
    pub notes: String,
}

fn default_status() -> String {
    "active".to_string()
}

// ---------------------------------------------------------------------------
// HolidayRange
// ---------------------------------------------------------------------------

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct HolidayRange {
    #[serde(default)]
    pub id: String,
    #[serde(default)]
    pub start_date: String,
    #[serde(default)]
    pub end_date: String,
    #[serde(default)]
    pub name: String,
}

// ---------------------------------------------------------------------------
// PostponeRecord
// ---------------------------------------------------------------------------

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PostponeRecord {
    #[serde(default)]
    pub id: String,
    #[serde(default)]
    pub class_id: String,
    #[serde(default)]
    pub original_date: String,
    #[serde(default)]
    pub reason: String,
    #[serde(default)]
    pub make_up_date: String,
}

// ---------------------------------------------------------------------------
// LessonOverride
// ---------------------------------------------------------------------------

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct LessonOverride {
    #[serde(default)]
    pub id: String,
    #[serde(default)]
    pub class_id: String,
    #[serde(default)]
    pub date: String,
    #[serde(default)]
    pub action: String,
}
