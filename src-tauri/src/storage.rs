use std::collections::HashMap;
use std::fs;
use std::io::Write;
use std::path::Path;

use crate::config::{data_file, get_data_dir};
use crate::models::{ClassRecord, HolidayRange, LessonOverride, PostponeRecord};

// ---------------------------------------------------------------------------
// CSV Headers
// ---------------------------------------------------------------------------

pub const CLASS_HEADERS: &[&str] = &[
    "id", "sku", "level", "location", "start_month", "class_letter", "start_year",
    "classroom", "start_date", "weekday", "start_time", "teacher", "relay_teacher",
    "relay_date", "student_count", "lesson_total", "status", "doorplate_done",
    "questionnaire_done", "intro_done", "merged_into_id", "promoted_to_id", "notes",
];

pub const HOLIDAY_HEADERS: &[&str] = &["id", "start_date", "end_date", "name"];
pub const POSTPONE_HEADERS: &[&str] = &["id", "class_id", "original_date", "reason", "make_up_date"];
pub const OVERRIDE_HEADERS: &[&str] = &["id", "class_id", "date", "action"];
pub const SETTINGS_HEADERS: &[&str] = &["type", "value"];
pub const APP_CONFIG_HEADERS: &[&str] = &["key", "value"];
pub const STOCK_HISTORY_HEADERS: &[&str] = &["month", "textbook_name", "count", "timestamp"];

pub const EPS_RECORD_HEADERS: &[&str] = &[
    "date", "item_name", "item_price", "item_section",
    "qty_K", "qty_L", "qty_HK", "subtotal", "period",
];
pub const EPS_AUDIT_HEADERS: &[&str] = &[
    "date",
    "operator_1_before", "operator_2_before", "operator_3_before",
    "operator_1_after", "operator_2_after", "operator_3_after",
    "operators_sum_before", "operators_sum_after",
    "sheet_before", "sheet_after",
    "past_day_carry", "calculated_total", "status",
    "status_after", "status_audit",
];

// ---------------------------------------------------------------------------
// Ensure file exists with headers
// ---------------------------------------------------------------------------

fn ensure_file(path: &Path, headers: &[&str]) {
    if path.exists() {
        return;
    }
    if let Some(parent) = path.parent() {
        fs::create_dir_all(parent).ok();
    }
    if let Ok(mut file) = fs::File::create(path) {
        let _ = writeln!(file, "{}", headers.join(","));
    }
}

// ---------------------------------------------------------------------------
// Generic CSV load/save
// ---------------------------------------------------------------------------

pub fn load_records<T: serde::de::DeserializeOwned>(path: &Path, headers: &[&str]) -> Vec<T> {
    ensure_file(path, headers);
    let mut rdr = match csv::Reader::from_path(path) {
        Ok(r) => r,
        Err(_) => return Vec::new(),
    };
    let mut records = Vec::new();
    for result in rdr.deserialize() {
        if let Ok(record) = result {
            records.push(record);
        }
    }
    records
}

pub fn save_records<T: serde::Serialize>(path: &Path, headers: &[&str], records: &[T]) {
    ensure_file(path, headers);
    let dir = path.parent().unwrap_or(Path::new("."));
    let tmp = match tempfile::NamedTempFile::new_in(dir) {
        Ok(t) => t,
        Err(_) => return,
    };
    {
        let mut wtr = csv::WriterBuilder::new().has_headers(false).from_writer(tmp.as_file());
        // Write header explicitly (has_headers(false) prevents serialize from auto-adding)
        let _ = wtr.write_record(headers);
        for record in records {
            let _ = wtr.serialize(record);
        }
        let _ = wtr.flush();
    }
    let _ = tmp.persist(path);
}

pub fn atomic_write_dicts(path: &Path, headers: &[&str], rows: &[HashMap<String, String>]) {
    ensure_file(path, headers);
    let dir = path.parent().unwrap_or(Path::new("."));
    let tmp = match tempfile::NamedTempFile::new_in(dir) {
        Ok(t) => t,
        Err(_) => return,
    };
    {
        let mut wtr = csv::WriterBuilder::new().from_writer(tmp.as_file());
        let _ = wtr.write_record(headers);
        for row in rows {
            let record: Vec<String> = headers.iter().map(|h| row.get(*h).cloned().unwrap_or_default()).collect();
            let _ = wtr.write_record(&record);
        }
        let _ = wtr.flush();
    }
    let _ = tmp.persist(path);
}

// ---------------------------------------------------------------------------
// Backup
// ---------------------------------------------------------------------------

pub fn backup_file(path: &Path) {
    if !path.exists() {
        return;
    }
    let backup_dir = get_data_dir().join("backups");
    fs::create_dir_all(&backup_dir).ok();
    let stamp = chrono::Local::now().format("%Y%m%d_%H%M%S").to_string();
    let stem = path.file_stem().unwrap_or_default().to_str().unwrap_or("");
    let ext = path.extension().unwrap_or_default().to_str().unwrap_or("");
    let dest = if ext.is_empty() {
        backup_dir.join(format!("{}_{}", stem, stamp))
    } else {
        backup_dir.join(format!("{}_{}.{}", stem, stamp, ext))
    };
    let _ = fs::copy(path, dest);
}

// ---------------------------------------------------------------------------
// Typed load/save functions
// ---------------------------------------------------------------------------

pub fn load_classes() -> Vec<ClassRecord> {
    load_records(&data_file("classes.csv"), CLASS_HEADERS)
}

pub fn save_classes(records: &[ClassRecord]) {
    save_records(&data_file("classes.csv"), CLASS_HEADERS, records);
}

pub fn load_holidays() -> Vec<HolidayRange> {
    load_records(&data_file("holidays.csv"), HOLIDAY_HEADERS)
}

pub fn save_holidays(records: &[HolidayRange]) {
    save_records(&data_file("holidays.csv"), HOLIDAY_HEADERS, records);
}

pub fn load_postpones() -> Vec<PostponeRecord> {
    load_records(&data_file("postpones.csv"), POSTPONE_HEADERS)
}

pub fn save_postpones(records: &[PostponeRecord]) {
    save_records(&data_file("postpones.csv"), POSTPONE_HEADERS, records);
}

pub fn load_overrides() -> Vec<LessonOverride> {
    load_records(&data_file("overrides.csv"), OVERRIDE_HEADERS)
}

pub fn save_overrides(records: &[LessonOverride]) {
    save_records(&data_file("overrides.csv"), OVERRIDE_HEADERS, records);
}

// ---------------------------------------------------------------------------
// Settings (polymorphic CSV)
// ---------------------------------------------------------------------------

fn parse_level_price(value: &str) -> Option<(String, i64)> {
    if value.is_empty() {
        return None;
    }
    for sep in &["|", "=", ":"] {
        if let Some(idx) = value.find(sep) {
            let level = value[..idx].trim().to_string();
            let price_str = value[idx + sep.len()..].trim();
            if level.is_empty() || price_str.is_empty() {
                return None;
            }
            if let Ok(price) = price_str.parse::<f64>() {
                return Some((level, (price as i64).max(0)));
            }
            return None;
        }
    }
    None
}

fn parse_message_category(value: &str) -> Option<(String, String)> {
    if value.is_empty() {
        return None;
    }
    for sep in &["|", "=", ":"] {
        if let Some(idx) = value.find(sep) {
            let name = value[..idx].trim().to_string();
            let category = value[idx + sep.len()..].trim().to_string();
            if name.is_empty() {
                return None;
            }
            return Some((name, category));
        }
    }
    None
}

#[derive(Debug, Clone, Default)]
pub struct Settings {
    pub teacher: Vec<String>,
    pub room: Vec<String>,
    pub level: Vec<String>,
    pub time: Vec<String>,
    pub level_price: HashMap<String, i64>,
    pub message_category: HashMap<String, String>,
    pub textbook: HashMap<String, i64>,
    pub level_textbook: HashMap<String, Vec<String>>,
    pub textbook_stock: HashMap<String, i64>,
    pub level_next: HashMap<String, String>,
    pub eps_config: HashMap<String, String>,
    pub eps_book: Vec<(String, i64)>,
    pub eps_other: Vec<(String, i64)>,
    pub eps_special: Vec<(String, i64)>,
}

impl Settings {
    pub fn to_json(&self) -> serde_json::Value {
        let eps_book_json: Vec<serde_json::Value> = self.eps_book.iter()
            .map(|(n, p)| serde_json::json!({"name": n, "price": p}))
            .collect();
        let eps_other_json: Vec<serde_json::Value> = self.eps_other.iter()
            .map(|(n, p)| serde_json::json!({"name": n, "price": p}))
            .collect();
        let eps_special_json: Vec<serde_json::Value> = self.eps_special.iter()
            .map(|(n, p)| serde_json::json!({"name": n, "price": p}))
            .collect();
        serde_json::json!({
            "teacher": self.teacher,
            "room": self.room,
            "level": self.level,
            "time": self.time,
            "level_price": self.level_price,
            "message_category": self.message_category,
            "textbook": self.textbook,
            "level_textbook": self.level_textbook,
            "textbook_stock": self.textbook_stock,
            "level_next": self.level_next,
            "eps_config": self.eps_config,
            "eps_book": eps_book_json,
            "eps_other": eps_other_json,
            "eps_special": eps_special_json,
        })
    }

    pub fn list_mut(&mut self, entry_type: &str) -> Option<&mut Vec<String>> {
        match entry_type {
            "teacher" => Some(&mut self.teacher),
            "room" => Some(&mut self.room),
            "level" => Some(&mut self.level),
            "time" => Some(&mut self.time),
            _ => None,
        }
    }
}

pub fn load_settings() -> Settings {
    let path = data_file("settings.csv");
    ensure_file(&path, SETTINGS_HEADERS);
    let mut settings = Settings::default();

    let mut rdr = match csv::Reader::from_path(&path) {
        Ok(r) => r,
        Err(_) => return settings,
    };

    for result in rdr.records() {
        let record = match result {
            Ok(r) => r,
            Err(_) => continue,
        };
        let entry_type = record.get(0).unwrap_or("").trim();
        let value = record.get(1).unwrap_or("").trim().to_string();

        match entry_type {
            "level_price" => {
                if let Some((level, price)) = parse_level_price(&value) {
                    settings.level_price.insert(level, price);
                }
            }
            "message_category" => {
                if let Some((name, category)) = parse_message_category(&value) {
                    settings.message_category.insert(name, category);
                }
            }
            "textbook" => {
                if let Some((name, price_str)) = value.split_once('|') {
                    let name = name.trim().to_string();
                    let price = price_str.trim().parse::<f64>().map(|p| (p as i64).max(0)).unwrap_or(0);
                    if !name.is_empty() {
                        settings.textbook.insert(name, price);
                    }
                }
            }
            "level_textbook" => {
                if let Some((k, rest)) = value.split_once('|') {
                    let k = k.trim().to_string();
                    let tb_list: Vec<String> = rest.split(',').map(|n| n.trim().to_string()).filter(|n| !n.is_empty()).collect();
                    if !k.is_empty() && !tb_list.is_empty() {
                        settings.level_textbook.insert(k, tb_list);
                    }
                }
            }
            "level_next" => {
                if let Some((k, v)) = value.split_once('|') {
                    let k = k.trim().to_string();
                    let v = v.trim().to_string();
                    if !k.is_empty() {
                        settings.level_next.insert(k, v);
                    }
                }
            }
            "textbook_stock" => {
                if let Some((name, count_str)) = value.split_once('|') {
                    let name = name.trim().to_string();
                    let count = count_str.trim().parse::<f64>().map(|c| (c as i64).max(0)).unwrap_or(0);
                    if !name.is_empty() {
                        settings.textbook_stock.insert(name, count);
                    }
                }
            }
            "eps_config" => {
                if let Some((k, v)) = value.split_once('|') {
                    let k = k.trim().to_string();
                    let v = v.trim().to_string();
                    if !k.is_empty() {
                        settings.eps_config.insert(k, v);
                    }
                }
            }
            "eps_book" => {
                if let Some((name, price_str)) = value.split_once('|') {
                    let name = name.trim().to_string();
                    let price = price_str.trim().parse::<f64>().map(|p| (p as i64).max(0)).unwrap_or(0);
                    if !name.is_empty() {
                        settings.eps_book.push((name, price));
                    }
                }
            }
            "eps_other" => {
                if let Some((name, price_str)) = value.split_once('|') {
                    let name = name.trim().to_string();
                    let price = price_str.trim().parse::<f64>().map(|p| (p as i64).max(0)).unwrap_or(0);
                    if !name.is_empty() {
                        settings.eps_other.push((name, price));
                    }
                }
            }
            "eps_special" => {
                if let Some((name, price_str)) = value.split_once('|') {
                    let name = name.trim().to_string();
                    let price = price_str.trim().parse::<f64>().map(|p| (p as i64).max(0)).unwrap_or(0);
                    if !name.is_empty() {
                        settings.eps_special.push((name, price));
                    }
                }
            }
            "teacher" | "room" | "level" | "time" => {
                if !value.is_empty() {
                    if let Some(list) = settings.list_mut(entry_type) {
                        if !list.contains(&value) {
                            list.push(value);
                        }
                    }
                }
            }
            _ => {}
        }
    }

    settings
}

pub fn save_settings(settings: &Settings) {
    let path = data_file("settings.csv");
    ensure_file(&path, SETTINGS_HEADERS);
    let mut rows: Vec<HashMap<String, String>> = Vec::new();

    for entry_type in &["teacher", "room", "level", "time"] {
        let list = match *entry_type {
            "teacher" => &settings.teacher,
            "room" => &settings.room,
            "level" => &settings.level,
            "time" => &settings.time,
            _ => continue,
        };
        for value in list {
            let mut row = HashMap::new();
            row.insert("type".to_string(), entry_type.to_string());
            row.insert("value".to_string(), value.clone());
            rows.push(row);
        }
    }
    for (level, price) in &settings.level_price {
        let mut row = HashMap::new();
        row.insert("type".to_string(), "level_price".to_string());
        row.insert("value".to_string(), format!("{}|{}", level, price));
        rows.push(row);
    }
    for (name, category) in &settings.message_category {
        let mut row = HashMap::new();
        row.insert("type".to_string(), "message_category".to_string());
        row.insert("value".to_string(), format!("{}|{}", name, category));
        rows.push(row);
    }
    for (name, price) in &settings.textbook {
        let mut row = HashMap::new();
        row.insert("type".to_string(), "textbook".to_string());
        row.insert("value".to_string(), format!("{}|{}", name, price));
        rows.push(row);
    }
    for (k, v) in &settings.level_textbook {
        if !v.is_empty() {
            let mut row = HashMap::new();
            row.insert("type".to_string(), "level_textbook".to_string());
            row.insert("value".to_string(), format!("{}|{}", k, v.join(",")));
            rows.push(row);
        }
    }
    for (k, v) in &settings.level_next {
        let mut row = HashMap::new();
        row.insert("type".to_string(), "level_next".to_string());
        row.insert("value".to_string(), format!("{}|{}", k, v));
        rows.push(row);
    }
    for (name, count) in &settings.textbook_stock {
        let mut row = HashMap::new();
        row.insert("type".to_string(), "textbook_stock".to_string());
        row.insert("value".to_string(), format!("{}|{}", name, count));
        rows.push(row);
    }
    for (k, v) in &settings.eps_config {
        let mut row = HashMap::new();
        row.insert("type".to_string(), "eps_config".to_string());
        row.insert("value".to_string(), format!("{}|{}", k, v));
        rows.push(row);
    }
    for (name, price) in &settings.eps_book {
        let mut row = HashMap::new();
        row.insert("type".to_string(), "eps_book".to_string());
        row.insert("value".to_string(), format!("{}|{}", name, price));
        rows.push(row);
    }
    for (name, price) in &settings.eps_other {
        let mut row = HashMap::new();
        row.insert("type".to_string(), "eps_other".to_string());
        row.insert("value".to_string(), format!("{}|{}", name, price));
        rows.push(row);
    }
    for (name, price) in &settings.eps_special {
        let mut row = HashMap::new();
        row.insert("type".to_string(), "eps_special".to_string());
        row.insert("value".to_string(), format!("{}|{}", name, price));
        rows.push(row);
    }

    atomic_write_dicts(&path, SETTINGS_HEADERS, &rows);
}

// ---------------------------------------------------------------------------
// App config
// ---------------------------------------------------------------------------

pub fn load_app_config() -> HashMap<String, String> {
    let path = data_file("app_config.csv");
    ensure_file(&path, APP_CONFIG_HEADERS);
    let mut config = HashMap::new();
    config.insert("location".to_string(), String::new());

    let mut rdr = match csv::Reader::from_path(&path) {
        Ok(r) => r,
        Err(_) => return config,
    };
    for result in rdr.records() {
        let record = match result {
            Ok(r) => r,
            Err(_) => continue,
        };
        let key = record.get(0).unwrap_or("").trim().to_string();
        let value = record.get(1).unwrap_or("").trim().to_string();
        if !key.is_empty() {
            config.insert(key, value);
        }
    }
    config
}

pub fn save_app_config(config: &HashMap<String, String>) {
    let path = data_file("app_config.csv");
    let rows: Vec<HashMap<String, String>> = config
        .iter()
        .map(|(k, v)| {
            let mut row = HashMap::new();
            row.insert("key".to_string(), k.clone());
            row.insert("value".to_string(), v.clone());
            row
        })
        .collect();
    atomic_write_dicts(&path, APP_CONFIG_HEADERS, &rows);
}

// ---------------------------------------------------------------------------
// Stock history
// ---------------------------------------------------------------------------

pub fn load_stock_history() -> HashMap<String, HashMap<String, i64>> {
    let path = data_file("stock_history.csv");
    ensure_file(&path, STOCK_HISTORY_HEADERS);
    let mut history: HashMap<String, HashMap<String, i64>> = HashMap::new();

    let mut rdr = match csv::Reader::from_path(&path) {
        Ok(r) => r,
        Err(_) => return history,
    };
    for result in rdr.records() {
        let record = match result {
            Ok(r) => r,
            Err(_) => continue,
        };
        let month = record.get(0).unwrap_or("").trim().to_string();
        let name = record.get(1).unwrap_or("").trim().to_string();
        let count_str = record.get(2).unwrap_or("0").trim();
        if month.is_empty() || name.is_empty() {
            continue;
        }
        let count = count_str.parse::<f64>().map(|c| (c as i64).max(0)).unwrap_or(0);
        history.entry(month).or_default().insert(name, count);
    }
    history
}

pub fn save_stock_snapshot(month: &str, stock_data: &HashMap<String, i64>) {
    let path = data_file("stock_history.csv");
    ensure_file(&path, STOCK_HISTORY_HEADERS);

    // Read existing, filter out current month
    let mut existing_rows: Vec<HashMap<String, String>> = Vec::new();
    if let Ok(mut rdr) = csv::Reader::from_path(&path) {
        for result in rdr.records() {
            if let Ok(record) = result {
                let m = record.get(0).unwrap_or("").trim();
                if m != month {
                    let mut row = HashMap::new();
                    for (i, h) in STOCK_HISTORY_HEADERS.iter().enumerate() {
                        row.insert(h.to_string(), record.get(i).unwrap_or("").to_string());
                    }
                    existing_rows.push(row);
                }
            }
        }
    }

    let timestamp = chrono::Local::now().format("%Y-%m-%dT%H:%M:%S").to_string();
    for (name, count) in stock_data {
        let mut row = HashMap::new();
        row.insert("month".to_string(), month.to_string());
        row.insert("textbook_name".to_string(), name.clone());
        row.insert("count".to_string(), count.to_string());
        row.insert("timestamp".to_string(), timestamp.clone());
        existing_rows.push(row);
    }

    atomic_write_dicts(&path, STOCK_HISTORY_HEADERS, &existing_rows);
}

// ---------------------------------------------------------------------------
// EPS records / audit
// ---------------------------------------------------------------------------

pub fn load_eps_records(date_str: &str) -> Vec<HashMap<String, String>> {
    let path = data_file("eps_records.csv");
    ensure_file(&path, EPS_RECORD_HEADERS);
    let mut results = Vec::new();
    if let Ok(mut rdr) = csv::Reader::from_path(&path) {
        for result in rdr.records() {
            if let Ok(record) = result {
                let d = record.get(0).unwrap_or("").trim();
                if d == date_str {
                    let mut row = HashMap::new();
                    for (i, h) in EPS_RECORD_HEADERS.iter().enumerate() {
                        row.insert(h.to_string(), record.get(i).unwrap_or("").to_string());
                    }
                    results.push(row);
                }
            }
        }
    }
    results
}

pub fn save_eps_records(date_str: &str, new_rows: &[HashMap<String, String>]) {
    let path = data_file("eps_records.csv");
    ensure_file(&path, EPS_RECORD_HEADERS);

    let mut kept: Vec<HashMap<String, String>> = Vec::new();
    if let Ok(mut rdr) = csv::Reader::from_path(&path) {
        for result in rdr.records() {
            if let Ok(record) = result {
                let d = record.get(0).unwrap_or("").trim();
                if d != date_str {
                    let mut row = HashMap::new();
                    for (i, h) in EPS_RECORD_HEADERS.iter().enumerate() {
                        row.insert(h.to_string(), record.get(i).unwrap_or("").to_string());
                    }
                    kept.push(row);
                }
            }
        }
    }
    kept.extend_from_slice(new_rows);
    atomic_write_dicts(&path, EPS_RECORD_HEADERS, &kept);
}

pub fn load_eps_audit(date_str: &str) -> Option<HashMap<String, String>> {
    let path = data_file("eps_audit.csv");
    ensure_file(&path, EPS_AUDIT_HEADERS);
    if let Ok(mut rdr) = csv::Reader::from_path(&path) {
        for result in rdr.records() {
            if let Ok(record) = result {
                let d = record.get(0).unwrap_or("").trim();
                if d == date_str {
                    let mut row = HashMap::new();
                    for (i, h) in EPS_AUDIT_HEADERS.iter().enumerate() {
                        row.insert(h.to_string(), record.get(i).unwrap_or("").to_string());
                    }
                    return Some(row);
                }
            }
        }
    }
    None
}

pub fn save_eps_audit(date_str: &str, audit_row: &HashMap<String, String>) {
    let path = data_file("eps_audit.csv");
    ensure_file(&path, EPS_AUDIT_HEADERS);
    let mut kept: Vec<HashMap<String, String>> = Vec::new();
    if let Ok(mut rdr) = csv::Reader::from_path(&path) {
        for result in rdr.records() {
            if let Ok(record) = result {
                let d = record.get(0).unwrap_or("").trim();
                if d != date_str {
                    let mut row = HashMap::new();
                    for (i, h) in EPS_AUDIT_HEADERS.iter().enumerate() {
                        row.insert(h.to_string(), record.get(i).unwrap_or("").to_string());
                    }
                    kept.push(row);
                }
            }
        }
    }
    kept.push(audit_row.clone());
    atomic_write_dicts(&path, EPS_AUDIT_HEADERS, &kept);
}

pub fn list_eps_dates() -> Vec<String> {
    let path = data_file("eps_records.csv");
    ensure_file(&path, EPS_RECORD_HEADERS);
    let mut dates = std::collections::HashSet::new();
    if let Ok(mut rdr) = csv::Reader::from_path(&path) {
        for result in rdr.records() {
            if let Ok(record) = result {
                let d = record.get(0).unwrap_or("").trim().to_string();
                if !d.is_empty() {
                    dates.insert(d);
                }
            }
        }
    }
    let mut sorted: Vec<String> = dates.into_iter().collect();
    sorted.sort();
    sorted
}

pub fn get_eps_after_total(date_str: &str) -> i64 {
    let mut total: i64 = 0;
    for row in load_eps_records(date_str) {
        if row.get("period").map(|s| s.trim()) == Some("after") {
            if let Some(s) = row.get("subtotal") {
                total += s.trim().parse::<f64>().map(|v| v as i64).unwrap_or(0);
            }
        }
    }
    total
}

pub fn get_eps_after_items(date_str: &str) -> HashMap<String, HashMap<String, i64>> {
    let mut result: HashMap<String, HashMap<String, i64>> = HashMap::new();
    for row in load_eps_records(date_str) {
        if row.get("period").map(|s| s.trim()) == Some("after") {
            let name = row.get("item_name").map(|s| s.trim().to_string()).unwrap_or_default();
            if !name.is_empty() {
                let mut qtys = HashMap::new();
                for key in &["qty_K", "qty_L", "qty_HK"] {
                    let v = row.get(*key).and_then(|s| s.trim().parse::<f64>().ok()).map(|v| v as i64).unwrap_or(0);
                    qtys.insert(key.to_string(), v);
                }
                result.insert(name, qtys);
            }
        }
    }
    result
}

pub fn has_eps_records_for_date(date_str: &str, period: &str) -> bool {
    for row in load_eps_records(date_str) {
        if row.get("period").map(|s| s.trim()) == Some(period) {
            return true;
        }
    }
    false
}
