use chrono::Datelike;
use regex::Regex;
use serde_json::{json, Value};
use std::collections::HashMap;
use std::sync::LazyLock;

use crate::auth::db::AuthDb;
use crate::auth::session::SessionStore;
use crate::config::{get_output_dir, get_template_dir};
use crate::docx::{extract_docx_text, render_docx_template};
use crate::models::parse_date;
use crate::schedule::*;
use crate::sku::parse_sku;
use crate::storage::*;

static LOCATION_NAME: LazyLock<HashMap<&'static str, &'static str>> = LazyLock::new(|| {
    let mut m = HashMap::new();
    m.insert("K", "旺角校");
    m.insert("L", "太子校");
    m.insert("H", "香港校");
    m
});

const WEEKDAYS_ZH: &[&str] = &["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"];
const WEEKDAYS_JP: &[&str] = &["月曜日", "火曜日", "水曜日", "木曜日", "金曜日", "土曜日", "日曜日"];
const WEEKDAYS_EN: &[&str] = &["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"];

fn format_class_name(record: &crate::models::ClassRecord) -> String {
    if let Some(parts) = parse_sku(&record.sku) {
        let year_short = format!("{:02}", parts["start_year"].as_i64().unwrap_or(0) % 100);
        let prefix = if record.level.starts_with('N') { "" } else { "N" };
        format!(
            "{}{}{}{}{}/{}",
            prefix,
            record.level,
            parts["location"].as_str().unwrap_or(""),
            parts["start_month"].as_i64().unwrap_or(0),
            parts["class_letter"].as_str().unwrap_or(""),
            year_short
        )
    } else {
        record.sku.clone()
    }
}

static TIME_RE: LazyLock<Regex> = LazyLock::new(|| {
    Regex::new(r"(\d{1,2}):?(\d{2})\s*[-~]\s*(\d{1,2}):?(\d{2})").unwrap()
});

fn format_class_time(time_value: &str) -> String {
    if let Some(caps) = TIME_RE.captures(time_value) {
        let label = |h: &str, m: &str| -> String {
            let hour: u32 = h.parse().unwrap_or(0);
            let minute: u32 = m.parse().unwrap_or(0);
            let is_pm = hour >= 12;
            let hour12 = ((hour + 11) % 12) + 1;
            format!("{}:{:02}{}", hour12, minute, if is_pm { "pm" } else { "am" })
        };
        format!("{} ~ {}", label(&caps[1], &caps[2]), label(&caps[3], &caps[4]))
    } else {
        time_value.to_string()
    }
}

fn format_class_time_zh(time_value: &str) -> String {
    if let Some(caps) = TIME_RE.captures(time_value) {
        let zh_label = |h: &str, m: &str| -> (String, String) {
            let hour: u32 = h.parse().unwrap_or(0);
            let minute: u32 = m.parse().unwrap_or(0);
            let prefix = if hour >= 12 { "下午" } else { "上午" };
            let hour12 = ((hour + 11) % 12) + 1;
            (prefix.to_string(), format!("{}:{:02}", hour12, minute))
        };
        let (prefix1, t1) = zh_label(&caps[1], &caps[2]);
        let (_, t2) = zh_label(&caps[3], &caps[4]);
        format!("{}{}～{}", prefix1, t1, t2)
    } else {
        time_value.to_string()
    }
}

fn monthly_hours_from_time(time_value: &str) -> i64 {
    if let Some(caps) = TIME_RE.captures(time_value) {
        let start_m: i64 = caps[1].parse::<i64>().unwrap_or(0) * 60 + caps[2].parse::<i64>().unwrap_or(0);
        let end_m: i64 = caps[3].parse::<i64>().unwrap_or(0) * 60 + caps[4].parse::<i64>().unwrap_or(0);
        let duration = end_m - start_m;
        if duration <= 0 { return 0; }
        let hours_per_lesson = duration as f64 / 60.0;
        (hours_per_lesson * 4.0).round() as i64
    } else {
        0
    }
}

fn teacher_for(record: &crate::models::ClassRecord, use_relay: bool) -> String {
    if use_relay && record.level == "初級" && !record.relay_teacher.is_empty() {
        record.relay_teacher.clone()
    } else {
        record.teacher.clone()
    }
}

#[tauri::command]
pub fn list_docx_templates(
    session_token: String,
    app_handle: tauri::AppHandle,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "documents.view");

    let template_dir = get_template_dir(&app_handle).join("print");
    if !template_dir.exists() {
        return json!({"ok": false, "error": "找不到模板資料夾。", "templates": []});
    }
    let mut templates: Vec<String> = Vec::new();
    if let Ok(entries) = std::fs::read_dir(&template_dir) {
        for entry in entries.flatten() {
            let path = entry.path();
            if path.is_file() && path.extension().map(|e| e == "docx").unwrap_or(false) {
                if let Some(name) = path.file_name().and_then(|n| n.to_str()) {
                    templates.push(name.to_string());
                }
            }
        }
    }
    templates.sort();
    json!({"ok": true, "templates": templates})
}

#[tauri::command]
pub fn generate_docx(
    session_token: String,
    app_handle: tauri::AppHandle,
    template_name: String,
    class_id: String,
    use_relay_teacher: Option<bool>,
    class_id_secondary: Option<String>,
    use_relay_teacher_secondary: Option<bool>,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "documents.generate");

    let use_relay = use_relay_teacher.unwrap_or(false);
    let class_id_secondary = class_id_secondary.unwrap_or_default();
    let use_relay_secondary = use_relay_teacher_secondary.unwrap_or(false);

    let template_dir = get_template_dir(&app_handle).join("print");
    let output_dir = get_output_dir();
    let template_path = template_dir.join(&template_name);
    if !template_path.exists() {
        return json!({"ok": false, "error": "找不到模板。"});
    }

    let classes = load_classes();
    let class_record = match classes.iter().find(|c| c.id == class_id) {
        Some(c) => c,
        None => return json!({"ok": false, "error": "找不到班別。"}),
    };

    let weekday_index = (class_record.weekday as usize).min(6);
    let mut context: HashMap<String, String> = HashMap::new();

    if template_name == "class.docx" {
        context.insert("CLASS_NAME_1".to_string(), format_class_name(class_record));
        context.insert("CLASS_WEEK_1".to_string(), WEEKDAYS_EN[weekday_index].to_string());
        context.insert("TEACHER_NAME_1".to_string(), teacher_for(class_record, use_relay));
        context.insert("CLASS_TIME_1".to_string(), format_class_time(&class_record.start_time));
        context.insert("CLASS_NAME_2".to_string(), String::new());
        context.insert("CLASS_WEEK_2".to_string(), String::new());
        context.insert("TEACHER_NAME_2".to_string(), String::new());
        context.insert("CLASS_TIME_2".to_string(), String::new());

        if !class_id_secondary.is_empty() {
            let secondary = match classes.iter().find(|c| c.id == class_id_secondary) {
                Some(c) => c,
                None => return json!({"ok": false, "error": "找不到追加班別。"}),
            };
            let sec_weekday = (secondary.weekday as usize).min(6);
            context.insert("CLASS_NAME_2".to_string(), format_class_name(secondary));
            context.insert("CLASS_WEEK_2".to_string(), WEEKDAYS_EN[sec_weekday].to_string());
            context.insert("TEACHER_NAME_2".to_string(), teacher_for(secondary, use_relay_secondary));
            context.insert("CLASS_TIME_2".to_string(), format_class_time(&secondary.start_time));
        }
    } else {
        context.insert("CLASS_NAME".to_string(), format_class_name(class_record));
        context.insert("CLASS_WEEK".to_string(), WEEKDAYS_ZH[weekday_index].to_string());
        context.insert("TEACHER_NAME".to_string(), teacher_for(class_record, use_relay));
        context.insert("CLASS_TIME".to_string(), format_class_time(&class_record.start_time));
        context.insert("ROOM_NUMBER".to_string(), class_record.classroom.clone());
        context.insert("WEEK_DAY".to_string(), WEEKDAYS_JP[weekday_index].to_string());
    }

    std::fs::create_dir_all(&output_dir).ok();
    let stamp = chrono::Local::now().format("%Y%m%d").to_string();
    let stem = std::path::Path::new(&template_name).file_stem().and_then(|s| s.to_str()).unwrap_or("doc");
    let output_name = format!("{}_{}_{}. docx", class_record.sku, stem, stamp);
    let output_path = output_dir.join(&output_name);

    match render_docx_template(&template_path, &output_path, &context) {
        Ok(()) => json!({"ok": true, "path": output_path.to_string_lossy()}),
        Err(e) => json!({"ok": false, "error": e}),
    }
}

#[tauri::command]
pub fn load_payment_template(
    session_token: String,
    app_handle: tauri::AppHandle,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "documents.view");

    let template_dir = get_template_dir(&app_handle);
    let candidates = [
        template_dir.join("messages").join("payinstruc_wtsapp.txt"),
        template_dir.join("payinstruc_wtsapp.txt"),
    ];
    for path in &candidates {
        if path.exists() {
            match std::fs::read_to_string(path) {
                Ok(content) => return json!({"ok": true, "content": content}),
                Err(_) => continue,
            }
        }
    }
    let docx_path = template_dir.join("payinstruc_wtsapp.docx");
    if docx_path.exists() {
        match extract_docx_text(&docx_path) {
            Ok(content) => return json!({"ok": true, "content": content}),
            Err(e) => return json!({"ok": false, "error": e}),
        }
    }
    json!({"ok": false, "error": "找不到模板。"})
}

#[tauri::command]
pub fn load_makeup_template(
    session_token: String,
    app_handle: tauri::AppHandle,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "documents.view");

    let template_dir = get_template_dir(&app_handle);
    let path = template_dir.join("messages").join("補堂注意事項及安排.docx");
    if path.exists() {
        match extract_docx_text(&path) {
            Ok(content) => return json!({"ok": true, "content": content}),
            Err(e) => return json!({"ok": false, "error": e}),
        }
    }
    json!({"ok": false, "error": "找不到模板。"})
}

#[tauri::command]
pub fn get_promote_notice_data(
    session_token: String,
    class_id: String,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "documents.view");

    let classes = load_classes();
    let record = match classes.iter().find(|c| c.id == class_id) {
        Some(c) => c,
        None => return json!({"ok": false, "error": "找不到班別。"}),
    };

    let settings = load_settings();
    let holidays = load_holidays();
    let postpones = load_postpones();
    let overrides = load_overrides();

    let start = parse_date(&record.start_date);
    let mut end_date_str = String::new();
    if let Some(start) = start {
        let class_postpones: Vec<_> = postpones.iter().filter(|p| p.class_id == record.id).cloned().collect();
        let class_overrides: Vec<_> = overrides.iter().filter(|o| o.class_id == record.id).cloned().collect();
        let schedule = generate_weekly_schedule(start, record.weekday as u32, record.lesson_total, &holidays);
        let schedule = apply_postpones(&schedule, record.weekday as u32, &holidays, &class_postpones);
        let schedule = apply_overrides(&schedule, &holidays, &class_overrides);
        if let Some(end) = schedule.last() {
            end_date_str = format!("{}年{}月{}日", end.year(), end.month(), end.day());
        }
    }

    let start_date_formatted = start.map(|s| {
        let weekday_index = (record.weekday as usize).min(6);
        format!("{}月{}日{}", s.month(), s.day(), WEEKDAYS_ZH[weekday_index])
    }).unwrap_or_default();

    let duration = start.map(|s| {
        if end_date_str.is_empty() { String::new() }
        else { format!("{}年{}月{}日至{}", s.year(), s.month(), s.day(), end_date_str) }
    }).unwrap_or_default();

    let time_zh = format_class_time_zh(&record.start_time);
    let teacher = if record.teacher.is_empty() { String::new() } else { format!("＜{}＞", record.teacher) };
    let location = LOCATION_NAME.get(record.location.as_str()).unwrap_or(&record.location.as_str()).to_string();

    let hours = monthly_hours_from_time(&record.start_time);
    let hours_str = if hours > 0 { format!("每月{}小時", hours) } else { "每月--小時".to_string() };

    let price = settings.level_price.get(&record.level);
    let price_str = match price {
        Some(p) => format!("學費${}", p),
        None => "學費$--".to_string(),
    };
    let remarks = format!("{} {}", hours_str, price_str);
    let signature_date = format!("{:02}/{}", record.start_month, record.start_year);

    let source_level = settings.level_next.iter()
        .find(|(_, v)| v.as_str() == record.level)
        .map(|(k, _)| k.clone())
        .unwrap_or_default();

    let short = |lvl: &str| -> String {
        lvl.trim_end_matches('級').to_string()
    };

    json!({
        "ok": true,
        "name": format_class_name(record),
        "start_date_formatted": start_date_formatted,
        "duration": duration,
        "time": time_zh,
        "teacher": teacher,
        "location": location,
        "remarks": remarks,
        "signature_date": signature_date,
        "source_level": source_level,
        "target_level": record.level,
        "source_level_short": short(&source_level),
        "target_level_short": short(&record.level),
        "start_month": record.start_month,
    })
}

#[tauri::command]
pub fn generate_promote_notice(
    session_token: String,
    app_handle: tauri::AppHandle,
    data: Value,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "documents.generate");

    let template_path = get_template_dir(&app_handle).join("print").join("promote_notice.docx");
    if !template_path.exists() {
        return json!({"ok": false, "error": "找不到升班通知模板 (promote_notice.docx)。請先生成模板。"});
    }

    let output_dir = get_output_dir();
    std::fs::create_dir_all(&output_dir).ok();

    let sku = data.get("sku").and_then(|v| v.as_str()).unwrap_or("promote").replace('/', "-");
    let stamp = chrono::Local::now().format("%Y%m%d").to_string();
    let output_name = format!("{}_promote_notice_{}.docx", sku, stamp);
    let output_path = output_dir.join(&output_name);

    let mut context: HashMap<String, String> = HashMap::new();
    for key in &["ADDRESSEE", "BODY_TEXT", "CLASS_NAME", "START_DATE", "DURATION", "TIME", "TEACHER", "LOCATION", "REMARKS", "TEXTBOOK_FEE", "SIGNATURE_DATE"] {
        let json_key = match *key {
            "CLASS_NAME" => "name",
            "START_DATE" => "start_date_formatted",
            "DURATION" => "duration",
            "TIME" => "time",
            "TEACHER" => "teacher",
            "LOCATION" => "location",
            "REMARKS" => "remarks",
            "TEXTBOOK_FEE" => "textbook_fee",
            "SIGNATURE_DATE" => "signature_date",
            "ADDRESSEE" => "addressee",
            "BODY_TEXT" => "body_text",
            _ => *key,
        };
        context.insert(key.to_string(), data.get(json_key).and_then(|v| v.as_str()).unwrap_or("").to_string());
    }

    match render_docx_template(&template_path, &output_path, &context) {
        Ok(()) => json!({"ok": true, "path": output_path.to_string_lossy()}),
        Err(e) => json!({"ok": false, "error": e}),
    }
}

#[tauri::command]
pub fn list_message_templates(
    session_token: String,
    app_handle: tauri::AppHandle,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "documents.view");

    let template_dir = get_template_dir(&app_handle).join("messages");
    if !template_dir.exists() {
        return json!({"ok": false, "error": "找不到訊息資料夾。", "templates": []});
    }
    let settings = load_settings();
    let category_map = &settings.message_category;
    let mut templates: Vec<Value> = Vec::new();

    if let Ok(mut entries) = std::fs::read_dir(&template_dir) {
        let mut paths: Vec<_> = entries.by_ref().flatten().map(|e| e.path()).collect();
        paths.sort();
        for path in paths {
            if !path.is_file() { continue; }
            let ext = path.extension().and_then(|e| e.to_str()).unwrap_or("").to_lowercase();
            if ext != "txt" && ext != "docx" { continue; }
            let name = path.file_name().and_then(|n| n.to_str()).unwrap_or("").to_string();
            if name == "補堂注意事項及安排.docx" { continue; }
            let label = path.file_stem().and_then(|s| s.to_str()).unwrap_or("").to_string();
            let category = category_map.get(&name).cloned().unwrap_or_default();
            templates.push(json!({"name": name, "label": label, "category": category}));
        }
    }

    json!({"ok": true, "templates": templates})
}

#[tauri::command]
pub fn load_message_content(
    session_token: String,
    app_handle: tauri::AppHandle,
    template_name: String,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "documents.view");

    let template_dir = get_template_dir(&app_handle).join("messages");
    let template_path = template_dir.join(&template_name);
    if !template_path.exists() {
        return json!({"ok": false, "error": "找不到訊息模板。"});
    }
    let ext = template_path.extension().and_then(|e| e.to_str()).unwrap_or("").to_lowercase();
    if ext == "txt" {
        match std::fs::read_to_string(&template_path) {
            Ok(content) => return json!({"ok": true, "content": content}),
            Err(e) => return json!({"ok": false, "error": format!("讀取失敗: {}", e)}),
        }
    }
    if ext == "docx" {
        match extract_docx_text(&template_path) {
            Ok(content) => return json!({"ok": true, "content": content}),
            Err(e) => return json!({"ok": false, "error": e}),
        }
    }
    json!({"ok": false, "error": "不支援的檔案格式。"})
}

#[tauri::command]
pub fn set_message_category(
    session_token: String,
    template_name: String,
    category: String,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "settings.modify");

    let name = template_name.trim().to_string();
    if name.is_empty() {
        return json!({"ok": false, "error": "模板名稱不正確。"});
    }
    let mut settings = load_settings();
    settings.message_category.insert(name, category.trim().to_string());
    save_settings(&settings);
    json!({"ok": true})
}
