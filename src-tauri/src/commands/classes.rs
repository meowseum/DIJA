use serde_json::{json, Value};
use tracing::info;
use uuid::Uuid;

use crate::auth::db::AuthDb;
use crate::auth::session::SessionStore;
use crate::models::{parse_date, parse_bool_loose, ClassRecord};
use crate::schedule::*;
use crate::sku::{build_sku, parse_sku};
use crate::storage::*;

fn build_schedule_with_index(
    class_record: &ClassRecord,
    holidays: &[crate::models::HolidayRange],
    postpones: &[crate::models::PostponeRecord],
    overrides: &[crate::models::LessonOverride],
) -> Vec<Value> {
    let start = parse_date(&class_record.start_date)
        .unwrap_or_else(|| chrono::Local::now().date_naive());
    let base = generate_weekly_schedule(start, class_record.weekday as u32, class_record.lesson_total, holidays);
    let schedule = apply_postpones(&base, class_record.weekday as u32, holidays, postpones);
    let mut schedule = apply_overrides(&schedule, holidays, overrides);
    schedule.sort();
    schedule
        .iter()
        .enumerate()
        .map(|(idx, d)| json!({"date": d.format("%Y-%m-%d").to_string(), "index": idx + 1}))
        .collect()
}

fn make_up_date_for(
    class_record: &ClassRecord,
    holidays: &[crate::models::HolidayRange],
    postpones: &[crate::models::PostponeRecord],
    overrides: &[crate::models::LessonOverride],
    original_date: chrono::NaiveDate,
) -> String {
    let start = parse_date(&class_record.start_date).unwrap_or(original_date);
    let base = generate_weekly_schedule(start, class_record.weekday as u32, class_record.lesson_total, holidays);
    let schedule = apply_postpones(&base, class_record.weekday as u32, holidays, postpones);
    let schedule = apply_overrides(&schedule, holidays, overrides);
    let mut scheduled: std::collections::HashSet<_> = schedule.into_iter().collect();
    scheduled.remove(&original_date);
    let holidays_set = holiday_set(holidays);

    let mut candidate = original_date + chrono::Duration::days(7);
    if scheduled.contains(&candidate) || holidays_set.contains(&candidate) {
        candidate = find_next_available_weekly(candidate, class_record.weekday as u32, &scheduled, &holidays_set);
    }
    candidate.format("%Y-%m-%d").to_string()
}

#[tauri::command]
pub fn create_class(
    session_token: String,
    data: Value,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "classes.create");

    let class_code = data.get("sku").and_then(|v| v.as_str()).unwrap_or("").trim().to_string();
    let sku_parts = match parse_sku(&class_code) {
        Some(p) => p,
        None => return json!({"ok": false, "error": "班別格式不正確。"}),
    };

    let start_date = data.get("start_date").and_then(|v| v.as_str()).unwrap_or("").trim().to_string();
    if parse_date(&start_date).is_none() {
        return json!({"ok": false, "error": "開課日期不正確。"});
    }

    let lesson_total = data.get("lesson_total").and_then(|v| v.as_i64()).unwrap_or(0);
    if lesson_total <= 0 {
        return json!({"ok": false, "error": "總課節必須大於 0。"});
    }

    let selected_level = data.get("level").and_then(|v| v.as_str()).unwrap_or("").trim().to_string();
    if selected_level.is_empty() {
        return json!({"ok": false, "error": "請選擇等級。"});
    }
    let sku_level = sku_parts["level"].as_str().unwrap_or("");
    if !sku_level.is_empty() && selected_level != sku_level {
        return json!({"ok": false, "error": "等級與班別不一致。"});
    }

    let app_location = load_app_config().get("location").cloned().unwrap_or_default().trim().to_uppercase();
    let sku_location = sku_parts["location"].as_str().unwrap_or("");
    let location_code = if !sku_location.is_empty() {
        sku_location.to_string()
    } else if ["K", "L", "H"].contains(&app_location.as_str()) {
        app_location
    } else {
        String::new()
    };

    let sku_value = build_sku(
        &selected_level,
        &location_code,
        sku_parts["start_month"].as_i64().unwrap_or(0),
        sku_parts["class_letter"].as_str().unwrap_or(""),
        sku_parts["start_year"].as_i64().unwrap_or(0),
    );

    let mut relay_teacher = data.get("relay_teacher").and_then(|v| v.as_str()).unwrap_or("").trim().to_string();
    let mut relay_date = data.get("relay_date").and_then(|v| v.as_str()).unwrap_or("").trim().to_string();
    if selected_level != "初級" {
        relay_teacher.clear();
        relay_date.clear();
    } else if !relay_teacher.is_empty() && relay_date.is_empty() {
        return json!({"ok": false, "error": "設有接力老師時，接力時間不能為空。"});
    } else if !relay_date.is_empty() && parse_date(&relay_date).is_none() {
        return json!({"ok": false, "error": "接力時間不正確。"});
    }

    let class_record = ClassRecord {
        id: Uuid::new_v4().to_string(),
        sku: sku_value.clone(),
        level: selected_level,
        location: location_code,
        start_month: sku_parts["start_month"].as_i64().unwrap_or(0),
        class_letter: sku_parts["class_letter"].as_str().unwrap_or("").to_string(),
        start_year: sku_parts["start_year"].as_i64().unwrap_or(0),
        classroom: data.get("classroom").and_then(|v| v.as_str()).unwrap_or("").trim().to_string(),
        start_date,
        weekday: data.get("weekday").and_then(|v| v.as_i64()).unwrap_or(0),
        start_time: data.get("start_time").and_then(|v| v.as_str()).unwrap_or("").trim().to_string(),
        teacher: data.get("teacher").and_then(|v| v.as_str()).unwrap_or("").trim().to_string(),
        relay_teacher,
        relay_date,
        student_count: data.get("student_count").and_then(|v| v.as_i64()).unwrap_or(0),
        lesson_total,
        status: "active".to_string(),
        doorplate_done: false,
        questionnaire_done: false,
        intro_done: false,
        merged_into_id: String::new(),
        promoted_to_id: String::new(),
        notes: String::new(),
    };

    let mut classes = load_classes();
    classes.push(class_record.clone());
    save_classes(&classes);
    info!("Class created: {} (id={})", sku_value, class_record.id);
    json!({"ok": true})
}

#[tauri::command]
pub fn update_class(
    session_token: String,
    class_id: String,
    updates: Value,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "classes.update");

    let mut classes = load_classes();
    let bool_fields = ["doorplate_done", "questionnaire_done", "intro_done"];
    let mut updated = false;

    for record in &mut classes {
        if record.id != class_id {
            continue;
        }

        // Handle SKU update
        if let Some(sku_val) = updates.get("sku").and_then(|v| v.as_str()) {
            let sku_value = sku_val.trim();
            if let Some(sku_parts) = parse_sku(sku_value) {
                let level_value = {
                    let l = sku_parts["level"].as_str().unwrap_or("");
                    if l.is_empty() { record.level.clone() } else { l.to_string() }
                };
                record.level = level_value.clone();
                record.location = sku_parts["location"].as_str().unwrap_or("").to_string();
                record.start_month = sku_parts["start_month"].as_i64().unwrap_or(0);
                record.class_letter = sku_parts["class_letter"].as_str().unwrap_or("").to_string();
                record.start_year = sku_parts["start_year"].as_i64().unwrap_or(0);
                record.sku = build_sku(
                    &level_value,
                    &record.location,
                    record.start_month,
                    &record.class_letter,
                    record.start_year,
                );
            } else {
                return json!({"ok": false, "error": "班別格式不正確。"});
            }
        }

        // Update other fields
        if let Some(obj) = updates.as_object() {
            for (field, value) in obj {
                if field == "sku" {
                    continue;
                }
                if bool_fields.contains(&field.as_str()) {
                    let b = value.as_str().map(|s| parse_bool_loose(s, false))
                        .or_else(|| value.as_bool())
                        .unwrap_or(false);
                    match field.as_str() {
                        "doorplate_done" => record.doorplate_done = b,
                        "questionnaire_done" => record.questionnaire_done = b,
                        "intro_done" => record.intro_done = b,
                        _ => {}
                    }
                } else if let Some(s) = value.as_str() {
                    match field.as_str() {
                        "classroom" => record.classroom = s.to_string(),
                        "start_date" => record.start_date = s.to_string(),
                        "start_time" => record.start_time = s.to_string(),
                        "teacher" => record.teacher = s.to_string(),
                        "relay_teacher" => record.relay_teacher = s.to_string(),
                        "relay_date" => record.relay_date = s.to_string(),
                        "status" => record.status = s.to_string(),
                        "merged_into_id" => record.merged_into_id = s.to_string(),
                        "promoted_to_id" => record.promoted_to_id = s.to_string(),
                        "notes" => record.notes = s.to_string(),
                        _ => {}
                    }
                } else if let Some(n) = value.as_i64() {
                    match field.as_str() {
                        "weekday" => record.weekday = n,
                        "student_count" => record.student_count = n,
                        "lesson_total" => record.lesson_total = n,
                        _ => {}
                    }
                }
            }
        }

        updated = true;
        break;
    }

    if updated {
        save_classes(&classes);
        info!("Class updated: {}", class_id);
    }
    json!({"ok": updated})
}

#[tauri::command]
pub fn delete_class(
    session_token: String,
    class_id: String,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "classes.delete");

    let classes = load_classes();
    if !classes.iter().any(|r| r.id == class_id) {
        return json!({"ok": false, "error": "找不到班別。"});
    }
    let new_classes: Vec<_> = classes.into_iter().filter(|r| r.id != class_id).collect();
    save_classes(&new_classes);
    info!("Class deleted: {}", class_id);

    let postpones = load_postpones();
    let new_postpones: Vec<_> = postpones.iter().filter(|p| p.class_id != class_id).cloned().collect();
    if new_postpones.len() != postpones.len() {
        save_postpones(&new_postpones);
    }

    let overrides = load_overrides();
    let new_overrides: Vec<_> = overrides.iter().filter(|o| o.class_id != class_id).cloned().collect();
    if new_overrides.len() != overrides.len() {
        save_overrides(&new_overrides);
    }

    json!({"ok": true})
}

#[tauri::command]
pub fn end_class_action(
    session_token: String,
    class_id: String,
    action: String,
    target_id: Option<String>,
    new_sku: Option<String>,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "classes.end");

    let mut classes = load_classes();
    let mut updated = false;
    let target_id = target_id.unwrap_or_default();
    let new_sku = new_sku.unwrap_or_default();

    match action.as_str() {
        "terminate" => {
            for record in &mut classes {
                if record.id == class_id {
                    record.status = "terminated".to_string();
                    updated = true;
                    break;
                }
            }
        }
        "merge" => {
            if !classes.iter().any(|c| c.id == target_id) {
                return json!({"ok": false, "error": "找不到合併目標班別。"});
            }
            for record in &mut classes {
                if record.id == class_id {
                    record.status = "merged".to_string();
                    record.merged_into_id = target_id.clone();
                    updated = true;
                    break;
                }
            }
        }
        "promote" => {
            let sku_parts = match parse_sku(&new_sku) {
                Some(p) => p,
                None => return json!({"ok": false, "error": "升級班別格式不正確。"}),
            };
            let sku_level = sku_parts["level"].as_str().unwrap_or("");
            if sku_level.is_empty() {
                return json!({"ok": false, "error": "升級班別需要包含等級。"});
            }
            let base = match classes.iter().find(|c| c.id == class_id) {
                Some(c) => c.clone(),
                None => return json!({"ok": false, "error": "找不到班別。"}),
            };
            let promoted_id = Uuid::new_v4().to_string();
            let code = sku_parts["code"].as_str().unwrap_or("");
            let promoted = ClassRecord {
                id: promoted_id.clone(),
                sku: format!("{}{}", sku_level, code),
                level: sku_level.to_string(),
                location: sku_parts["location"].as_str().unwrap_or("").to_string(),
                start_month: sku_parts["start_month"].as_i64().unwrap_or(0),
                class_letter: sku_parts["class_letter"].as_str().unwrap_or("").to_string(),
                start_year: sku_parts["start_year"].as_i64().unwrap_or(0),
                classroom: base.classroom.clone(),
                start_date: base.start_date.clone(),
                weekday: base.weekday,
                start_time: base.start_time.clone(),
                teacher: base.teacher.clone(),
                relay_teacher: base.relay_teacher.clone(),
                relay_date: base.relay_date.clone(),
                student_count: base.student_count,
                lesson_total: base.lesson_total,
                status: "active".to_string(),
                doorplate_done: base.doorplate_done,
                questionnaire_done: base.questionnaire_done,
                intro_done: base.intro_done,
                merged_into_id: String::new(),
                promoted_to_id: String::new(),
                notes: String::new(),
            };
            for record in &mut classes {
                if record.id == class_id {
                    record.status = "promoted".to_string();
                    record.promoted_to_id = promoted_id.clone();
                    updated = true;
                    break;
                }
            }
            if updated {
                classes.push(promoted);
            }
        }
        _ => {}
    }

    if updated {
        save_classes(&classes);
    }
    json!({"ok": updated})
}

#[tauri::command]
pub fn save_student_counts(
    session_token: String,
    updates: Vec<Value>,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "classes.update");

    let mut classes = load_classes();
    let mut class_map: std::collections::HashMap<String, &mut ClassRecord> =
        classes.iter_mut().map(|c| (c.id.clone(), c)).collect();

    for item in &updates {
        let class_id = item.get("id").and_then(|v| v.as_str()).unwrap_or("").trim().to_string();
        let count = item.get("student_count").and_then(|v| v.as_i64());
        if let (Some(record), Some(count)) = (class_map.get_mut(&class_id), count) {
            record.student_count = count.max(0);
        }
    }

    let records: Vec<ClassRecord> = classes;
    save_classes(&records);
    json!({"ok": true})
}

#[tauri::command]
pub fn terminate_class_with_last_date(
    session_token: String,
    class_id: String,
    last_date: String,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "classes.end");

    let end_date = match parse_date(&last_date) {
        Some(d) => d,
        None => return json!({"ok": false, "error": "日期不正確。"}),
    };
    let mut classes = load_classes();
    let class_record = match classes.iter().find(|c| c.id == class_id) {
        Some(c) => c.clone(),
        None => return json!({"ok": false, "error": "找不到班別。"}),
    };
    let holidays = load_holidays();
    let postpones = load_postpones();
    let overrides = load_overrides();
    let schedule = build_schedule_with_index(&class_record, &holidays, &postpones, &overrides);
    let count = schedule.iter().filter(|item| {
        item.get("date").and_then(|v| v.as_str()).and_then(|d| parse_date(d)).map(|d| d <= end_date).unwrap_or(false)
    }).count() as i64;

    if count <= 0 {
        return json!({"ok": false, "error": "結束日期早於開課日期。"});
    }

    for record in &mut classes {
        if record.id == class_id {
            record.lesson_total = count;
            record.status = "terminated".to_string();
            break;
        }
    }
    save_classes(&classes);
    json!({"ok": true})
}

#[tauri::command]
pub fn get_class_schedule(
    session_token: String,
    class_id: String,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "classes.view");

    let classes = load_classes();
    let class_record = match classes.iter().find(|c| c.id == class_id) {
        Some(c) => c,
        None => return json!({"ok": false, "error": "找不到班別。"}),
    };
    let holidays = load_holidays();
    let postpones = load_postpones();
    let overrides = load_overrides();
    let class_postpones: Vec<_> = postpones.iter().filter(|p| p.class_id == class_id).cloned().collect();
    let class_overrides: Vec<_> = overrides.iter().filter(|o| o.class_id == class_id).cloned().collect();
    let make_up_dates: std::collections::HashSet<String> = class_postpones.iter().map(|p| p.make_up_date.clone()).collect();
    let schedule_with_index = build_schedule_with_index(class_record, &holidays, &class_postpones, &class_overrides);
    let schedule_payload: Vec<Value> = schedule_with_index.iter().map(|item| {
        let date_str = item.get("date").and_then(|v| v.as_str()).unwrap_or("");
        let is_makeup = make_up_dates.contains(date_str);
        json!({
            "date": date_str,
            "type": if is_makeup { "makeup" } else { "normal" },
            "index": item.get("index").and_then(|v| v.as_i64()).unwrap_or(0),
            "total": class_record.lesson_total,
        })
    }).collect();

    json!({
        "ok": true,
        "class": serde_json::to_value(class_record).unwrap_or(json!({})),
        "schedule": schedule_payload,
        "postpones": class_postpones.iter().map(|p| serde_json::to_value(p).unwrap_or(json!({}))).collect::<Vec<_>>(),
        "overrides": class_overrides.iter().map(|o| serde_json::to_value(o).unwrap_or(json!({}))).collect::<Vec<_>>(),
    })
}

// Re-export helpers used by other command modules
pub fn build_schedule_with_index_pub(
    class_record: &ClassRecord,
    holidays: &[crate::models::HolidayRange],
    postpones: &[crate::models::PostponeRecord],
    overrides: &[crate::models::LessonOverride],
) -> Vec<Value> {
    build_schedule_with_index(class_record, holidays, postpones, overrides)
}

pub fn make_up_date_for_pub(
    class_record: &ClassRecord,
    holidays: &[crate::models::HolidayRange],
    postpones: &[crate::models::PostponeRecord],
    overrides: &[crate::models::LessonOverride],
    original_date: chrono::NaiveDate,
) -> String {
    make_up_date_for(class_record, holidays, postpones, overrides, original_date)
}
