use serde_json::{json, Value};
use tracing::info;
use uuid::Uuid;

use crate::auth::db::AuthDb;
use crate::auth::session::SessionStore;
use crate::models::{parse_date, PostponeRecord};
use crate::schedule::*;
use crate::storage::*;
use super::classes::{build_schedule_with_index_pub, make_up_date_for_pub};

#[tauri::command]
pub fn add_postpone(
    session_token: String,
    class_id: String,
    original_date: String,
    reason: String,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "postpones.add");

    let mut classes = load_classes();
    let class_record = match classes.iter().find(|c| c.id == class_id) {
        Some(c) => c.clone(),
        None => return json!({"ok": false, "error": "找不到班別。"}),
    };

    let original_dt = match parse_date(&original_date) {
        Some(d) => d,
        None => return json!({"ok": false, "error": "原定日期不正確。"}),
    };

    let holidays = load_holidays();
    let mut postpones = load_postpones();
    let overrides = load_overrides();
    let schedule = build_schedule_with_index_pub(&class_record, &holidays, &postpones, &overrides);
    let schedule_dates: std::collections::HashSet<String> = schedule.iter()
        .filter_map(|item| item.get("date").and_then(|v| v.as_str()).map(|s| s.to_string()))
        .collect();

    let (make_up_date, reactivated) = if !schedule_dates.contains(&original_date) {
        let progress = calculate_progress(
            &class_record.start_date, class_record.weekday as u32, class_record.lesson_total,
            &holidays, &postpones.iter().filter(|p| p.class_id == class_id).cloned().collect::<Vec<_>>(),
            &overrides.iter().filter(|o| o.class_id == class_id).cloned().collect::<Vec<_>>(),
        );
        let is_ended = progress["lessons_remaining"].as_i64().unwrap_or(0) == 0
            || class_record.status == "terminated";
        if !is_ended {
            return json!({"ok": false, "error": "原定日期不在日程。"});
        }
        let mud = make_up_date_for_pub(&class_record, &holidays, &postpones, &overrides, original_dt);
        // Increment lesson_total and possibly reactivate
        for record in &mut classes {
            if record.id == class_id {
                record.lesson_total += 1;
                let reactivated = record.status == "terminated";
                if reactivated {
                    record.status = "active".to_string();
                    info!("Class {} reactivated via add_postpone", class_id);
                }
                save_classes(&classes);
                return finish_add_postpone(&class_id, &original_date, &reason, &mud, reactivated, &mut postpones);
            }
        }
        (mud, false)
    } else {
        let mud = make_up_date_for_pub(&class_record, &holidays, &postpones, &overrides, original_dt);
        (mud, false)
    };

    finish_add_postpone(&class_id, &original_date, &reason, &make_up_date, reactivated, &mut postpones)
}

fn finish_add_postpone(
    class_id: &str, original_date: &str, reason: &str, make_up_date: &str,
    reactivated: bool, postpones: &mut Vec<PostponeRecord>,
) -> Value {
    let postpone = PostponeRecord {
        id: Uuid::new_v4().to_string(),
        class_id: class_id.to_string(),
        original_date: original_date.to_string(),
        reason: reason.trim().to_string(),
        make_up_date: make_up_date.to_string(),
    };
    postpones.push(postpone);
    save_postpones(postpones);
    info!("Postpone added for class {}: {} → {}", class_id, original_date, make_up_date);
    json!({"ok": true, "reactivated": reactivated})
}

#[tauri::command]
pub fn add_postpone_manual(
    session_token: String,
    class_id: String,
    original_date: String,
    make_up_date: String,
    reason: String,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "postpones.add");

    let mut classes = load_classes();
    let class_record = match classes.iter().find(|c| c.id == class_id) {
        Some(c) => c.clone(),
        None => return json!({"ok": false, "error": "找不到班別。"}),
    };
    if parse_date(&original_date).is_none() {
        return json!({"ok": false, "error": "原定日期不正確。"});
    }
    if parse_date(&make_up_date).is_none() {
        return json!({"ok": false, "error": "補課日期不正確。"});
    }

    let holidays = load_holidays();
    let mut postpones = load_postpones();
    let overrides = load_overrides();
    let schedule = build_schedule_with_index_pub(&class_record, &holidays, &postpones, &overrides);
    let schedule_dates: std::collections::HashSet<String> = schedule.iter()
        .filter_map(|item| item.get("date").and_then(|v| v.as_str()).map(|s| s.to_string()))
        .collect();

    let reactivated;
    if !schedule_dates.contains(&original_date) {
        let progress = calculate_progress(
            &class_record.start_date, class_record.weekday as u32, class_record.lesson_total,
            &holidays, &postpones.iter().filter(|p| p.class_id == class_id).cloned().collect::<Vec<_>>(),
            &overrides.iter().filter(|o| o.class_id == class_id).cloned().collect::<Vec<_>>(),
        );
        let is_ended = progress["lessons_remaining"].as_i64().unwrap_or(0) == 0
            || class_record.status == "terminated";
        if !is_ended {
            return json!({"ok": false, "error": "原定日期不在日程。"});
        }
        if schedule_dates.contains(&make_up_date) {
            return json!({"ok": false, "error": "補課日期重複。"});
        }
        if let Some(d) = parse_date(&make_up_date) {
            if holiday_set(&holidays).contains(&d) {
                return json!({"ok": false, "error": "補課日期遇到假期。"});
            }
        }
        for record in &mut classes {
            if record.id == class_id {
                record.lesson_total += 1;
                reactivated = record.status == "terminated";
                if reactivated {
                    record.status = "active".to_string();
                    info!("Class {} reactivated via add_postpone_manual", class_id);
                }
                save_classes(&classes);
                return finish_add_postpone(&class_id, &original_date, &reason, &make_up_date, reactivated, &mut postpones);
            }
        }
        reactivated = false;
    } else {
        reactivated = false;
    }

    finish_add_postpone(&class_id, &original_date, &reason, &make_up_date, reactivated, &mut postpones)
}

#[tauri::command]
pub fn get_make_up_date(
    session_token: String,
    class_id: String,
    original_date: String,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "postpones.view");

    let classes = load_classes();
    let class_record = match classes.iter().find(|c| c.id == class_id) {
        Some(c) => c,
        None => return json!({"ok": false, "error": "找不到班別。"}),
    };
    let original_dt = match parse_date(&original_date) {
        Some(d) => d,
        None => return json!({"ok": false, "error": "原定日期不正確。"}),
    };
    let holidays = load_holidays();
    let postpones = load_postpones();
    let overrides = load_overrides();
    let schedule = build_schedule_with_index_pub(class_record, &holidays, &postpones, &overrides);
    let schedule_dates: std::collections::HashSet<String> = schedule.iter()
        .filter_map(|item| item.get("date").and_then(|v| v.as_str()).map(|s| s.to_string()))
        .collect();

    if !schedule_dates.contains(&original_date) {
        let progress = calculate_progress(
            &class_record.start_date, class_record.weekday as u32, class_record.lesson_total,
            &holidays, &postpones.iter().filter(|p| p.class_id == class_id).cloned().collect::<Vec<_>>(),
            &overrides.iter().filter(|o| o.class_id == class_id).cloned().collect::<Vec<_>>(),
        );
        let is_ended = progress["lessons_remaining"].as_i64().unwrap_or(0) == 0
            || class_record.status == "terminated";
        if !is_ended {
            return json!({"ok": false, "error": "原定日期不在日程。"});
        }
    }

    let mud = make_up_date_for_pub(class_record, &holidays, &postpones, &overrides, original_dt);
    json!({"ok": true, "make_up_date": mud})
}

#[tauri::command]
pub fn delete_postpone(
    session_token: String,
    postpone_id: String,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "postpones.delete");

    let postpones = load_postpones();
    let new_postpones: Vec<_> = postpones.into_iter().filter(|p| p.id != postpone_id).collect();
    save_postpones(&new_postpones);
    json!({"ok": true})
}
