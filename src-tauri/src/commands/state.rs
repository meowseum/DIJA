use serde_json::{json, Value};
use tracing::{info, warn};

use crate::auth::db::AuthDb;
use crate::auth::session::SessionStore;
use crate::models::parse_date;
use crate::schedule::calculate_progress;
use crate::storage::*;

fn class_progress(
    record: &crate::models::ClassRecord,
    holidays: &[crate::models::HolidayRange],
    postpones: &[crate::models::PostponeRecord],
    overrides: &[crate::models::LessonOverride],
) -> Value {
    let class_postpones: Vec<_> = postpones.iter().filter(|p| p.class_id == record.id).cloned().collect();
    let class_overrides: Vec<_> = overrides.iter().filter(|o| o.class_id == record.id).cloned().collect();
    calculate_progress(
        &record.start_date,
        record.weekday as u32,
        record.lesson_total,
        holidays,
        &class_postpones,
        &class_overrides,
    )
}

fn class_payload(
    record: &crate::models::ClassRecord,
    holidays: &[crate::models::HolidayRange],
    postpones: &[crate::models::PostponeRecord],
    overrides: &[crate::models::LessonOverride],
) -> Value {
    let progress = class_progress(record, holidays, postpones, overrides);
    let mut payload = serde_json::to_value(record).unwrap_or(json!({}));
    if let (Some(obj), Some(prog_obj)) = (payload.as_object_mut(), progress.as_object()) {
        for (k, v) in prog_obj {
            obj.insert(k.clone(), v.clone());
        }
    }
    payload
}

#[tauri::command]
pub fn load_state(
    session_token: String,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "state.view");

    let mut classes = load_classes();
    let holidays = load_holidays();
    let postpones = load_postpones();
    let settings = load_settings();
    let app_config = load_app_config();
    let overrides = load_overrides();

    // Data integrity checks
    let class_ids: std::collections::HashSet<_> = classes.iter().map(|c| c.id.clone()).collect();
    for p in &postpones {
        if !class_ids.contains(&p.class_id) {
            warn!("Orphaned postpone {} references missing class {}", p.id, p.class_id);
        }
    }
    for o in &overrides {
        if !class_ids.contains(&o.class_id) {
            warn!("Orphaned override {} references missing class {}", o.id, o.class_id);
        }
    }
    for c in &classes {
        if !c.start_date.is_empty() && parse_date(&c.start_date).is_none() {
            warn!("Class {} has invalid start_date: {}", c.id, c.start_date);
        }
        if !(0..=6).contains(&c.weekday) {
            warn!("Class {} has out-of-range weekday: {}", c.id, c.weekday);
        }
    }

    // Relay teacher auto-apply
    let today = chrono::Local::now().date_naive();
    let mut updated = false;
    for record in &mut classes {
        if record.level == "初級" && !record.relay_teacher.is_empty() && !record.relay_date.is_empty() {
            if let Some(relay_date) = parse_date(&record.relay_date) {
                if relay_date <= today && record.teacher != record.relay_teacher {
                    record.teacher = record.relay_teacher.clone();
                    updated = true;
                    info!("Relay teacher applied for class {}", record.id);
                }
            }
        }
    }
    if updated {
        save_classes(&classes);
    }

    json!({
        "classes": classes.iter().map(|c| class_payload(c, &holidays, &postpones, &overrides)).collect::<Vec<_>>(),
        "holidays": holidays.iter().map(|h| serde_json::to_value(h).unwrap_or(json!({}))).collect::<Vec<_>>(),
        "postpones": postpones.iter().map(|p| serde_json::to_value(p).unwrap_or(json!({}))).collect::<Vec<_>>(),
        "settings": settings.to_json(),
        "app_config": app_config,
        "stock_history": load_stock_history(),
    })
}

#[tauri::command]
pub fn set_app_location(
    session_token: String,
    location: String,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "state.modify");

    let location_code = location.trim().to_uppercase();
    let allowed = ["", "K", "L", "H"];
    if !allowed.contains(&location_code.as_str()) {
        return json!({"ok": false, "error": "地點不正確。"});
    }
    let mut config = load_app_config();
    config.insert("location".to_string(), location_code.clone());
    save_app_config(&config);
    json!({"ok": true, "location": location_code})
}

#[tauri::command]
pub fn set_tab_order(
    session_token: String,
    order: Vec<String>,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "state.modify");

    let clean: Vec<String> = order.iter().map(|s| s.trim().to_string()).filter(|s| !s.is_empty()).collect();
    let mut config = load_app_config();
    config.insert("tab_order".to_string(), clean.join(","));
    save_app_config(&config);
    json!({"ok": true})
}

#[tauri::command]
pub fn set_eps_output_path(
    session_token: String,
    path: String,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "state.modify");

    let output_path = path.trim().to_string();
    let mut config = load_app_config();
    config.insert("eps_output_path".to_string(), output_path.clone());
    save_app_config(&config);
    json!({"ok": true, "eps_output_path": output_path})
}

#[tauri::command]
pub fn set_zoom_level(
    session_token: String,
    level: f64,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "state.modify");

    let clamped = level.clamp(0.5, 2.0);
    let mut config = load_app_config();
    config.insert("zoom_level".to_string(), format!("{:.2}", clamped));
    save_app_config(&config);
    json!({"ok": true, "zoom_level": clamped})
}

#[tauri::command]
pub fn set_last_review_ts(
    session_token: String,
    ts: String,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "state.modify");

    let mut config = load_app_config();
    config.insert("last_review_ts".to_string(), ts.trim().to_string());
    save_app_config(&config);
    json!({"ok": true})
}
