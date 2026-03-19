use serde_json::{json, Value};
use uuid::Uuid;

use crate::auth::db::AuthDb;
use crate::auth::session::SessionStore;
use crate::models::{parse_date, LessonOverride};
use crate::storage::*;

#[tauri::command]
pub fn add_schedule_override(
    session_token: String,
    class_id: String,
    date_str: String,
    action: String,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "overrides.add");

    if parse_date(&date_str).is_none() {
        return json!({"ok": false, "error": "日期不正確。"});
    }
    if action != "add" && action != "remove" {
        return json!({"ok": false, "error": "操作不正確。"});
    }
    let mut overrides = load_overrides();
    let exists = overrides.iter().any(|o| o.class_id == class_id && o.date == date_str && o.action == action);
    if exists {
        return json!({"ok": true});
    }
    overrides.push(LessonOverride {
        id: Uuid::new_v4().to_string(),
        class_id,
        date: date_str,
        action,
    });
    save_overrides(&overrides);
    json!({"ok": true})
}

#[tauri::command]
pub fn delete_schedule_override(
    session_token: String,
    override_id: String,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "overrides.delete");

    let overrides = load_overrides();
    let new_overrides: Vec<_> = overrides.into_iter().filter(|o| o.id != override_id).collect();
    save_overrides(&new_overrides);
    json!({"ok": true})
}
