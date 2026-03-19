use serde_json::{json, Value};
use uuid::Uuid;

use crate::auth::db::AuthDb;
use crate::auth::session::SessionStore;
use crate::models::HolidayRange;
use crate::storage::*;

#[tauri::command]
pub fn add_holiday(
    session_token: String,
    data: Value,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "holidays.add");

    let holiday = HolidayRange {
        id: Uuid::new_v4().to_string(),
        start_date: data.get("start_date").and_then(|v| v.as_str()).unwrap_or("").trim().to_string(),
        end_date: data.get("end_date").and_then(|v| v.as_str()).unwrap_or("").trim().to_string(),
        name: data.get("name").and_then(|v| v.as_str()).unwrap_or("").trim().to_string(),
    };
    let mut holidays = load_holidays();
    holidays.push(holiday);
    save_holidays(&holidays);
    json!({"ok": true})
}

#[tauri::command]
pub fn delete_holiday(
    session_token: String,
    holiday_id: String,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "holidays.delete");

    let holidays = load_holidays();
    let new_holidays: Vec<_> = holidays.into_iter().filter(|h| h.id != holiday_id).collect();
    save_holidays(&new_holidays);
    json!({"ok": true})
}
