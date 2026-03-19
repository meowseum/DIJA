use serde_json::{json, Value};
use tracing::info;

use crate::auth::db::AuthDb;
use crate::auth::session::SessionStore;
use crate::models::parse_date;
use crate::storage::*;
use super::classes::build_schedule_with_index_pub;

#[tauri::command]
pub fn get_calendar_data(
    session_token: String,
    start_date: String,
    end_date: String,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "calendar.view");

    info!("get_calendar_data called: {} to {}", start_date, end_date);
    let start = match parse_date(&start_date) {
        Some(d) => d,
        None => return json!({"ok": false, "error": "日期不正確。"}),
    };
    let end = match parse_date(&end_date) {
        Some(d) => d,
        None => return json!({"ok": false, "error": "日期不正確。"}),
    };
    let (start, end) = if end < start { (end, start) } else { (start, end) };

    let classes = load_classes();
    let holidays = load_holidays();
    let postpones = load_postpones();
    let overrides = load_overrides();
    let mut calendar_sessions: Vec<Value> = Vec::new();

    for class_record in &classes {
        let class_postpones: Vec<_> = postpones.iter().filter(|p| p.class_id == class_record.id).cloned().collect();
        let class_overrides: Vec<_> = overrides.iter().filter(|o| o.class_id == class_record.id).cloned().collect();
        let schedule = build_schedule_with_index_pub(class_record, &holidays, &class_postpones, &class_overrides);

        for item in &schedule {
            let date_str = item.get("date").and_then(|v| v.as_str()).unwrap_or("");
            let lesson_date = match parse_date(date_str) {
                Some(d) => d,
                None => continue,
            };
            if lesson_date < start || lesson_date > end {
                continue;
            }
            let idx = item.get("index").and_then(|v| v.as_i64()).unwrap_or(0);
            let payment_due = idx % 4 == 3 || idx % 4 == 0;

            calendar_sessions.push(json!({
                "date": date_str,
                "sku": class_record.sku,
                "class_id": class_record.id,
                "location": class_record.location,
                "room": class_record.classroom,
                "teacher": class_record.teacher,
                "time": class_record.start_time,
                "lesson_index": idx,
                "lesson_total": class_record.lesson_total,
                "payment_due": payment_due,
            }));
        }
    }

    json!({
        "ok": true,
        "sessions": calendar_sessions,
        "holidays": holidays.iter().map(|h| serde_json::to_value(h).unwrap_or(json!({}))).collect::<Vec<_>>(),
    })
}
