use serde_json::{json, Value};
use tracing::info;

use crate::auth::db::AuthDb;
use crate::auth::session::SessionStore;
use crate::config::data_file;
use crate::models::ClassRecord;
use crate::storage::*;

#[tauri::command]
pub fn export_classes_csv(
    session_token: String,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "export.classes");

    let classes = load_classes();
    let mut wtr = csv::WriterBuilder::new().from_writer(Vec::new());
    let _ = wtr.write_record(CLASS_HEADERS);
    for record in &classes {
        let _ = wtr.serialize(record);
    }
    let _ = wtr.flush();
    let content = String::from_utf8(wtr.into_inner().unwrap_or_default()).unwrap_or_default();
    json!({"ok": true, "content": content})
}

#[tauri::command]
pub fn import_classes_csv(
    session_token: String,
    content: String,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "import.classes");

    let mut rdr = csv::Reader::from_reader(content.as_bytes());
    let mut classes: Vec<ClassRecord> = Vec::new();
    for result in rdr.deserialize() {
        match result {
            Ok(record) => classes.push(record),
            Err(_) => return json!({"ok": false, "error": "CSV 內容不正確。"}),
        }
    }
    backup_file(&data_file("classes.csv"));
    save_classes(&classes);
    info!("Classes CSV imported: {} records", classes.len());
    json!({"ok": true})
}
