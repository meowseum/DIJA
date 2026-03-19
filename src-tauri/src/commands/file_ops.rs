use serde_json::{json, Value};

use crate::auth::db::AuthDb;
use crate::auth::session::SessionStore;
use crate::config::get_output_dir;

#[tauri::command]
pub fn open_output_folder(
    session_token: String,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "documents.view");

    let output_dir = get_output_dir();
    match opener::open(&output_dir) {
        Ok(()) => json!({"ok": true}),
        Err(e) => json!({"ok": false, "error": format!("{}", e)}),
    }
}
