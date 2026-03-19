pub mod db;
pub mod password;
pub mod permissions;
pub mod session;

/// Macro to authorize a command invocation. Call at the top of every Tauri command.
///
/// Usage:
/// ```ignore
/// require_auth!(sessions, auth_db, &session_token, "classes.create");
/// ```
///
/// On failure, returns `{ "ok": false, "error": "...", "auth_error": true }`.
#[macro_export]
macro_rules! require_auth {
    ($sessions:expr, $auth_db:expr, $token:expr, $perm:expr) => {{
        let sessions_guard = $sessions.lock().unwrap();
        let db_guard = $auth_db.lock().unwrap();
        match crate::auth::permissions::authorize(&sessions_guard, &db_guard, $token, $perm) {
            Ok(session) => {
                drop(db_guard);
                drop(sessions_guard);
                session
            }
            Err(e) => {
                return serde_json::json!({"ok": false, "error": e, "auth_error": true});
            }
        }
    }};
}
