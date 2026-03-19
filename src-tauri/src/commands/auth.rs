use serde_json::{json, Value};
use tracing::{info, warn};

use crate::auth::db::{self, AuthDb};
use crate::auth::password;
use crate::auth::permissions;
use crate::auth::session::{self, SessionStore};

/// Maximum failed login attempts before account lockout.
const MAX_FAILED_ATTEMPTS: i64 = 5;
/// Lockout duration in minutes.
const LOCKOUT_MINUTES: i64 = 15;

// ---------------------------------------------------------------------------
// check_setup_needed — no auth required
// ---------------------------------------------------------------------------
#[tauri::command]
pub fn check_setup_needed(auth_db: tauri::State<'_, AuthDb>) -> Value {
    let conn = auth_db.lock().unwrap();
    json!({ "ok": true, "setup_needed": db::setup_needed(&conn) })
}

// ---------------------------------------------------------------------------
// setup_admin — only works when zero users exist
// ---------------------------------------------------------------------------
#[tauri::command]
pub fn setup_admin(
    username: String,
    password: String,
    display_name: String,
    auth_db: tauri::State<'_, AuthDb>,
    sessions: tauri::State<'_, SessionStore>,
) -> Value {
    let username = username.trim().to_string();
    let display_name = display_name.trim().to_string();

    if username.is_empty() || password.is_empty() {
        return json!({"ok": false, "error": "Username and password are required."});
    }
    if password.len() != 4 || !password.chars().all(|c| c.is_ascii_digit()) {
        return json!({"ok": false, "error": "Password must be exactly 4 digits."});
    }

    let conn = auth_db.lock().unwrap();

    // Atomically verify no users exist and insert
    let count: i64 = conn
        .query_row("SELECT COUNT(*) FROM users", [], |r| r.get(0))
        .unwrap_or(0);
    if count > 0 {
        return json!({"ok": false, "error": "Setup already completed. Please log in."});
    }

    let hash = match password::hash_password(&password) {
        Ok(h) => h,
        Err(e) => return json!({"ok": false, "error": e}),
    };

    let user_id = uuid::Uuid::new_v4().to_string();
    let now = chrono::Local::now().format("%Y-%m-%d %H:%M:%S").to_string();

    if let Err(e) = conn.execute(
        "INSERT INTO users (id, username, password_hash, role, display_name, created_at, created_by, last_login)
         VALUES (?1, ?2, ?3, 'admin', ?4, ?5, ?1, ?5)",
        rusqlite::params![user_id, username, hash, display_name, now],
    ) {
        return json!({"ok": false, "error": format!("Failed to create admin: {}", e)});
    }

    db::write_audit(&conn, Some(&user_id), Some(&username), "user_created", "Initial admin setup", true);

    // Auto-login
    let mut store = sessions.lock().unwrap();
    let sess = session::create_session(&mut store, &user_id, &username, "admin");
    let perms = permissions::get_role_permissions(&conn, "admin");

    info!("Initial admin '{}' created and logged in", username);

    json!({
        "ok": true,
        "token": sess.token,
        "user": {
            "id": user_id,
            "username": username,
            "role": "admin",
            "display_name": display_name,
        },
        "permissions": perms,
    })
}

// ---------------------------------------------------------------------------
// login
// ---------------------------------------------------------------------------
#[tauri::command]
pub fn login(
    username: String,
    password: String,
    auth_db: tauri::State<'_, AuthDb>,
    sessions: tauri::State<'_, SessionStore>,
) -> Value {
    let username = username.trim().to_string();
    if username.is_empty() || password.is_empty() {
        return json!({"ok": false, "error": "Username and password are required."});
    }

    let conn = auth_db.lock().unwrap();

    // Look up user
    let row = conn.query_row(
        "SELECT id, password_hash, role, display_name, failed_attempts, locked_until, active
         FROM users WHERE username = ?1",
        [&username],
        |r| {
            Ok((
                r.get::<_, String>(0)?,
                r.get::<_, String>(1)?,
                r.get::<_, String>(2)?,
                r.get::<_, String>(3)?,
                r.get::<_, i64>(4)?,
                r.get::<_, Option<String>>(5)?,
                r.get::<_, i64>(6)?,
            ))
        },
    );

    let (user_id, hash, role, display_name, failed_attempts, locked_until, active) = match row {
        Ok(data) => data,
        Err(_) => {
            // Don't reveal whether user exists
            warn!("Login failed: unknown user '{}'", username);
            db::write_audit(&conn, None, Some(&username), "login_failed", "Unknown user", false);
            return json!({"ok": false, "error": "Invalid username or password."});
        }
    };

    // Check if account is deactivated
    if active == 0 {
        db::write_audit(&conn, Some(&user_id), Some(&username), "login_failed", "Account deactivated", false);
        return json!({"ok": false, "error": "Account has been deactivated. Contact an administrator."});
    }

    // Check lockout
    if let Some(ref until) = locked_until {
        if let Ok(lock_time) = chrono::NaiveDateTime::parse_from_str(until, "%Y-%m-%d %H:%M:%S") {
            let now = chrono::Local::now().naive_local();
            if now < lock_time {
                let remaining = (lock_time - now).num_minutes() + 1;
                db::write_audit(&conn, Some(&user_id), Some(&username), "login_failed", "Account locked", false);
                return json!({
                    "ok": false,
                    "error": format!("Account is locked. Try again in {} minute(s).", remaining)
                });
            }
        }
    }

    // Verify password
    let valid = match password::verify_password(&password, &hash) {
        Ok(v) => v,
        Err(e) => {
            warn!("Password verification error: {}", e);
            return json!({"ok": false, "error": "Internal error during authentication."});
        }
    };

    if !valid {
        let new_count = failed_attempts + 1;
        if new_count >= MAX_FAILED_ATTEMPTS {
            let lock_until = chrono::Local::now().naive_local()
                + chrono::Duration::minutes(LOCKOUT_MINUTES);
            let lock_str = lock_until.format("%Y-%m-%d %H:%M:%S").to_string();
            conn.execute(
                "UPDATE users SET failed_attempts = ?1, locked_until = ?2 WHERE id = ?3",
                rusqlite::params![new_count, lock_str, user_id],
            )
            .ok();
            db::write_audit(
                &conn,
                Some(&user_id),
                Some(&username),
                "login_failed",
                &format!("Account locked after {} attempts", new_count),
                false,
            );
            warn!("Account '{}' locked after {} failed attempts", username, new_count);
            return json!({
                "ok": false,
                "error": format!("Too many failed attempts. Account locked for {} minutes.", LOCKOUT_MINUTES)
            });
        } else {
            conn.execute(
                "UPDATE users SET failed_attempts = ?1 WHERE id = ?2",
                rusqlite::params![new_count, user_id],
            )
            .ok();
            db::write_audit(
                &conn,
                Some(&user_id),
                Some(&username),
                "login_failed",
                &format!("Attempt {} of {}", new_count, MAX_FAILED_ATTEMPTS),
                false,
            );
        }
        return json!({"ok": false, "error": "Invalid username or password."});
    }

    // Success — reset failed attempts, update last_login
    let now = chrono::Local::now().format("%Y-%m-%d %H:%M:%S").to_string();
    conn.execute(
        "UPDATE users SET failed_attempts = 0, locked_until = NULL, last_login = ?1 WHERE id = ?2",
        rusqlite::params![now, user_id],
    )
    .ok();

    // Create session
    let mut store = sessions.lock().unwrap();
    session::cleanup_expired(&mut store);
    let sess = session::create_session(&mut store, &user_id, &username, &role);
    let perms = permissions::get_role_permissions(&conn, &role);

    db::write_audit(&conn, Some(&user_id), Some(&username), "login", "", true);
    info!("User '{}' logged in (role: {})", username, role);

    json!({
        "ok": true,
        "token": sess.token,
        "user": {
            "id": user_id,
            "username": username,
            "role": role,
            "display_name": display_name,
        },
        "permissions": perms,
    })
}

// ---------------------------------------------------------------------------
// logout
// ---------------------------------------------------------------------------
#[tauri::command]
pub fn logout(
    session_token: String,
    auth_db: tauri::State<'_, AuthDb>,
    sessions: tauri::State<'_, SessionStore>,
) -> Value {
    let mut store = sessions.lock().unwrap();
    if let Ok(sess) = session::validate_session(&store, &session_token) {
        let conn = auth_db.lock().unwrap();
        db::write_audit(&conn, Some(&sess.user_id), Some(&sess.username), "logout", "", true);
        info!("User '{}' logged out", sess.username);
    }
    session::invalidate_session(&mut store, &session_token);
    json!({"ok": true})
}

// ---------------------------------------------------------------------------
// get_current_user
// ---------------------------------------------------------------------------
#[tauri::command]
pub fn get_current_user(
    session_token: String,
    auth_db: tauri::State<'_, AuthDb>,
    sessions: tauri::State<'_, SessionStore>,
) -> Value {
    let store = sessions.lock().unwrap();
    let sess = match session::validate_session(&store, &session_token) {
        Ok(s) => s,
        Err(e) => return json!({"ok": false, "error": e, "auth_error": true}),
    };
    let conn = auth_db.lock().unwrap();
    let perms = permissions::get_role_permissions(&conn, &sess.role);
    json!({
        "ok": true,
        "user": {
            "id": sess.user_id,
            "username": sess.username,
            "role": sess.role,
        },
        "permissions": perms,
    })
}

// ---------------------------------------------------------------------------
// change_password — user changes their own password
// ---------------------------------------------------------------------------
#[tauri::command]
pub fn change_password(
    session_token: String,
    old_password: String,
    new_password: String,
    auth_db: tauri::State<'_, AuthDb>,
    sessions: tauri::State<'_, SessionStore>,
) -> Value {
    let store = sessions.lock().unwrap();
    let sess = match session::validate_session(&store, &session_token) {
        Ok(s) => s,
        Err(e) => return json!({"ok": false, "error": e, "auth_error": true}),
    };
    drop(store);

    if new_password.len() != 4 || !new_password.chars().all(|c| c.is_ascii_digit()) {
        return json!({"ok": false, "error": "New password must be exactly 4 digits."});
    }

    let conn = auth_db.lock().unwrap();

    // Verify old password
    let hash: String = match conn.query_row(
        "SELECT password_hash FROM users WHERE id = ?1",
        [&sess.user_id],
        |r| r.get(0),
    ) {
        Ok(h) => h,
        Err(_) => return json!({"ok": false, "error": "User not found."}),
    };

    match password::verify_password(&old_password, &hash) {
        Ok(true) => {}
        Ok(false) => return json!({"ok": false, "error": "Current password is incorrect."}),
        Err(e) => return json!({"ok": false, "error": e}),
    }

    // Hash and store new password
    let new_hash = match password::hash_password(&new_password) {
        Ok(h) => h,
        Err(e) => return json!({"ok": false, "error": e}),
    };

    conn.execute(
        "UPDATE users SET password_hash = ?1 WHERE id = ?2",
        rusqlite::params![new_hash, sess.user_id],
    )
    .ok();

    db::write_audit(
        &conn,
        Some(&sess.user_id),
        Some(&sess.username),
        "password_changed",
        "Self-service password change",
        true,
    );

    info!("User '{}' changed their password", sess.username);
    json!({"ok": true})
}
