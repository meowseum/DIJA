use serde_json::{json, Value};
use tracing::info;

use crate::auth::db::{self, AuthDb};
use crate::auth::password;
use crate::auth::permissions::{self, ALL_PERMISSIONS};
use crate::auth::session::SessionStore;

// ---------------------------------------------------------------------------
// list_users
// ---------------------------------------------------------------------------
#[tauri::command]
pub fn list_users(
    session_token: String,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "admin.users");

    let conn = auth_db.lock().unwrap();
    let mut stmt = conn
        .prepare(
            "SELECT id, username, role, display_name, created_at, last_login, active
             FROM users ORDER BY created_at ASC",
        )
        .unwrap();

    let users: Vec<Value> = stmt
        .query_map([], |r| {
            Ok(json!({
                "id": r.get::<_, String>(0)?,
                "username": r.get::<_, String>(1)?,
                "role": r.get::<_, String>(2)?,
                "display_name": r.get::<_, String>(3)?,
                "created_at": r.get::<_, String>(4)?,
                "last_login": r.get::<_, Option<String>>(5)?,
                "active": r.get::<_, i64>(6)? == 1,
            }))
        })
        .unwrap()
        .filter_map(|r| r.ok())
        .collect();

    json!({"ok": true, "users": users})
}

// ---------------------------------------------------------------------------
// create_user
// ---------------------------------------------------------------------------
#[tauri::command]
pub fn create_user(
    session_token: String,
    username: String,
    password_val: String,
    role: String,
    display_name: String,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let session = crate::require_auth!(sessions, auth_db, &session_token, "admin.users");

    let username = username.trim().to_string();
    let display_name = display_name.trim().to_string();
    let role = role.trim().to_lowercase();

    if username.is_empty() || password_val.is_empty() {
        return json!({"ok": false, "error": "Username and password are required."});
    }
    if password_val.len() != 4 || !password_val.chars().all(|c| c.is_ascii_digit()) {
        return json!({"ok": false, "error": "Password must be exactly 4 digits."});
    }
    if role == "dev" {
        return json!({"ok": false, "error": "The dev role cannot be assigned manually."});
    }

    let conn = auth_db.lock().unwrap();

    // Check username uniqueness
    let exists: i64 = conn
        .query_row(
            "SELECT COUNT(*) FROM users WHERE username = ?1",
            [&username],
            |r| r.get(0),
        )
        .unwrap_or(0);
    if exists > 0 {
        return json!({"ok": false, "error": "Username already exists."});
    }

    // Validate that the role has permissions defined (or is 'admin'/'staff')
    let role_exists: i64 = conn
        .query_row(
            "SELECT COUNT(*) FROM role_permissions WHERE role = ?1",
            [&role],
            |r| r.get(0),
        )
        .unwrap_or(0);
    if role_exists == 0 && role != "admin" {
        return json!({"ok": false, "error": format!("Role '{}' does not exist. Create role permissions first.", role)});
    }

    let hash = match password::hash_password(&password_val) {
        Ok(h) => h,
        Err(e) => return json!({"ok": false, "error": e}),
    };

    let user_id = uuid::Uuid::new_v4().to_string();
    let now = chrono::Local::now().format("%Y-%m-%d %H:%M:%S").to_string();

    if let Err(e) = conn.execute(
        "INSERT INTO users (id, username, password_hash, role, display_name, created_at, created_by)
         VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7)",
        rusqlite::params![user_id, username, hash, role, display_name, now, session.user_id],
    ) {
        return json!({"ok": false, "error": format!("Failed to create user: {}", e)});
    }

    db::write_audit(
        &conn,
        Some(&session.user_id),
        Some(&session.username),
        "user_created",
        &format!("Created user '{}' with role '{}'", username, role),
        true,
    );

    info!(
        "Admin '{}' created user '{}' (role: {})",
        session.username, username, role
    );

    json!({"ok": true, "user_id": user_id})
}

// ---------------------------------------------------------------------------
// update_user — change role, display_name, or active status
// ---------------------------------------------------------------------------
#[tauri::command]
pub fn update_user(
    session_token: String,
    user_id: String,
    role: Option<String>,
    display_name: Option<String>,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let session = crate::require_auth!(sessions, auth_db, &session_token, "admin.users");

    let conn = auth_db.lock().unwrap();

    // Get current user data
    let current_role: String = match conn.query_row(
        "SELECT role FROM users WHERE id = ?1",
        [&user_id],
        |r| r.get(0),
    ) {
        Ok(r) => r,
        Err(_) => return json!({"ok": false, "error": "User not found."}),
    };

    // Dev accounts cannot be modified by anyone except themselves
    if current_role == "dev" && session.role != "dev" {
        return json!({"ok": false, "error": "Cannot modify a dev account."});
    }

    let mut details = Vec::new();

    if let Some(new_role) = &role {
        let new_role = new_role.trim().to_lowercase();
        if new_role == "dev" {
            return json!({"ok": false, "error": "The dev role cannot be assigned manually."});
        }
        if new_role != current_role {
            // Validate role exists
            let role_exists: i64 = conn
                .query_row(
                    "SELECT COUNT(*) FROM role_permissions WHERE role = ?1",
                    [&new_role],
                    |r| r.get(0),
                )
                .unwrap_or(0);
            if role_exists == 0 && new_role != "admin" {
                return json!({"ok": false, "error": format!("Role '{}' does not exist.", new_role)});
            }

            // Prevent removing the last admin
            if current_role == "admin" && new_role != "admin" {
                let admin_count: i64 = conn
                    .query_row(
                        "SELECT COUNT(*) FROM users WHERE role = 'admin' AND active = 1",
                        [],
                        |r| r.get(0),
                    )
                    .unwrap_or(0);
                if admin_count <= 1 {
                    return json!({"ok": false, "error": "Cannot change role of the last admin."});
                }
            }

            conn.execute(
                "UPDATE users SET role = ?1 WHERE id = ?2",
                rusqlite::params![new_role, user_id],
            )
            .ok();
            details.push(format!("role: {} → {}", current_role, new_role));
        }
    }

    if let Some(dn) = &display_name {
        conn.execute(
            "UPDATE users SET display_name = ?1 WHERE id = ?2",
            rusqlite::params![dn.trim(), user_id],
        )
        .ok();
        details.push(format!("display_name updated"));
    }

    if !details.is_empty() {
        db::write_audit(
            &conn,
            Some(&session.user_id),
            Some(&session.username),
            "user_updated",
            &format!("Updated user {}: {}", user_id, details.join(", ")),
            true,
        );
    }

    json!({"ok": true})
}

// ---------------------------------------------------------------------------
// deactivate_user — soft delete
// ---------------------------------------------------------------------------
#[tauri::command]
pub fn deactivate_user(
    session_token: String,
    user_id: String,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let session = crate::require_auth!(sessions, auth_db, &session_token, "admin.users");

    // Cannot deactivate yourself
    if session.user_id == user_id {
        return json!({"ok": false, "error": "Cannot deactivate your own account."});
    }

    let conn = auth_db.lock().unwrap();

    // Check if target is the last admin
    let target_role: String = match conn.query_row(
        "SELECT role FROM users WHERE id = ?1",
        [&user_id],
        |r| r.get(0),
    ) {
        Ok(r) => r,
        Err(_) => return json!({"ok": false, "error": "User not found."}),
    };

    if target_role == "dev" {
        return json!({"ok": false, "error": "Cannot deactivate a dev account."});
    }

    if target_role == "admin" {
        let admin_count: i64 = conn
            .query_row(
                "SELECT COUNT(*) FROM users WHERE role = 'admin' AND active = 1",
                [],
                |r| r.get(0),
            )
            .unwrap_or(0);
        if admin_count <= 1 {
            return json!({"ok": false, "error": "Cannot deactivate the last admin account."});
        }
    }

    conn.execute(
        "UPDATE users SET active = 0 WHERE id = ?1",
        [&user_id],
    )
    .ok();

    // Invalidate their sessions
    let mut store = sessions.lock().unwrap();
    store.retain(|_, s| s.user_id != user_id);

    let target_name: String = conn
        .query_row("SELECT username FROM users WHERE id = ?1", [&user_id], |r| {
            r.get(0)
        })
        .unwrap_or_default();

    db::write_audit(
        &conn,
        Some(&session.user_id),
        Some(&session.username),
        "user_deactivated",
        &format!("Deactivated user '{}' ({})", target_name, user_id),
        true,
    );

    info!("Admin '{}' deactivated user '{}'", session.username, target_name);
    json!({"ok": true})
}

// ---------------------------------------------------------------------------
// reactivate_user
// ---------------------------------------------------------------------------
#[tauri::command]
pub fn reactivate_user(
    session_token: String,
    user_id: String,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let session = crate::require_auth!(sessions, auth_db, &session_token, "admin.users");

    let conn = auth_db.lock().unwrap();

    conn.execute(
        "UPDATE users SET active = 1 WHERE id = ?1",
        [&user_id],
    )
    .ok();

    let target_name: String = conn
        .query_row("SELECT username FROM users WHERE id = ?1", [&user_id], |r| {
            r.get(0)
        })
        .unwrap_or_default();

    db::write_audit(
        &conn,
        Some(&session.user_id),
        Some(&session.username),
        "user_reactivated",
        &format!("Reactivated user '{}' ({})", target_name, user_id),
        true,
    );

    info!("Admin '{}' reactivated user '{}'", session.username, target_name);
    json!({"ok": true})
}

// ---------------------------------------------------------------------------
// reset_password — admin resets another user's password
// ---------------------------------------------------------------------------
#[tauri::command]
pub fn reset_password(
    session_token: String,
    user_id: String,
    new_password: String,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let session = crate::require_auth!(sessions, auth_db, &session_token, "admin.users");

    if new_password.len() != 4 || !new_password.chars().all(|c| c.is_ascii_digit()) {
        return json!({"ok": false, "error": "Password must be exactly 4 digits."});
    }

    let conn = auth_db.lock().unwrap();

    // Block non-dev users from resetting a dev account's password
    let target_role: String = conn
        .query_row("SELECT role FROM users WHERE id = ?1", [&user_id], |r| r.get(0))
        .unwrap_or_default();
    if target_role == "dev" && session.role != "dev" {
        return json!({"ok": false, "error": "Cannot reset a dev account's password."});
    }

    let hash = match password::hash_password(&new_password) {
        Ok(h) => h,
        Err(e) => return json!({"ok": false, "error": e}),
    };

    let affected = conn
        .execute(
            "UPDATE users SET password_hash = ?1, failed_attempts = 0, locked_until = NULL WHERE id = ?2",
            rusqlite::params![hash, user_id],
        )
        .unwrap_or(0);

    if affected == 0 {
        return json!({"ok": false, "error": "User not found."});
    }

    let target_name: String = conn
        .query_row("SELECT username FROM users WHERE id = ?1", [&user_id], |r| {
            r.get(0)
        })
        .unwrap_or_default();

    db::write_audit(
        &conn,
        Some(&session.user_id),
        Some(&session.username),
        "password_reset",
        &format!("Reset password for user '{}' ({})", target_name, user_id),
        true,
    );

    info!("Admin '{}' reset password for '{}'", session.username, target_name);
    json!({"ok": true})
}

// ---------------------------------------------------------------------------
// list_all_permissions — returns the master permission list for UI
// ---------------------------------------------------------------------------
#[tauri::command]
pub fn list_all_permissions(
    session_token: String,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "admin.roles");

    let perms: Vec<Value> = ALL_PERMISSIONS
        .iter()
        .map(|(key, name, category)| {
            json!({
                "key": key,
                "name": name,
                "category": category,
            })
        })
        .collect();

    json!({"ok": true, "permissions": perms})
}

// ---------------------------------------------------------------------------
// list_role_permissions
// ---------------------------------------------------------------------------
#[tauri::command]
pub fn list_role_permissions(
    session_token: String,
    role: String,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "admin.roles");

    let conn = auth_db.lock().unwrap();
    let perms = permissions::get_role_permissions(&conn, &role);

    json!({"ok": true, "role": role, "permissions": perms})
}

// ---------------------------------------------------------------------------
// set_role_permissions — replace all permissions for a role
// ---------------------------------------------------------------------------
#[tauri::command]
pub fn set_role_permissions(
    session_token: String,
    role: String,
    permission_list: Vec<String>,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let session = crate::require_auth!(sessions, auth_db, &session_token, "admin.roles");

    let role = role.trim().to_lowercase();

    // Cannot modify the admin or dev roles
    if role == "admin" || role == "dev" {
        return json!({"ok": false, "error": format!("The {} role cannot be modified.", role)});
    }

    if role.is_empty() {
        return json!({"ok": false, "error": "Role name is required."});
    }

    // Validate all permission keys
    let valid_keys: std::collections::HashSet<&str> =
        ALL_PERMISSIONS.iter().map(|(k, _, _)| *k).collect();
    for p in &permission_list {
        if !valid_keys.contains(p.as_str()) {
            return json!({"ok": false, "error": format!("Unknown permission: '{}'", p)});
        }
    }

    let conn = auth_db.lock().unwrap();

    // Replace in a transaction
    let tx = conn.unchecked_transaction().unwrap();
    tx.execute("DELETE FROM role_permissions WHERE role = ?1", [&role])
        .ok();
    for p in &permission_list {
        tx.execute(
            "INSERT INTO role_permissions (role, permission) VALUES (?1, ?2)",
            rusqlite::params![role, p],
        )
        .ok();
    }
    if let Err(e) = tx.commit() {
        return json!({"ok": false, "error": format!("Failed to update permissions: {}", e)});
    }

    db::write_audit(
        &conn,
        Some(&session.user_id),
        Some(&session.username),
        "permissions_changed",
        &format!("Set {} permissions for role '{}'", permission_list.len(), role),
        true,
    );

    info!(
        "Admin '{}' updated permissions for role '{}' ({} permissions)",
        session.username,
        role,
        permission_list.len()
    );

    json!({"ok": true})
}

// ---------------------------------------------------------------------------
// list_roles — returns all defined roles
// ---------------------------------------------------------------------------
#[tauri::command]
pub fn list_roles(
    session_token: String,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "admin.roles");

    let conn = auth_db.lock().unwrap();
    let mut stmt = conn
        .prepare("SELECT DISTINCT role FROM role_permissions ORDER BY role")
        .unwrap();
    let roles: Vec<String> = stmt
        .query_map([], |r| r.get::<_, String>(0))
        .unwrap()
        .filter_map(|r| r.ok())
        .collect();

    // Ensure 'admin' is always in the list
    let mut all_roles = roles;
    if !all_roles.contains(&"admin".to_string()) {
        all_roles.insert(0, "admin".to_string());
    }

    json!({"ok": true, "roles": all_roles})
}

// ---------------------------------------------------------------------------
// get_audit_log
// ---------------------------------------------------------------------------
#[tauri::command]
pub fn get_audit_log(
    session_token: String,
    limit: Option<i64>,
    offset: Option<i64>,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "admin.audit");

    let limit = limit.unwrap_or(100).min(500);
    let offset = offset.unwrap_or(0);

    let conn = auth_db.lock().unwrap();

    let total: i64 = conn
        .query_row("SELECT COUNT(*) FROM audit_log", [], |r| r.get(0))
        .unwrap_or(0);

    let mut stmt = conn
        .prepare(
            "SELECT id, timestamp, user_id, username, event_type, details, success
             FROM audit_log ORDER BY id DESC LIMIT ?1 OFFSET ?2",
        )
        .unwrap();

    let entries: Vec<Value> = stmt
        .query_map(rusqlite::params![limit, offset], |r| {
            Ok(json!({
                "id": r.get::<_, i64>(0)?,
                "timestamp": r.get::<_, String>(1)?,
                "user_id": r.get::<_, String>(2)?,
                "username": r.get::<_, String>(3)?,
                "event_type": r.get::<_, String>(4)?,
                "details": r.get::<_, String>(5)?,
                "success": r.get::<_, i64>(6)? == 1,
            }))
        })
        .unwrap()
        .filter_map(|r| r.ok())
        .collect();

    json!({"ok": true, "entries": entries, "total": total})
}
