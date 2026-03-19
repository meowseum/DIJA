use rusqlite::Connection;
use std::sync::Mutex;
use tracing::info;

use crate::config::get_data_dir;

/// Type alias used with `tauri::State`.
pub type AuthDb = Mutex<Connection>;

/// Initialise the auth database. Creates tables if they don't exist.
pub fn init_database() -> Connection {
    let db_path = get_data_dir().join("auth.db");
    info!("Auth database path: {:?}", db_path);

    let conn =
        Connection::open(&db_path).expect("Failed to open auth database");

    // Enable WAL mode for crash safety
    conn.execute_batch("PRAGMA journal_mode=WAL;")
        .expect("Failed to set WAL mode");

    conn.execute_batch(
        "
        CREATE TABLE IF NOT EXISTS users (
            id              TEXT PRIMARY KEY,
            username        TEXT UNIQUE NOT NULL,
            password_hash   TEXT NOT NULL,
            role            TEXT NOT NULL DEFAULT 'staff',
            display_name    TEXT NOT NULL DEFAULT '',
            created_at      TEXT NOT NULL,
            created_by      TEXT NOT NULL,
            last_login      TEXT,
            failed_attempts INTEGER DEFAULT 0,
            locked_until    TEXT,
            active          INTEGER DEFAULT 1
        );

        CREATE TABLE IF NOT EXISTS role_permissions (
            role       TEXT NOT NULL,
            permission TEXT NOT NULL,
            PRIMARY KEY (role, permission)
        );

        CREATE TABLE IF NOT EXISTS audit_log (
            id         INTEGER PRIMARY KEY AUTOINCREMENT,
            timestamp  TEXT NOT NULL,
            user_id    TEXT,
            username   TEXT,
            event_type TEXT NOT NULL,
            details    TEXT DEFAULT '',
            success    INTEGER DEFAULT 1
        );
        ",
    )
    .expect("Failed to create auth tables");

    // Seed default role permissions if the table is empty
    let count: i64 = conn
        .query_row("SELECT COUNT(*) FROM role_permissions", [], |r| r.get(0))
        .unwrap_or(0);
    if count == 0 {
        seed_default_permissions(&conn);
    }

    conn
}

/// Check whether zero users exist (first-run setup needed).
pub fn setup_needed(conn: &Connection) -> bool {
    let count: i64 = conn
        .query_row("SELECT COUNT(*) FROM users WHERE active = 1", [], |r| {
            r.get(0)
        })
        .unwrap_or(0);
    count == 0
}

/// Write an entry to the audit log.
pub fn write_audit(
    conn: &Connection,
    user_id: Option<&str>,
    username: Option<&str>,
    event_type: &str,
    details: &str,
    success: bool,
) {
    let now = chrono::Local::now().format("%Y-%m-%d %H:%M:%S").to_string();
    conn.execute(
        "INSERT INTO audit_log (timestamp, user_id, username, event_type, details, success)
         VALUES (?1, ?2, ?3, ?4, ?5, ?6)",
        rusqlite::params![
            now,
            user_id.unwrap_or(""),
            username.unwrap_or(""),
            event_type,
            details,
            success as i32
        ],
    )
    .ok(); // audit should never crash the app
}

use super::permissions::ALL_PERMISSIONS;

/// Seed the default admin and staff role permissions.
fn seed_default_permissions(conn: &Connection) {
    let tx = conn.unchecked_transaction().expect("Failed to start tx");

    // Admin gets everything
    for (key, _, _) in ALL_PERMISSIONS {
        tx.execute(
            "INSERT OR IGNORE INTO role_permissions (role, permission) VALUES ('admin', ?1)",
            [key],
        )
        .ok();
    }

    // Staff gets everything except admin.* permissions
    for (key, _, _) in ALL_PERMISSIONS {
        if !key.starts_with("admin.") {
            tx.execute(
                "INSERT OR IGNORE INTO role_permissions (role, permission) VALUES ('staff', ?1)",
                [key],
            )
            .ok();
        }
    }

    tx.commit().expect("Failed to commit default permissions");
    info!("Seeded default role permissions (admin, staff)");
}
