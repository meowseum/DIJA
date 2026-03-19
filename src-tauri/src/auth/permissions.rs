use rusqlite::Connection;
use std::collections::HashMap;

use super::session::Session;

/// Master list of all permissions: (key, display_name, category).
pub const ALL_PERMISSIONS: &[(&str, &str, &str)] = &[
    // Classes
    ("classes.view", "View Classes", "Classes"),
    ("classes.create", "Create Class", "Classes"),
    ("classes.update", "Update Class", "Classes"),
    ("classes.delete", "Delete Class", "Classes"),
    ("classes.end", "End / Promote Class", "Classes"),
    // Holidays
    ("holidays.add", "Add Holiday", "Holidays"),
    ("holidays.delete", "Delete Holiday", "Holidays"),
    // Postpones
    ("postpones.add", "Add Postpone", "Postpones"),
    ("postpones.delete", "Delete Postpone", "Postpones"),
    ("postpones.view", "View Make-up Date", "Postpones"),
    // Overrides
    ("overrides.add", "Add Override", "Overrides"),
    ("overrides.delete", "Delete Override", "Overrides"),
    // Settings
    ("settings.modify", "Modify Settings", "Settings"),
    ("settings.export", "Export Settings", "Settings"),
    ("settings.import", "Import Settings", "Settings"),
    // Textbooks
    ("textbooks.modify", "Modify Textbooks", "Textbooks"),
    ("textbooks.view", "View Stock History", "Textbooks"),
    // Documents
    ("documents.generate", "Generate Documents", "Documents"),
    ("documents.view", "View Templates", "Documents"),
    // Calendar
    ("calendar.view", "View Calendar", "Calendar"),
    // Export/Import
    ("export.classes", "Export Classes", "Export/Import"),
    ("import.classes", "Import Classes", "Export/Import"),
    // EPS
    ("eps.view", "View EPS Records", "EPS"),
    ("eps.modify", "Modify EPS Records", "EPS"),
    ("eps.export", "Export EPS", "EPS"),
    // System
    ("state.view", "View App State", "System"),
    ("state.modify", "Modify App Config", "System"),
    // Admin
    ("admin.users", "Manage Users", "Admin"),
    ("admin.roles", "Manage Roles", "Admin"),
    ("admin.audit", "View Audit Logs", "Admin"),
];

/// Check whether a role has a specific permission.
pub fn role_has_permission(conn: &Connection, role: &str, permission: &str) -> bool {
    // Admin always has all permissions (hardcoded safety net)
    if role == "admin" {
        return true;
    }
    let count: i64 = conn
        .query_row(
            "SELECT COUNT(*) FROM role_permissions WHERE role = ?1 AND permission = ?2",
            rusqlite::params![role, permission],
            |r| r.get(0),
        )
        .unwrap_or(0);
    count > 0
}

/// Get all permissions for a role.
pub fn get_role_permissions(conn: &Connection, role: &str) -> Vec<String> {
    // Admin always has everything
    if role == "admin" {
        return ALL_PERMISSIONS.iter().map(|(k, _, _)| k.to_string()).collect();
    }
    let mut stmt = conn
        .prepare("SELECT permission FROM role_permissions WHERE role = ?1")
        .unwrap();
    stmt.query_map([role], |row| row.get::<_, String>(0))
        .unwrap()
        .filter_map(|r| r.ok())
        .collect()
}

/// Authorise a session token against a required permission.
/// Returns the session on success, or an error string on failure.
pub fn authorize(
    sessions: &HashMap<String, Session>,
    conn: &Connection,
    token: &str,
    required_permission: &str,
) -> Result<Session, String> {
    let session = super::session::validate_session(sessions, token)?;

    if !role_has_permission(conn, &session.role, required_permission) {
        return Err(format!(
            "Permission denied: '{}' requires '{}'",
            session.username, required_permission
        ));
    }

    Ok(session)
}
