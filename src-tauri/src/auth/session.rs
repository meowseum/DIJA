use chrono::NaiveDateTime;
use rand::RngCore;
use std::collections::HashMap;
use std::sync::Mutex;

/// A live user session.
#[derive(Clone, Debug)]
pub struct Session {
    pub token: String,
    pub user_id: String,
    pub username: String,
    pub role: String,
    pub created_at: NaiveDateTime,
    pub expires_at: NaiveDateTime,
}

/// Type alias used with `tauri::State`.
pub type SessionStore = Mutex<HashMap<String, Session>>;

/// Default session lifetime: 8 hours.
const SESSION_LIFETIME_SECS: i64 = 8 * 3600;

/// Generate a cryptographically secure 256-bit hex token.
pub fn generate_token() -> String {
    let mut bytes = [0u8; 32];
    rand::rngs::OsRng.fill_bytes(&mut bytes);
    hex::encode(&bytes)
}

/// We encode hex manually to avoid adding a dep — tiny helper.
mod hex {
    pub fn encode(bytes: &[u8]) -> String {
        bytes.iter().map(|b| format!("{:02x}", b)).collect()
    }
}

/// Create a new session for the given user, invalidating any previous session for that user.
pub fn create_session(
    store: &mut HashMap<String, Session>,
    user_id: &str,
    username: &str,
    role: &str,
) -> Session {
    // Invalidate any existing session for this user (one session per user)
    store.retain(|_, s| s.user_id != user_id);

    let now = chrono::Local::now().naive_local();
    let session = Session {
        token: generate_token(),
        user_id: user_id.to_string(),
        username: username.to_string(),
        role: role.to_string(),
        created_at: now,
        expires_at: now + chrono::Duration::seconds(SESSION_LIFETIME_SECS),
    };
    store.insert(session.token.clone(), session.clone());
    session
}

/// Validate a session token. Returns the session if valid and not expired.
pub fn validate_session(
    store: &HashMap<String, Session>,
    token: &str,
) -> Result<Session, String> {
    let session = store
        .get(token)
        .ok_or_else(|| "Invalid session token".to_string())?;
    let now = chrono::Local::now().naive_local();
    if now > session.expires_at {
        return Err("Session expired".to_string());
    }
    Ok(session.clone())
}

/// Remove a session (logout).
pub fn invalidate_session(store: &mut HashMap<String, Session>, token: &str) {
    store.remove(token);
}

/// Remove all expired sessions (housekeeping).
pub fn cleanup_expired(store: &mut HashMap<String, Session>) {
    let now = chrono::Local::now().naive_local();
    store.retain(|_, s| s.expires_at > now);
}
