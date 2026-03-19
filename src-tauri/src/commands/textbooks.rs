use regex::Regex;
use serde_json::{json, Value};
use std::collections::HashMap;

use crate::auth::db::AuthDb;
use crate::auth::session::SessionStore;
use crate::storage::*;

#[tauri::command]
pub fn set_textbook(
    session_token: String,
    name: String,
    price: f64,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "textbooks.modify");

    let name = name.trim().to_string();
    if name.is_empty() {
        return json!({"ok": false, "error": "教材名稱不能為空。"});
    }
    let price_value = (price as i64).max(0);
    let mut settings = load_settings();
    settings.textbook.insert(name, price_value);
    save_settings(&settings);
    json!({"ok": true})
}

#[tauri::command]
pub fn delete_textbook(
    session_token: String,
    name: String,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "textbooks.modify");

    let name = name.trim().to_string();
    let mut settings = load_settings();
    let mut changed = false;

    if settings.textbook.remove(&name).is_some() {
        changed = true;
    }
    if settings.textbook_stock.remove(&name).is_some() {
        changed = true;
    }

    let mut lt_changed = false;
    let keys: Vec<String> = settings.level_textbook.keys().cloned().collect();
    for lv in keys {
        if let Some(v) = settings.level_textbook.get_mut(&lv) {
            if v.contains(&name) {
                v.retain(|n| n != &name);
                lt_changed = true;
                if v.is_empty() {
                    settings.level_textbook.remove(&lv);
                }
            }
        }
    }

    if changed || lt_changed {
        save_settings(&settings);
    }
    json!({"ok": true})
}

#[tauri::command]
pub fn set_textbook_stock(
    session_token: String,
    name: String,
    count: f64,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "textbooks.modify");

    let name = name.trim().to_string();
    if name.is_empty() {
        return json!({"ok": false, "error": "教材名稱不能為空。"});
    }
    let count_value = (count as i64).max(0);
    let mut settings = load_settings();
    settings.textbook_stock.insert(name, count_value);
    save_settings(&settings);
    json!({"ok": true})
}

#[tauri::command]
pub fn save_monthly_stock(
    session_token: String,
    month: String,
    stock_data: HashMap<String, Value>,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "textbooks.modify");

    let month = month.trim().to_string();
    let re = Regex::new(r"^\d{4}-\d{2}$").unwrap();
    if !re.is_match(&month) {
        return json!({"ok": false, "error": "月份格式不正確（應為 YYYY-MM）。"});
    }
    let mut cleaned: HashMap<String, i64> = HashMap::new();
    for (name, count) in &stock_data {
        let name = name.trim().to_string();
        if name.is_empty() {
            continue;
        }
        let c = count.as_f64().map(|v| (v as i64).max(0)).unwrap_or(0);
        cleaned.insert(name, c);
    }
    save_stock_snapshot(&month, &cleaned);
    json!({"ok": true})
}

#[tauri::command]
pub fn get_stock_history(
    session_token: String,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "textbooks.view");

    let history = load_stock_history();
    json!(history)
}

#[tauri::command]
pub fn set_level_textbook(
    session_token: String,
    level: String,
    textbook_names: Vec<String>,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "textbooks.modify");

    let level = level.trim().to_string();
    if level.is_empty() {
        return json!({"ok": false, "error": "等級不能為空。"});
    }
    let names: Vec<String> = textbook_names.iter().map(|n| n.trim().to_string()).filter(|n| !n.is_empty()).collect();
    let mut settings = load_settings();
    if !names.is_empty() {
        settings.level_textbook.insert(level, names);
    } else {
        settings.level_textbook.remove(&level);
    }
    save_settings(&settings);
    json!({"ok": true})
}

#[tauri::command]
pub fn set_level_next(
    session_token: String,
    level: String,
    next_level: String,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "settings.modify");

    let level = level.trim().to_string();
    let next_level = next_level.trim().to_string();
    if level.is_empty() {
        return json!({"ok": false, "error": "等級不能為空。"});
    }
    let mut settings = load_settings();
    if !next_level.is_empty() {
        settings.level_next.insert(level, next_level);
    } else {
        settings.level_next.remove(&level);
    }
    save_settings(&settings);
    json!({"ok": true})
}
