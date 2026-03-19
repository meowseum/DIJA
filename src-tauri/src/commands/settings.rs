use serde_json::{json, Value};

use crate::auth::db::AuthDb;
use crate::auth::session::SessionStore;
use crate::config::data_file;
use crate::storage::*;

#[tauri::command]
pub fn add_setting(
    session_token: String,
    entry_type: String,
    value: String,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "settings.modify");

    let entry_type = entry_type.trim().to_lowercase();
    let value = value.trim().to_string();
    if !["teacher", "room", "level", "time"].contains(&entry_type.as_str()) {
        return json!({"ok": false, "error": "設定類型不正確。"});
    }
    if value.is_empty() {
        return json!({"ok": false, "error": "請輸入內容。"});
    }
    let mut settings = load_settings();
    if let Some(list) = settings.list_mut(&entry_type) {
        if !list.contains(&value) {
            list.push(value);
            save_settings(&settings);
        }
    }
    json!({"ok": true})
}

#[tauri::command]
pub fn delete_setting(
    session_token: String,
    entry_type: String,
    value: String,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "settings.modify");

    let entry_type = entry_type.trim().to_lowercase();
    let value = value.trim().to_string();
    let mut settings = load_settings();
    if let Some(list) = settings.list_mut(&entry_type) {
        list.retain(|v| v != &value);
        save_settings(&settings);
    }
    json!({"ok": true})
}

#[tauri::command]
pub fn move_setting(
    session_token: String,
    entry_type: String,
    value: String,
    direction: String,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "settings.modify");

    let entry_type = entry_type.trim().to_lowercase();
    let value = value.trim().to_string();
    let direction = direction.trim().to_lowercase();
    let mut settings = load_settings();
    if let Some(list) = settings.list_mut(&entry_type) {
        if let Some(index) = list.iter().position(|v| v == &value) {
            if direction == "up" && index > 0 {
                list.swap(index - 1, index);
            } else if direction == "down" && index < list.len() - 1 {
                list.swap(index, index + 1);
            }
            save_settings(&settings);
            return json!({"ok": true});
        }
        return json!({"ok": false, "error": "找不到設定。"});
    }
    json!({"ok": false, "error": "找不到設定。"})
}

#[tauri::command]
pub fn export_settings_csv(
    session_token: String,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "settings.export");

    let settings = load_settings();
    let mut lines = vec!["type,value".to_string()];

    for entry_type in &["teacher", "room", "level", "time"] {
        let list = match *entry_type {
            "teacher" => &settings.teacher,
            "room" => &settings.room,
            "level" => &settings.level,
            "time" => &settings.time,
            _ => continue,
        };
        for value in list {
            let safe = value.replace('"', "\"\"");
            lines.push(format!("{},\"{}\"", entry_type, safe));
        }
    }
    for (level, price) in &settings.level_price {
        let safe = format!("{}|{}", level, price).replace('"', "\"\"");
        lines.push(format!("level_price,\"{}\"", safe));
    }
    for (name, price) in &settings.textbook {
        let safe = format!("{}|{}", name, price).replace('"', "\"\"");
        lines.push(format!("textbook,\"{}\"", safe));
    }
    for (k, v) in &settings.level_textbook {
        if !v.is_empty() {
            let safe = format!("{}|{}", k, v.join(",")).replace('"', "\"\"");
            lines.push(format!("level_textbook,\"{}\"", safe));
        }
    }
    for (k, v) in &settings.level_next {
        let safe = format!("{}|{}", k, v).replace('"', "\"\"");
        lines.push(format!("level_next,\"{}\"", safe));
    }
    for (name, count) in &settings.textbook_stock {
        let safe = format!("{}|{}", name, count).replace('"', "\"\"");
        lines.push(format!("textbook_stock,\"{}\"", safe));
    }

    json!({"ok": true, "content": lines.join("\n")})
}

#[tauri::command]
pub fn import_settings_csv(
    session_token: String,
    content: String,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "settings.import");

    backup_file(&data_file("settings.csv"));

    let mut new_settings = crate::storage::Settings::default();

    let mut rdr = csv::Reader::from_reader(content.as_bytes());
    for result in rdr.records() {
        let record = match result {
            Ok(r) => r,
            Err(_) => return json!({"ok": false, "error": "CSV 內容不正確。"}),
        };
        let entry_type = record.get(0).unwrap_or("").trim().to_lowercase();
        let value = record.get(1).unwrap_or("").trim().to_string();

        match entry_type.as_str() {
            "level_price" => {
                let parts: Vec<&str> = if value.contains('|') { value.splitn(2, '|').collect() } else { value.splitn(2, '=').collect() };
                if parts.len() == 2 {
                    let level = parts[0].trim().to_string();
                    let price = parts[1].trim();
                    if !level.is_empty() {
                        if let Ok(p) = price.parse::<f64>() {
                            new_settings.level_price.insert(level, (p as i64).max(0));
                        }
                    }
                }
            }
            "textbook" => {
                if let Some((name, price_str)) = value.split_once('|') {
                    let name = name.trim().to_string();
                    if !name.is_empty() {
                        let price = price_str.trim().parse::<f64>().map(|p| (p as i64).max(0)).unwrap_or(0);
                        new_settings.textbook.insert(name, price);
                    }
                }
            }
            "textbook_stock" => {
                if let Some((name, count_str)) = value.split_once('|') {
                    let name = name.trim().to_string();
                    if !name.is_empty() {
                        let count = count_str.trim().parse::<f64>().map(|c| (c as i64).max(0)).unwrap_or(0);
                        new_settings.textbook_stock.insert(name, count);
                    }
                }
            }
            "level_textbook" => {
                if let Some((k, rest)) = value.split_once('|') {
                    let k = k.trim().to_string();
                    let tb_list: Vec<String> = rest.split(',').map(|n| n.trim().to_string()).filter(|n| !n.is_empty()).collect();
                    if !k.is_empty() && !tb_list.is_empty() {
                        new_settings.level_textbook.insert(k, tb_list);
                    }
                }
            }
            "level_next" => {
                if let Some((k, v)) = value.split_once('|') {
                    let k = k.trim().to_string();
                    let v = v.trim().to_string();
                    if !k.is_empty() {
                        new_settings.level_next.insert(k, v);
                    }
                }
            }
            "teacher" | "room" | "level" | "time" => {
                if !value.is_empty() {
                    if let Some(list) = new_settings.list_mut(&entry_type) {
                        if !list.contains(&value) {
                            list.push(value);
                        }
                    }
                }
            }
            _ => {}
        }
    }

    save_settings(&new_settings);
    json!({"ok": true})
}

#[tauri::command]
pub fn set_level_price(
    session_token: String,
    level: String,
    price: f64,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "settings.modify");

    let level = level.trim().to_string();
    if level.is_empty() {
        return json!({"ok": false, "error": "等級不能為空。"});
    }
    let price_value = (price as i64).max(0);
    let mut settings = load_settings();
    settings.level_price.insert(level, price_value);
    save_settings(&settings);
    json!({"ok": true})
}

#[tauri::command]
pub fn adjust_level_prices(
    session_token: String,
    delta: f64,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "settings.modify");

    let delta_value = delta as i64;
    let mut settings = load_settings();
    for price in settings.level_price.values_mut() {
        *price = (*price + delta_value).max(0);
    }
    save_settings(&settings);
    json!({"ok": true})
}

// ---------------------------------------------------------------------------
// EPS config management
// ---------------------------------------------------------------------------

#[tauri::command]
pub fn set_eps_config(
    session_token: String,
    key: String,
    value: String,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "settings.modify");
    let key = key.trim().to_string();
    let value = value.trim().to_string();
    if key.is_empty() {
        return json!({"ok": false, "error": "Key 不能為空。"});
    }
    let mut settings = load_settings();
    settings.eps_config.insert(key, value);
    save_settings(&settings);
    json!({"ok": true})
}

#[tauri::command]
pub fn set_eps_item(
    session_token: String,
    item_type: String,
    name: String,
    price: f64,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "settings.modify");
    let name = name.trim().to_string();
    if name.is_empty() {
        return json!({"ok": false, "error": "名稱不能為空。"});
    }
    let price_val = (price as i64).max(0);
    let mut settings = load_settings();
    let list = match item_type.as_str() {
        "eps_book" => &mut settings.eps_book,
        "eps_other" => &mut settings.eps_other,
        "eps_special" => &mut settings.eps_special,
        _ => return json!({"ok": false, "error": "類型不正確。"}),
    };
    // Update existing or append
    if let Some(entry) = list.iter_mut().find(|(n, _)| n == &name) {
        entry.1 = price_val;
    } else {
        list.push((name, price_val));
    }
    save_settings(&settings);
    json!({"ok": true})
}

#[tauri::command]
pub fn delete_eps_item(
    session_token: String,
    item_type: String,
    name: String,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "settings.modify");
    let name = name.trim().to_string();
    let mut settings = load_settings();
    let list = match item_type.as_str() {
        "eps_book" => &mut settings.eps_book,
        "eps_other" => &mut settings.eps_other,
        "eps_special" => &mut settings.eps_special,
        _ => return json!({"ok": false, "error": "類型不正確。"}),
    };
    list.retain(|(n, _)| n != &name);
    save_settings(&settings);
    json!({"ok": true})
}
