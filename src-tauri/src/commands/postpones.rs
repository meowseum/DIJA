use serde_json::{json, Value};
use tracing::info;
use uuid::Uuid;

use crate::auth::db::AuthDb;
use crate::auth::session::SessionStore;
use crate::models::{parse_date, PostponeRecord};
use crate::schedule::*;
use crate::storage::*;
use super::classes::{build_schedule_with_index_pub, make_up_date_for_pub};

#[tauri::command]
pub fn add_postpone(
    session_token: String,
    class_id: String,
    original_date: String,
    reason: String,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "postpones.add");

    let mut classes = load_classes();
    let class_record = match classes.iter().find(|c| c.id == class_id) {
        Some(c) => c.clone(),
        None => return json!({"ok": false, "error": "找不到班別。"}),
    };

    let original_dt = match parse_date(&original_date) {
        Some(d) => d,
        None => return json!({"ok": false, "error": "原定日期不正確。"}),
    };

    let holidays = load_holidays();
    let mut postpones = load_postpones();
    let overrides = load_overrides();
    let schedule = build_schedule_with_index_pub(&class_record, &holidays, &postpones, &overrides);
    let schedule_dates: std::collections::HashSet<String> = schedule.iter()
        .filter_map(|item| item.get("date").and_then(|v| v.as_str()).map(|s| s.to_string()))
        .collect();

    let (make_up_date, reactivated) = if !schedule_dates.contains(&original_date) {
        let progress = calculate_progress(
            &class_record.start_date, class_record.weekday as u32, class_record.lesson_total,
            &holidays, &postpones.iter().filter(|p| p.class_id == class_id).cloned().collect::<Vec<_>>(),
            &overrides.iter().filter(|o| o.class_id == class_id).cloned().collect::<Vec<_>>(),
        );
        let is_ended = progress["lessons_remaining"].as_i64().unwrap_or(0) == 0
            || class_record.status == "terminated";
        if !is_ended {
            return json!({"ok": false, "error": "原定日期不在日程。"});
        }
        let mud = make_up_date_for_pub(&class_record, &holidays, &postpones, &overrides, original_dt);
        // Increment lesson_total and possibly reactivate
        for record in &mut classes {
            if record.id == class_id {
                record.lesson_total += 1;
                let reactivated = record.status == "terminated";
                if reactivated {
                    record.status = "active".to_string();
                    info!("Class {} reactivated via add_postpone", class_id);
                }
                save_classes(&classes);
                return finish_add_postpone(&class_id, &original_date, &reason, &mud, reactivated, &mut postpones);
            }
        }
        (mud, false)
    } else {
        let mud = make_up_date_for_pub(&class_record, &holidays, &postpones, &overrides, original_dt);
        (mud, false)
    };

    finish_add_postpone(&class_id, &original_date, &reason, &make_up_date, reactivated, &mut postpones)
}

fn finish_add_postpone(
    class_id: &str, original_date: &str, reason: &str, make_up_date: &str,
    reactivated: bool, postpones: &mut Vec<PostponeRecord>,
) -> Value {
    let postpone = PostponeRecord {
        id: Uuid::new_v4().to_string(),
        class_id: class_id.to_string(),
        original_date: original_date.to_string(),
        reason: reason.trim().to_string(),
        make_up_date: make_up_date.to_string(),
    };
    postpones.push(postpone);
    save_postpones(postpones);
    info!("Postpone added for class {}: {} → {}", class_id, original_date, make_up_date);
    json!({"ok": true, "reactivated": reactivated})
}

#[tauri::command]
pub fn add_postpone_manual(
    session_token: String,
    class_id: String,
    original_date: String,
    make_up_date: String,
    reason: String,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "postpones.add");

    let mut classes = load_classes();
    let class_record = match classes.iter().find(|c| c.id == class_id) {
        Some(c) => c.clone(),
        None => return json!({"ok": false, "error": "找不到班別。"}),
    };
    if parse_date(&original_date).is_none() {
        return json!({"ok": false, "error": "原定日期不正確。"});
    }
    if parse_date(&make_up_date).is_none() {
        return json!({"ok": false, "error": "補課日期不正確。"});
    }

    let holidays = load_holidays();
    let mut postpones = load_postpones();
    let overrides = load_overrides();
    let schedule = build_schedule_with_index_pub(&class_record, &holidays, &postpones, &overrides);
    let schedule_dates: std::collections::HashSet<String> = schedule.iter()
        .filter_map(|item| item.get("date").and_then(|v| v.as_str()).map(|s| s.to_string()))
        .collect();

    // Validate the manual make-up date for ALL cases (in-schedule and ended). Previously
    // these checks only ran for the ended branch, so an in-schedule postpone could store a
    // colliding "next week" date that the scheduler silently relocated — making the course
    // look like it merged two lessons and lost one. Reject duplicates, holidays, and dates
    // that are not strictly after the original date.
    if schedule_dates.contains(&make_up_date) {
        return json!({"ok": false, "error": "補課日期重複，請選擇沒有課堂的日期。"});
    }
    if let Some(mu) = parse_date(&make_up_date) {
        if holiday_set(&holidays).contains(&mu) {
            return json!({"ok": false, "error": "補課日期遇到假期。"});
        }
        if let Some(orig) = parse_date(&original_date) {
            if mu <= orig {
                return json!({"ok": false, "error": "補課日期必須在原定日期之後。"});
            }
        }
    }

    let reactivated;
    if !schedule_dates.contains(&original_date) {
        let progress = calculate_progress(
            &class_record.start_date, class_record.weekday as u32, class_record.lesson_total,
            &holidays, &postpones.iter().filter(|p| p.class_id == class_id).cloned().collect::<Vec<_>>(),
            &overrides.iter().filter(|o| o.class_id == class_id).cloned().collect::<Vec<_>>(),
        );
        let is_ended = progress["lessons_remaining"].as_i64().unwrap_or(0) == 0
            || class_record.status == "terminated";
        if !is_ended {
            return json!({"ok": false, "error": "原定日期不在日程。"});
        }
        for record in &mut classes {
            if record.id == class_id {
                record.lesson_total += 1;
                reactivated = record.status == "terminated";
                if reactivated {
                    record.status = "active".to_string();
                    info!("Class {} reactivated via add_postpone_manual", class_id);
                }
                save_classes(&classes);
                return finish_add_postpone(&class_id, &original_date, &reason, &make_up_date, reactivated, &mut postpones);
            }
        }
        reactivated = false;
    } else {
        reactivated = false;
    }

    finish_add_postpone(&class_id, &original_date, &reason, &make_up_date, reactivated, &mut postpones)
}

#[tauri::command]
pub fn get_make_up_date(
    session_token: String,
    class_id: String,
    original_date: String,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "postpones.view");

    let classes = load_classes();
    let class_record = match classes.iter().find(|c| c.id == class_id) {
        Some(c) => c,
        None => return json!({"ok": false, "error": "找不到班別。"}),
    };
    let original_dt = match parse_date(&original_date) {
        Some(d) => d,
        None => return json!({"ok": false, "error": "原定日期不正確。"}),
    };
    let holidays = load_holidays();
    let postpones = load_postpones();
    let overrides = load_overrides();
    let schedule = build_schedule_with_index_pub(class_record, &holidays, &postpones, &overrides);
    let schedule_dates: std::collections::HashSet<String> = schedule.iter()
        .filter_map(|item| item.get("date").and_then(|v| v.as_str()).map(|s| s.to_string()))
        .collect();

    if !schedule_dates.contains(&original_date) {
        let progress = calculate_progress(
            &class_record.start_date, class_record.weekday as u32, class_record.lesson_total,
            &holidays, &postpones.iter().filter(|p| p.class_id == class_id).cloned().collect::<Vec<_>>(),
            &overrides.iter().filter(|o| o.class_id == class_id).cloned().collect::<Vec<_>>(),
        );
        let is_ended = progress["lessons_remaining"].as_i64().unwrap_or(0) == 0
            || class_record.status == "terminated";
        if !is_ended {
            return json!({"ok": false, "error": "原定日期不在日程。"});
        }
    }

    let mud = make_up_date_for_pub(class_record, &holidays, &postpones, &overrides, original_dt);
    json!({"ok": true, "make_up_date": mud})
}

/// Suggest substitute classes a student can attend as a make-up when they will miss a
/// lesson. Ranks ACTIVE classes of the SAME level by how close their nearby session's
/// progress is to the absent lesson's progress, then by date proximity, then same-location.
#[tauri::command]
pub fn suggest_makeup_classes(
    session_token: String,
    class_id: String,
    absent_date: String,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "classes.view");

    let classes = load_classes();
    let target = match classes.iter().find(|c| c.id == class_id) {
        Some(c) => c.clone(),
        None => return json!({"ok": false, "error": "找不到班別。"}),
    };
    let absent = match parse_date(&absent_date) {
        Some(d) => d,
        None => return json!({"ok": false, "error": "缺席日期不正確。"}),
    };

    let holidays = load_holidays();
    let postpones = load_postpones();
    let overrides = load_overrides();

    // Progress index of the absent lesson within the target class.
    let target_pps: Vec<_> = postpones.iter().filter(|p| p.class_id == class_id).cloned().collect();
    let target_ovs: Vec<_> = overrides.iter().filter(|o| o.class_id == class_id).cloned().collect();
    let target_sched = build_schedule_with_index_pub(&target, &holidays, &target_pps, &target_ovs);
    let target_index = target_sched
        .iter()
        .find(|item| item.get("date").and_then(|v| v.as_str()) == Some(absent_date.as_str()))
        .and_then(|item| item.get("index").and_then(|v| v.as_i64()))
        .unwrap_or(0);

    const WINDOW_DAYS: i64 = 28;
    let mut suggestions: Vec<Value> = Vec::new();
    for cand in classes.iter() {
        if cand.id == class_id || cand.status != "active" || cand.level != target.level {
            continue;
        }
        let cand_pps: Vec<_> = postpones.iter().filter(|p| p.class_id == cand.id).cloned().collect();
        let cand_ovs: Vec<_> = overrides.iter().filter(|o| o.class_id == cand.id).cloned().collect();
        let cand_sched = build_schedule_with_index_pub(cand, &holidays, &cand_pps, &cand_ovs);

        // Candidate session closest to the absent date (within the window).
        let mut best: Option<(i64, i64, String)> = None; // (|day distance|, index, date)
        for item in &cand_sched {
            let ds = match item.get("date").and_then(|v| v.as_str()) {
                Some(s) => s,
                None => continue,
            };
            let dd = match parse_date(ds) {
                Some(d) => d,
                None => continue,
            };
            let dist = (dd - absent).num_days().abs();
            if dist > WINDOW_DAYS {
                continue;
            }
            let idx = item.get("index").and_then(|v| v.as_i64()).unwrap_or(0);
            if best.as_ref().map_or(true, |(bk, _, _)| dist < *bk) {
                best = Some((dist, idx, ds.to_string()));
            }
        }
        let (day_distance, suggested_index, suggested_date) = match best {
            Some(b) => b,
            None => continue,
        };

        suggestions.push(json!({
            "class_id": cand.id,
            "sku": cand.sku,
            "level": cand.level,
            "location": cand.location,
            "weekday": cand.weekday,
            "time": cand.start_time,
            "teacher": cand.teacher,
            "classroom": cand.classroom,
            "suggested_date": suggested_date,
            "suggested_index": suggested_index,
            "target_index": target_index,
            "index_diff": (suggested_index - target_index).abs(),
            "day_distance": day_distance,
            "same_location": cand.location == target.location,
        }));
    }

    // Rank: progress closeness → date proximity → same-location preferred.
    suggestions.sort_by(|a, b| {
        let get = |v: &Value, k: &str| v[k].as_i64().unwrap_or(i64::MAX);
        get(a, "index_diff")
            .cmp(&get(b, "index_diff"))
            .then_with(|| get(a, "day_distance").cmp(&get(b, "day_distance")))
            .then_with(|| {
                let bl = b["same_location"].as_bool().unwrap_or(false);
                let al = a["same_location"].as_bool().unwrap_or(false);
                bl.cmp(&al)
            })
    });
    suggestions.truncate(5);

    json!({
        "ok": true,
        "target_index": target_index,
        "level": target.level,
        "suggestions": suggestions,
    })
}

#[tauri::command]
pub fn delete_postpone(
    session_token: String,
    postpone_id: String,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "postpones.delete");

    let postpones = load_postpones();
    let new_postpones: Vec<_> = postpones.into_iter().filter(|p| p.id != postpone_id).collect();
    save_postpones(&new_postpones);
    json!({"ok": true})
}
