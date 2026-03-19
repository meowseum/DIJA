use chrono::{Datelike, Duration, NaiveDate};
use serde_json::{json, Value};
use std::collections::HashMap;
use std::sync::Mutex;
use tracing::info;

use crate::auth::db::AuthDb;
use crate::auth::session::SessionStore;
use crate::config::get_eps_template_path;
use crate::storage::*;

static EPS_ITEMS_CACHE: Mutex<Option<Vec<EpsItem>>> = Mutex::new(None);

#[derive(Debug, Clone)]
struct EpsItem {
    name: String,
    price: i64,
    section: String,
}

fn parse_hk_price(s: &str) -> Option<i64> {
    let cleaned = s.trim().replace("HK$", "").replace(",", "").replace(" ", "");
    if cleaned.is_empty() {
        return None;
    }
    cleaned.parse::<f64>().ok().map(|v| v as i64)
}

fn parse_eps_template(app_handle: &tauri::AppHandle) -> Vec<EpsItem> {
    {
        let cache = EPS_ITEMS_CACHE.lock().unwrap();
        if let Some(ref items) = *cache {
            return items.clone();
        }
    }

    let path = get_eps_template_path(app_handle);
    if !path.exists() {
        tracing::warn!("EPS template not found: {:?}", path);
        return Vec::new();
    }

    let mut rdr = match csv::ReaderBuilder::new().has_headers(false).from_path(&path) {
        Ok(r) => r,
        Err(_) => return Vec::new(),
    };

    let rows: Vec<csv::StringRecord> = rdr.records().flatten().collect();
    let mut items = Vec::new();
    let mut section = "class".to_string();

    for (i, row) in rows.iter().enumerate() {
        if i < 3 {
            continue;
        }
        let col_a = row.get(0).unwrap_or("").trim();
        let col_b = row.get(1).unwrap_or("").trim();

        if col_a == "書" {
            section = "book".to_string();
            continue;
        }
        if col_a == "其他" {
            section = "other".to_string();
            continue;
        }

        if col_a.is_empty() || col_a.contains("Sub-total") || col_a.contains("Total") {
            continue;
        }
        if row.iter().any(|c| c.contains("Sub-total")) {
            continue;
        }

        let price = match parse_hk_price(col_b) {
            Some(p) => p,
            None => continue,
        };

        items.push(EpsItem {
            name: col_a.to_string(),
            price,
            section: section.clone(),
        });
    }

    let mut cache = EPS_ITEMS_CACHE.lock().unwrap();
    *cache = Some(items.clone());
    items
}

#[tauri::command]
pub fn load_eps_items(
    session_token: String,
    app_handle: tauri::AppHandle,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "eps.view");

    let items = parse_eps_template(&app_handle);
    let items_json: Vec<Value> = items.iter().map(|i| json!({"name": i.name, "price": i.price, "section": i.section})).collect();
    json!({"ok": true, "items": items_json})
}

#[tauri::command]
pub fn load_eps_record(
    session_token: String,
    app_handle: tauri::AppHandle,
    date_str: String,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "eps.view");

    let items = parse_eps_template(&app_handle);
    let rows = load_eps_records(&date_str);
    let audit = load_eps_audit(&date_str);

    // Build lookup: (item_name, period) -> row
    let mut lookup: HashMap<(String, String), HashMap<String, String>> = HashMap::new();
    for r in &rows {
        let key = (
            r.get("item_name").unwrap_or(&String::new()).trim().to_string(),
            r.get("period").unwrap_or(&String::new()).trim().to_string(),
        );
        lookup.insert(key, r.clone());
    }

    // Yesterday for carry-over
    let yesterday = NaiveDate::parse_from_str(&date_str, "%Y-%m-%d")
        .ok()
        .map(|d| (d - Duration::days(1)).format("%Y-%m-%d").to_string())
        .unwrap_or_default();

    let past_day_carry = if !yesterday.is_empty() { get_eps_after_total(&yesterday) } else { 0 };

    // Carry over from yesterday if no before records
    let carry_over = if !yesterday.is_empty() && !has_eps_records_for_date(&date_str, "before") {
        get_eps_after_items(&yesterday)
    } else {
        HashMap::new()
    };

    let mut records_before: Vec<Value> = Vec::new();
    let mut records_after: Vec<Value> = Vec::new();

    for item in &items {
        for period in &["before", "after"] {
            let r = lookup.get(&(item.name.clone(), period.to_string()));
            let mut qty_k = r.and_then(|r| r.get("qty_K")).and_then(|s| s.trim().parse::<f64>().ok()).map(|v| v as i64).unwrap_or(0);
            let mut qty_l = r.and_then(|r| r.get("qty_L")).and_then(|s| s.trim().parse::<f64>().ok()).map(|v| v as i64).unwrap_or(0);
            let mut qty_hk = r.and_then(|r| r.get("qty_HK")).and_then(|s| s.trim().parse::<f64>().ok()).map(|v| v as i64).unwrap_or(0);

            if *period == "before" {
                if let Some(co) = carry_over.get(&item.name) {
                    qty_k += co.get("qty_K").copied().unwrap_or(0);
                    qty_l += co.get("qty_L").copied().unwrap_or(0);
                    qty_hk += co.get("qty_HK").copied().unwrap_or(0);
                }
            }

            let entry = json!({
                "item_name": item.name,
                "qty_K": qty_k,
                "qty_L": qty_l,
                "qty_HK": qty_hk,
            });

            if *period == "before" {
                records_before.push(entry);
            } else {
                records_after.push(entry);
            }
        }
    }

    let audit_int = |key: &str| -> i64 {
        audit.as_ref().and_then(|a| a.get(key)).and_then(|s| s.trim().parse::<f64>().ok()).map(|v| v as i64).unwrap_or(0)
    };

    let audit_data = json!({
        "operator_1_before": if audit_int("operator_1_before") != 0 { audit_int("operator_1_before") } else { audit_int("operator_1") },
        "operator_2_before": if audit_int("operator_2_before") != 0 { audit_int("operator_2_before") } else { audit_int("operator_2") },
        "operator_3_before": if audit_int("operator_3_before") != 0 { audit_int("operator_3_before") } else { audit_int("operator_3") },
        "operator_1_after": audit_int("operator_1_after"),
        "operator_2_after": audit_int("operator_2_after"),
        "operator_3_after": audit_int("operator_3_after"),
    });

    json!({
        "ok": true,
        "records": {"before": records_before, "after": records_after},
        "audit": audit_data,
        "past_day_carry": past_day_carry,
    })
}

#[tauri::command]
pub fn save_eps_record(
    session_token: String,
    app_handle: tauri::AppHandle,
    date_str: String,
    records: Value,
    audit: Value,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "eps.modify");

    let items = parse_eps_template(&app_handle);
    let mut csv_rows: Vec<HashMap<String, String>> = Vec::new();
    let mut sheet_before: i64 = 0;
    let mut sheet_after: i64 = 0;

    for period in &["before", "after"] {
        let period_items = records.get(*period).and_then(|v| v.as_array()).cloned().unwrap_or_default();
        for (idx, entry) in period_items.iter().enumerate() {
            if idx >= items.len() {
                break;
            }
            let item = &items[idx];
            let qty_k = entry.get("qty_K").and_then(|v| v.as_i64()).unwrap_or(0).max(0);
            let qty_l = entry.get("qty_L").and_then(|v| v.as_i64()).unwrap_or(0).max(0);
            let qty_hk = entry.get("qty_HK").and_then(|v| v.as_i64()).unwrap_or(0).max(0);
            let subtotal = item.price * (qty_k + qty_l + qty_hk);

            if *period == "before" { sheet_before += subtotal; } else { sheet_after += subtotal; }

            let mut row = HashMap::new();
            row.insert("date".to_string(), date_str.clone());
            row.insert("item_name".to_string(), item.name.clone());
            row.insert("item_price".to_string(), item.price.to_string());
            row.insert("item_section".to_string(), item.section.clone());
            row.insert("qty_K".to_string(), qty_k.to_string());
            row.insert("qty_L".to_string(), qty_l.to_string());
            row.insert("qty_HK".to_string(), qty_hk.to_string());
            row.insert("subtotal".to_string(), subtotal.to_string());
            row.insert("period".to_string(), period.to_string());
            csv_rows.push(row);
        }
    }

    save_eps_records(&date_str, &csv_rows);

    // Compute audit
    let yesterday = NaiveDate::parse_from_str(&date_str, "%Y-%m-%d")
        .ok()
        .map(|d| (d - Duration::days(1)).format("%Y-%m-%d").to_string())
        .unwrap_or_default();
    let past_day_carry = if !yesterday.is_empty() { get_eps_after_total(&yesterday) } else { 0 };

    let op1b = audit.get("operator_1_before").and_then(|v| v.as_i64()).unwrap_or(0).max(0);
    let op2b = audit.get("operator_2_before").and_then(|v| v.as_i64()).unwrap_or(0).max(0);
    let op3b = audit.get("operator_3_before").and_then(|v| v.as_i64()).unwrap_or(0).max(0);
    let op1a = audit.get("operator_1_after").and_then(|v| v.as_i64()).unwrap_or(0).max(0);
    let op2a = audit.get("operator_2_after").and_then(|v| v.as_i64()).unwrap_or(0).max(0);
    let op3a = audit.get("operator_3_after").and_then(|v| v.as_i64()).unwrap_or(0).max(0);
    let ops_sum_before = op1b + op2b + op3b;
    let ops_sum_after = op1a + op2a + op3a;
    let calculated_total = ops_sum_before + past_day_carry;
    let status = if calculated_total == sheet_before { "OK" } else { "MISMATCH" };
    let status_after = if ops_sum_after == sheet_after { "OK" } else { "MISMATCH" };
    let status_audit = if (ops_sum_before + ops_sum_after) == (sheet_before + sheet_after) { "OK" } else { "MISMATCH" };

    let mut audit_row = HashMap::new();
    audit_row.insert("date".to_string(), date_str.clone());
    audit_row.insert("operator_1_before".to_string(), op1b.to_string());
    audit_row.insert("operator_2_before".to_string(), op2b.to_string());
    audit_row.insert("operator_3_before".to_string(), op3b.to_string());
    audit_row.insert("operator_1_after".to_string(), op1a.to_string());
    audit_row.insert("operator_2_after".to_string(), op2a.to_string());
    audit_row.insert("operator_3_after".to_string(), op3a.to_string());
    audit_row.insert("operators_sum_before".to_string(), ops_sum_before.to_string());
    audit_row.insert("operators_sum_after".to_string(), ops_sum_after.to_string());
    audit_row.insert("sheet_before".to_string(), sheet_before.to_string());
    audit_row.insert("sheet_after".to_string(), sheet_after.to_string());
    audit_row.insert("past_day_carry".to_string(), past_day_carry.to_string());
    audit_row.insert("calculated_total".to_string(), calculated_total.to_string());
    audit_row.insert("status".to_string(), status.to_string());
    audit_row.insert("status_after".to_string(), status_after.to_string());
    audit_row.insert("status_audit".to_string(), status_audit.to_string());
    save_eps_audit(&date_str, &audit_row);

    info!("EPS record saved for {}: status={} status_after={} status_audit={}", date_str, status, status_after, status_audit);
    json!({
        "ok": true,
        "status": status,
        "status_after": status_after,
        "status_audit": status_audit,
        "calculated_total": calculated_total,
        "operators_sum_before": ops_sum_before,
        "operators_sum_after": ops_sum_after,
        "past_day_carry": past_day_carry,
        "sheet_before": sheet_before,
        "sheet_after": sheet_after,
    })
}

#[tauri::command]
pub fn export_eps_csv(
    session_token: String,
    app_handle: tauri::AppHandle,
    date_str: String,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "eps.export");

    use std::fmt::Write;

    let items = parse_eps_template(&app_handle);
    let rows = load_eps_records(&date_str);

    // Merge quantities per item
    let mut merged: HashMap<String, [i64; 3]> = HashMap::new();
    for r in &rows {
        let name = r.get("item_name").unwrap_or(&String::new()).trim().to_string();
        let entry = merged.entry(name).or_insert([0, 0, 0]);
        for (i, key) in ["qty_K", "qty_L", "qty_HK"].iter().enumerate() {
            entry[i] += r.get(*key).and_then(|s| s.trim().parse::<f64>().ok()).map(|v| v as i64).unwrap_or(0);
        }
    }

    let day_names = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"];
    let (date_display, dow) = NaiveDate::parse_from_str(&date_str, "%Y-%m-%d")
        .ok()
        .map(|d| (d.format("%Y/%m/%d").to_string(), day_names[d.weekday().num_days_from_monday() as usize].to_string()))
        .unwrap_or((date_str.clone(), String::new()));
    let year = if date_str.len() >= 4 { &date_str[..4] } else { "" };

    let hk_fmt = |v: i64| -> String { format!("HK${}", v) };
    let price_fmt = |v: i64| -> String { if v >= 1000 { format!("HK${},{}  ", v / 1000, v % 1000) } else { format!("HK${} ", v) } };

    let mut h = String::new();
    let _ = write!(h, r#"<html><head><meta charset="utf-8">
<style>
@page {{ margin: 0.17in 0.16in 0.17in 0.12in; }}
body {{ margin: 0; }}
table {{ border-collapse: collapse; table-layout: fixed; }}
td {{ font-family: "Times New Roman", serif; font-size: 14pt; vertical-align: middle;
     white-space: nowrap; padding: 1px 4px; }}
.hdr-year {{ font-size: 12pt; text-align: right; }}
.hdr-label {{ font-size: 12pt; text-align: left; }}
.title {{ font-size: 16pt; font-weight: 700; text-align: center; }}
.col-hdr {{ font-size: 16pt; font-weight: 700; text-align: center;
           border: 1pt solid black; border-bottom: .5pt solid black; }}
.col-hdr-left {{ font-size: 16pt; border-top: 1pt solid black; border-bottom: .5pt solid black;
                border-left: 1pt solid black; }}
.item-name {{ font-weight: 700; border-left: 1pt solid black; border-bottom: .5pt solid black;
             border-right: none; }}
.item-price {{ font-weight: 700; text-align: left; border-bottom: .5pt solid black; }}
.qty-x {{ text-align: right; border-left: .5pt solid black; border-bottom: .5pt solid black; }}
.qty-n {{ text-align: center; border-bottom: .5pt solid black; }}
.qty-eq {{ text-align: center; border-bottom: .5pt solid black; }}
.qty-val {{ text-align: right; border-right: .5pt solid black; border-bottom: .5pt solid black; }}
.row-total {{ text-align: right; border: .5pt solid black; }}
.row-remark {{ text-align: left; border-right: 1pt solid black; border-bottom: .5pt solid black;
              border-left: .5pt solid black; }}
.section-hdr {{ font-size: 16pt; font-weight: 700; text-align: left;
               border-top: 1pt solid black; border-left: 1pt solid black;
               border-bottom: .5pt solid black; }}
.section-hdr-r {{ border-top: 1pt solid black; border-right: 1pt solid black;
                 border-bottom: .5pt solid black; }}
.sub-label {{ font-weight: 700; text-align: right; }}
.sub-val {{ text-align: right; border-bottom: .5pt solid black; }}
.loc-total {{ font-size: 9pt; color: windowtext; text-align: left; }}
.total-label {{ font-weight: 700; text-align: right; }}
.total-val {{ text-align: right; border-top: .5pt solid black; border-bottom: 1pt solid black; }}
.date-label {{ font-weight: 700; text-align: right; }}
.date-val {{ font-weight: 700; text-align: right; border-bottom: 1pt solid black; }}
.dow {{ font-weight: 700; text-align: center; }}
.spacer-row td {{ height: 9pt; }}
</style></head><body>
<table>
<col width=9><col width=241><col width=106>
<col width=21><col width=37><col width=21><col width=85>
<col width=21><col width=37><col width=21><col width=81>
<col width=21><col width=37><col width=21><col width=83>
<col width=94><col width=93>
"#);

    // Row 1: Year + Date
    let _ = write!(h, r#"<tr height=24>
 <td></td><td class="hdr-year">{}</td><td class="hdr-label">學年</td>
 <td colspan=11></td>
 <td class="date-label">Date:</td><td class="date-val">{}</td>
 <td class="dow">{}</td>
</tr>"#, year, date_display, dow);

    // Row 2: Title
    let _ = write!(h, r#"<tr height=22>
 <td></td><td colspan=16 class="title">EPS 收支紀錄 (旺角校 - 星期一至星期六)</td>
</tr>"#);

    // Row 3: Column headers
    let _ = write!(h, r#"<tr height=28>
 <td></td><td class="col-hdr-left">&nbsp;</td><td class="col-hdr-left">&nbsp;</td>
 <td colspan=4 class="col-hdr">K</td>
 <td colspan=4 class="col-hdr" style="border-left:none">L</td>
 <td colspan=4 class="col-hdr" style="border-left:none">HK</td>
 <td class="col-hdr" style="border-left:none">Total</td>
 <td class="col-hdr" style="border-left:none">Remarks</td>
</tr>"#);

    let mut section_totals: HashMap<String, i64> = HashMap::new();
    let mut loc_totals: HashMap<String, [i64; 3]> = HashMap::new();
    for s in &["class", "book", "other"] {
        section_totals.insert(s.to_string(), 0);
        loc_totals.insert(s.to_string(), [0, 0, 0]);
    }
    let mut current_section: Option<String> = None;
    let mut grand_total: i64 = 0;

    let subtotal_row = |section: &str, st: &HashMap<String, i64>, lt: &HashMap<String, [i64; 3]>| -> String {
        let total = st.get(section).copied().unwrap_or(0);
        let locs = lt.get(section).copied().unwrap_or([0, 0, 0]);
        format!(r#"<tr height=26>
 <td></td><td></td><td></td>
 <td colspan=3 class="loc-total">{}</td><td></td>
 <td colspan=3 class="loc-total">{}</td><td></td>
 <td colspan=3 class="loc-total">{}</td>
 <td class="sub-label">Sub-total:</td><td class="sub-val">{}</td><td></td>
</tr>"#, hk_fmt(locs[0]), hk_fmt(locs[1]), hk_fmt(locs[2]), total)
    };

    let spacer = r#"<tr class="spacer-row"><td colspan=17></td></tr>"#;

    for item in &items {
        if current_section.as_deref() != Some(&item.section) {
            if current_section.as_deref() == Some("class") {
                h.push_str(&subtotal_row("class", &section_totals, &loc_totals));
                h.push_str(spacer);
                let _ = write!(h, r#"<tr height=29>
 <td></td><td colspan=16 class="section-hdr">書</td>
</tr>"#);
            } else if current_section.as_deref() == Some("book") && item.section == "other" {
                h.push_str(&subtotal_row("book", &section_totals, &loc_totals));
                h.push_str(spacer);
                let _ = write!(h, r#"<tr height=29>
 <td></td><td colspan=2 class="section-hdr">其他</td>
 <td colspan=4 class="section-hdr-r">&nbsp;</td>
 <td colspan=4 class="section-hdr-r">&nbsp;</td>
 <td colspan=4 class="section-hdr-r">&nbsp;</td>
 <td class="section-hdr-r">&nbsp;</td><td class="section-hdr-r">&nbsp;</td>
</tr>"#);
            }
            current_section = Some(item.section.clone());
        }

        let m = merged.get(&item.name).copied().unwrap_or([0, 0, 0]);
        let sub_k = item.price * m[0];
        let sub_l = item.price * m[1];
        let sub_hk = item.price * m[2];
        let total = sub_k + sub_l + sub_hk;
        *section_totals.get_mut(&item.section).unwrap() += total;
        let lt = loc_totals.get_mut(&item.section).unwrap();
        lt[0] += sub_k; lt[1] += sub_l; lt[2] += sub_hk;
        grand_total += total;

        let qk_s = if m[0] != 0 { m[0].to_string() } else { "&nbsp;".to_string() };
        let ql_s = if m[1] != 0 { m[1].to_string() } else { "&nbsp;".to_string() };
        let qh_s = if m[2] != 0 { m[2].to_string() } else { "&nbsp;".to_string() };

        let _ = write!(h, r#"<tr height=22>
 <td></td>
 <td class="item-name">{}</td>
 <td class="item-price">{}</td>
 <td class="qty-x">X</td><td class="qty-n">{}</td><td class="qty-eq">=</td><td class="qty-val">{}</td>
 <td class="qty-x">X</td><td class="qty-n">{}</td><td class="qty-eq">=</td><td class="qty-val">{}</td>
 <td class="qty-x">X</td><td class="qty-n">{}</td><td class="qty-eq">=</td><td class="qty-val">{}</td>
 <td class="row-total">{}</td>
 <td class="row-remark">&nbsp;</td>
</tr>"#, item.name, price_fmt(item.price), qk_s, sub_k, ql_s, sub_l, qh_s, sub_hk, total);
    }

    if let Some(ref section) = current_section {
        h.push_str(&subtotal_row(section, &section_totals, &loc_totals));
    }

    let _ = write!(h, r#"<tr height=21>
 <td></td><td></td><td></td>
 <td colspan=11></td>
 <td colspan=2 class="total-label">Total:</td>
 <td class="total-val">{}</td><td></td>
</tr>"#, grand_total);

    h.push_str("</table></body></html>");

    let filename = format!("EPS_{}.htm", date_str);

    // Save to custom output path if configured
    let config = load_app_config();
    let eps_output_path = config.get("eps_output_path").cloned().unwrap_or_default();
    let mut saved_path = String::new();
    if !eps_output_path.is_empty() {
        let out_dir = std::path::Path::new(&eps_output_path);
        if out_dir.is_dir() {
            let out_file = out_dir.join(&filename);
            if std::fs::write(&out_file, &h).is_ok() {
                saved_path = out_file.to_string_lossy().to_string();
            }
        }
    }

    json!({"ok": true, "content": h, "filename": filename, "saved_path": saved_path})
}

#[tauri::command]
pub fn list_eps_dates_endpoint(
    session_token: String,
    sessions: tauri::State<'_, SessionStore>,
    auth_db: tauri::State<'_, AuthDb>,
) -> Value {
    let _session = crate::require_auth!(sessions, auth_db, &session_token, "eps.view");

    json!({"ok": true, "dates": list_eps_dates()})
}
