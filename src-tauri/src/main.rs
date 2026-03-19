// Prevents additional console window on Windows in release
#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

mod auth;
mod commands;
mod config;
mod docx;
mod logger;
mod models;
mod schedule;
mod sku;
mod storage;

use std::collections::HashMap;
use std::sync::Mutex;

fn main() {
    logger::init_logger();

    // Initialise auth database (creates tables on first run)
    let auth_conn = auth::db::init_database();

    tauri::Builder::default()
        .plugin(tauri_plugin_updater::Builder::new().build())
        .plugin(tauri_plugin_process::init())
        // Managed state for auth
        .manage::<auth::db::AuthDb>(Mutex::new(auth_conn))
        .manage::<auth::session::SessionStore>(Mutex::new(HashMap::new()))
        .invoke_handler(tauri::generate_handler![
            // Auth (6)
            commands::auth::check_setup_needed,
            commands::auth::setup_admin,
            commands::auth::login,
            commands::auth::logout,
            commands::auth::get_current_user,
            commands::auth::change_password,
            // Admin (10)
            commands::admin::list_users,
            commands::admin::create_user,
            commands::admin::update_user,
            commands::admin::deactivate_user,
            commands::admin::reactivate_user,
            commands::admin::reset_password,
            commands::admin::list_all_permissions,
            commands::admin::list_role_permissions,
            commands::admin::set_role_permissions,
            commands::admin::list_roles,
            commands::admin::get_audit_log,
            // State (6)
            commands::state::load_state,
            commands::state::set_app_location,
            commands::state::set_tab_order,
            commands::state::set_eps_output_path,
            commands::state::set_zoom_level,
            commands::state::set_last_review_ts,
            // Classes (7)
            commands::classes::create_class,
            commands::classes::update_class,
            commands::classes::delete_class,
            commands::classes::end_class_action,
            commands::classes::save_student_counts,
            commands::classes::terminate_class_with_last_date,
            commands::classes::get_class_schedule,
            // Holidays (2)
            commands::holidays::add_holiday,
            commands::holidays::delete_holiday,
            // Postpones (4)
            commands::postpones::add_postpone,
            commands::postpones::add_postpone_manual,
            commands::postpones::get_make_up_date,
            commands::postpones::delete_postpone,
            // Overrides (2)
            commands::overrides::add_schedule_override,
            commands::overrides::delete_schedule_override,
            // Settings (7)
            commands::settings::add_setting,
            commands::settings::delete_setting,
            commands::settings::move_setting,
            commands::settings::export_settings_csv,
            commands::settings::import_settings_csv,
            commands::settings::set_level_price,
            commands::settings::adjust_level_prices,
            // Textbooks (7)
            commands::textbooks::set_textbook,
            commands::textbooks::delete_textbook,
            commands::textbooks::set_textbook_stock,
            commands::textbooks::save_monthly_stock,
            commands::textbooks::get_stock_history,
            commands::textbooks::set_level_textbook,
            commands::textbooks::set_level_next,
            // Documents (9)
            commands::documents::list_docx_templates,
            commands::documents::generate_docx,
            commands::documents::load_payment_template,
            commands::documents::load_makeup_template,
            commands::documents::get_promote_notice_data,
            commands::documents::generate_promote_notice,
            commands::documents::list_message_templates,
            commands::documents::load_message_content,
            commands::documents::set_message_category,
            // Calendar (1)
            commands::calendar::get_calendar_data,
            // Export/Import (2)
            commands::export_import::export_classes_csv,
            commands::export_import::import_classes_csv,
            // File ops (1)
            commands::file_ops::open_output_folder,
            // EPS (5)
            commands::eps::load_eps_items,
            commands::eps::load_eps_record,
            commands::eps::save_eps_record,
            commands::eps::export_eps_csv,
            commands::eps::list_eps_dates_endpoint,
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
