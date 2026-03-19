use std::path::PathBuf;

fn get_env() -> String {
    std::env::var("APP_ENV").unwrap_or_else(|_| "dev".to_string()).to_lowercase()
}

pub fn get_data_dir() -> PathBuf {
    // In dev/test builds, use CARGO_MANIFEST_DIR parent (project root)
    // In release builds, use the directory containing the executable
    let base = if cfg!(debug_assertions) {
        PathBuf::from(env!("CARGO_MANIFEST_DIR")).parent().unwrap().to_path_buf()
    } else {
        std::env::current_exe()
            .ok()
            .and_then(|p| p.parent().map(|pp| pp.to_path_buf()))
            .unwrap_or_else(|| {
                PathBuf::from(env!("CARGO_MANIFEST_DIR")).parent().unwrap().to_path_buf()
            })
    };
    let data_dir = base.join("data");
    std::fs::create_dir_all(&data_dir).ok();
    data_dir
}

pub fn data_file(filename: &str) -> PathBuf {
    let env = get_env();
    let data_dir = get_data_dir();
    if let Some((stem, suffix)) = filename.rsplit_once('.') {
        data_dir.join(format!("{}_{}.{}", stem, env, suffix))
    } else {
        data_dir.join(format!("{}_{}", filename, env))
    }
}

pub fn get_template_dir(app_handle: &tauri::AppHandle) -> PathBuf {
    // In production, use Tauri's resource directory
    if let Ok(resource_dir) = app_handle.path().resource_dir() {
        let template_dir = resource_dir.join("template");
        if template_dir.exists() {
            return template_dir;
        }
    }
    // Fallback for dev: project root / template
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .unwrap()
        .join("template")
}

pub fn get_eps_template_path(app_handle: &tauri::AppHandle) -> PathBuf {
    // In production, Tauri resource dir
    if let Ok(resource_dir) = app_handle.path().resource_dir() {
        let pattern = resource_dir.join("EPS*Blank*.csv");
        if let Ok(matches) = glob::glob(pattern.to_str().unwrap_or("")) {
            for entry in matches.flatten() {
                return entry;
            }
        }
    }
    // Fallback for dev
    let base = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .unwrap()
        .to_path_buf();
    let pattern = base.join("EPS*Blank*.csv");
    if let Ok(matches) = glob::glob(pattern.to_str().unwrap_or("")) {
        for entry in matches.flatten() {
            return entry;
        }
    }
    base.join("EPS  Blank 2026.csv")
}

pub fn get_output_dir() -> PathBuf {
    let base = std::env::current_exe()
        .ok()
        .and_then(|p| p.parent().map(|pp| pp.to_path_buf()))
        .unwrap_or_else(|| {
            PathBuf::from(env!("CARGO_MANIFEST_DIR"))
                .parent()
                .unwrap()
                .join("template")
        });
    let output_dir = base.join("output");
    std::fs::create_dir_all(&output_dir).ok();
    output_dir
}

use tauri::Manager;
