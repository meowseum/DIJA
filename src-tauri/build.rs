use std::path::Path;

/// Emit `rerun-if-changed` for every file under the frontend directory.
///
/// Tauri embeds `frontendDist` into the release binary at compile time via
/// `generate_context!`, but `tauri_build` does not track the frontend folder — so a
/// frontend-only edit would not recompile the crate and the release `.exe` would ship a
/// STALE frontend. Tracking each frontend file forces a recompile (and re-embed) whenever
/// any asset changes.
fn track_frontend(dir: &Path) {
    if let Ok(entries) = std::fs::read_dir(dir) {
        for entry in entries.flatten() {
            let path = entry.path();
            if path.is_dir() {
                track_frontend(&path);
            } else if let Some(s) = path.to_str() {
                println!("cargo:rerun-if-changed={}", s);
            }
        }
    }
}

fn main() {
    track_frontend(Path::new("../frontend"));
    tauri_build::build()
}
