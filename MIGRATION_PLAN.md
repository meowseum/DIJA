# Migration Plan: Python+Eel → Full Rust+Tauri

## Context

The DIJ前線助手 app currently uses Python (Eel) which launches Chrome in app mode as the UI layer. This causes two key problems: (1) the app shows a pixelated Chrome icon instead of the app's own icon, and (2) Chrome process instability can crash the app. Since the software is still early in development and will grow significantly, now is the right time to migrate to Tauri (Rust backend + embedded OS WebView) — producing a proper standalone desktop application while retaining the existing HTML/CSS/JS frontend.

---

## Phase 0: Backup

- Copy the entire `c:\Users\cjeff\Desktop\スクリプト\DIJ\Log` to `c:\Users\cjeff\Desktop\スクリプト\DIJ\Log_backup_pre_tauri`
- Separately copy the `data/` directory as an additional safety net
- Initialize a git repository in the project root (if not already) so each phase can be committed and rolled back

---

## Phase 1: Install Toolchain & Scaffold Tauri Project

### 1.1 Prerequisites
- Install Rust via `rustup` (stable toolchain)
- Install Tauri CLI: `cargo install tauri-cli`
- No Node.js needed (frontend is plain HTML/CSS/JS)

### 1.2 Create `src-tauri/` directory with:

```
src-tauri/
  Cargo.toml
  tauri.conf.json
  build.rs
  icons/icon.ico          ← copy from root icon.ico
  src/
    main.rs               ← Tauri entry point, command registration
    commands/
      mod.rs              ← re-exports all command modules
      state.rs            ← load_state, set_app_location, set_tab_order, set_eps_output_path, set_last_review_ts
      classes.rs          ← create_class, update_class, delete_class, end_class_action, get_class_schedule, terminate_class_with_last_date, save_student_counts
      holidays.rs         ← add_holiday, delete_holiday
      postpones.rs        ← add_postpone, add_postpone_manual, get_make_up_date, delete_postpone
      overrides.rs        ← add_schedule_override, delete_schedule_override
      settings.rs         ← add_setting, delete_setting, move_setting, export/import_settings_csv, set_level_price, adjust_level_prices
      textbooks.rs        ← set_textbook, delete_textbook, set_textbook_stock, save_monthly_stock, get_stock_history, set_level_textbook, set_level_next
      documents.rs        ← load_payment_template, load_makeup_template, list_docx_templates, generate_docx, get_promote_notice_data, generate_promote_notice, list_message_templates, load_message_content, set_message_category
      calendar.rs         ← get_calendar_data
      export_import.rs    ← export_classes_csv, import_classes_csv
      file_ops.rs         ← open_output_folder
      eps.rs              ← load_eps_items, load_eps_record, save_eps_record, export_eps_csv, list_eps_dates_endpoint
    models.rs             ← ClassRecord, HolidayRange, PostponeRecord, LessonOverride
    storage.rs            ← CSV read/write with atomic writes
    schedule.rs           ← weekly schedule, postpones, overrides, progress
    sku.rs                ← SKU parsing/formatting
    config.rs             ← path resolution (data, template, output dirs)
    docx.rs               ← DOCX template rendering & text extraction
    logger.rs             ← rotating file logger
```

### 1.3 Cargo.toml dependencies
```toml
tauri = { version = "2", features = ["shell-open-api"] }
serde = { version = "1", features = ["derive"] }
serde_json = "1"
csv = "1.3"
chrono = { version = "0.4", features = ["serde"] }
uuid = { version = "1", features = ["v4"] }
regex = "1"
tracing = "0.1"
tracing-subscriber = "0.3"
tracing-appender = "0.2"
tempfile = "3"
zip = "2"                  # DOCX = ZIP archive
quick-xml = "0.36"         # DOCX XML manipulation
glob = "0.3"
opener = "0.7"             # os.startfile() equivalent
```

### 1.4 tauri.conf.json key settings
- Window: 1920×1080, position (0,0)
- Bundle icon: `icons/icon.ico`
- Frontend `distDir`: `../frontend`
- Bundle resources: `template/**`, `EPS  Blank 2026.csv`
- `data/` is NOT bundled — lives next to exe at runtime, persists across updates

---

## Phase 2: Port Core Backend Modules (pure logic, no Tauri dependency)

Port in this order. Each module is self-contained and unit-testable independently.

### 2.1 models.rs
**Source**: `backend/models.py`
- Port `ClassRecord` (22 fields), `HolidayRange`, `PostponeRecord`, `LessonOverride` as Rust structs with `#[derive(Serialize, Deserialize, Clone)]`
- Port `parse_date()` → `chrono::NaiveDate::parse_from_str`
- Port `_parse_bool()` (accepts "1"/"true"/"yes"/"y") and `_parse_int()` as helper functions
- Custom serde deserializer for the bool field variants

### 2.2 sku.rs
**Source**: `backend/sku.py`
- Port `parse_sku()` and `build_sku()` using `regex` crate with named captures
- Pure functions, straightforward port

### 2.3 schedule.rs
**Source**: `backend/schedule.py`
- `holiday_set()` → expand HolidayRange list to `HashSet<NaiveDate>`
- `generate_weekly_schedule()` → weekly date iteration skipping holidays
- `apply_postpones()` → remove original dates, insert makeup dates
- `apply_overrides()` → add/remove specific dates
- `calculate_progress()` → count elapsed/remaining lessons, compute end_date/next_lesson
- `_find_next_available_weekly()` → next weekday not in holidays/scheduled dates

### 2.4 config.rs
**Source**: `backend/config.py`
- `get_data_dir()` → `std::env::current_exe().parent() / "data"` (Tauri builds to single exe, data lives next to it)
- `data_file()` → append env suffix (dev/prod)
- `get_template_dir()` → Tauri resource directory for bundled read-only templates
- `get_output_dir()` → writable directory next to exe
- `get_eps_template_path()` → glob for `EPS*Blank*.csv`

### 2.5 storage.rs
**Source**: `backend/storage.py`
- Port all CSV headers as constants (CLASS_HEADERS, HOLIDAY_HEADERS, etc.)
- `_ensure_file()` → create CSV with headers if missing
- `_load_records()` → generic CSV reader using `csv::Reader` + serde
- `_save_records()` → atomic write via `tempfile::NamedTempFile` + `persist()`
- Port all typed load/save functions: `load_classes`, `save_classes`, `load_holidays`, etc.
- `load_settings()` → complex parser for polymorphic settings CSV (`type|value` format)
- `load_app_config()` / `save_app_config()`
- All EPS storage functions
- `backup_file()` → copy to `data/backups/` with timestamp

### 2.6 logger.rs
- Replace Python's `RotatingFileHandler(maxBytes=1MB, backupCount=3)` with `tracing` + `tracing-appender` or `file-rotate` crate for size-based rotation

---

## Phase 3: DOCX Module

**This is the hardest part of the migration.** The current app uses `docxtpl` (Jinja2 templating inside DOCX).

### What needs to work:
1. **Template rendering** (`generate_docx`, `generate_promote_notice`) — open .docx, replace `{{ VAR }}` placeholders, save
2. **Text extraction** (`load_payment_template`, `load_makeup_template`, `load_message_content`) — open .docx, extract paragraph text

### Approach: ZIP + XML string replacement
A .docx is a ZIP of XML files. Template variables appear in `word/document.xml`.

**For rendering:**
1. Open .docx as ZIP (`zip` crate)
2. Read `word/document.xml`
3. Handle "split run" problem: Word may split `{{ VAR }}` across multiple `<w:r>` XML elements
   - Parse with `quick-xml`, walk `<w:p>` paragraphs
   - Concatenate all `<w:t>` text within a paragraph
   - Apply regex replacements (`\{\{\s*VAR\s*\}\}` → value)
   - Redistribute text back into runs
4. Write modified XML into new ZIP, copy all other entries unchanged

**For text extraction:**
1. Open .docx as ZIP
2. Parse `word/document.xml`, extract all `<w:t>` text
3. Join paragraphs with newlines

### Why this works for our templates:
- Templates use only simple `{{ VAR }}` placeholders — no loops, no conditionals, no images
- Variables: `CLASS_NAME`, `CLASS_WEEK`, `TEACHER_NAME`, `CLASS_TIME`, `ROOM_NUMBER`, `WEEK_DAY`, plus `_1`/`_2` variants for dual-class template, plus promotion notice fields

### Public API:
```rust
pub fn render_docx_template(template: &Path, output: &Path, context: &HashMap<String, String>) -> Result<()>
pub fn extract_docx_text(path: &Path) -> Result<String>
```

---

## Phase 4: Port All 51 Tauri Commands

Each `@eel.expose` function becomes a `#[tauri::command]` function. Port in this order, testing each group before proceeding:

### Group 1: State (5 functions)
`load_state`, `set_app_location`, `set_tab_order`, `set_eps_output_path`, `set_last_review_ts`
— Exercises every storage function, proves the data layer works

### Group 2: Settings (7 functions)
`add_setting`, `delete_setting`, `move_setting`, `export_settings_csv`, `import_settings_csv`, `set_level_price`, `adjust_level_prices`
— Tests the complex polymorphic settings parser

### Group 3: Class CRUD (6 functions)
`create_class`, `update_class`, `delete_class`, `end_class_action`, `save_student_counts`, `terminate_class_with_last_date`
— Tests models + SKU parsing + storage + cascade deletes

### Group 4: Holidays/Postpones/Overrides (8 functions)
`add_holiday`, `delete_holiday`, `add_postpone`, `add_postpone_manual`, `get_make_up_date`, `delete_postpone`, `add_schedule_override`, `delete_schedule_override`
— Tests schedule logic integration

### Group 5: Schedule/Calendar (2 functions)
`get_class_schedule`, `get_calendar_data`
— Tests the full schedule pipeline with payment_due flag logic

### Group 6: Textbooks (7 functions)
`set_textbook`, `delete_textbook`, `set_textbook_stock`, `save_monthly_stock`, `get_stock_history`, `set_level_textbook`, `set_level_next`

### Group 7: Documents (10 functions)
`list_docx_templates`, `generate_docx`, `load_payment_template`, `load_makeup_template`, `list_message_templates`, `load_message_content`, `set_message_category`, `get_promote_notice_data`, `generate_promote_notice`, `list_eps_templates`
— Tests the DOCX module end-to-end

### Group 8: Export/Import + File ops (3 functions)
`export_classes_csv`, `import_classes_csv`, `open_output_folder`

### Group 9: EPS (5 functions)
`load_eps_items`, `load_eps_record`, `save_eps_record`, `export_eps_csv`, `list_eps_dates_endpoint`
— Tests EPS template parsing, carry-over logic, audit validation (OK/MISMATCH)

### Command pattern:
```rust
// Python original:
// @eel.expose
// def create_class(data): ... return {"ok": True, ...}

#[tauri::command]
fn create_class(data: serde_json::Value) -> Result<serde_json::Value, String> {
    // ... port logic ...
    Ok(serde_json::json!({"ok": true}))
}
```

All 51 commands registered in `main.rs` via `tauri::generate_handler![...]`.

---

## Phase 5: Frontend Migration

### 5.1 app.js — convert all 55 eel calls

Mechanical replacement of call pattern:

```js
// Before (Eel double-invoke):
const state = await eel.load_state()();
await eel.add_setting(entryType, value)();
const resp = await eel.create_class({sku: ...})();

// After (Tauri invoke):
const state = await window.__TAURI__.invoke('load_state');
await window.__TAURI__.invoke('add_setting', { entry_type: entryType, value: value });
const resp = await window.__TAURI__.invoke('create_class', { data: {sku: ...} });
```

Key change: positional args become named object properties matching the Rust function parameter names.

### 5.2 index.html — NO changes
### 5.3 styles.css — NO changes

---

## Phase 6: Build & Distribution

### 6.1 Development
```bash
cargo tauri dev    # launches app with hot-reload frontend + Rust recompile
```

### 6.2 Production
```bash
cargo tauri build  # produces .exe in src-tauri/target/release/bundle/
```

### 6.3 Distribution
Same workflow as current: copy the `.exe` to target devices. Data directory persists next to exe.

### 6.4 Expected output
- Single `.exe` (~3-8MB vs current 22MB)
- Proper app icon (no more pixelated Chrome icon)
- Native window, no Chrome process dependency

### 6.5 build.bat replacement
```bat
@echo off
cd /d %~dp0
cargo tauri build
echo Build complete. Output: src-tauri\target\release\bundle\
pause
```

---

## Phase 7: Testing — Full Verification Checklist

### 7.1 Unit tests (Rust `#[cfg(test)]` modules)

| Module | What to test |
|--------|-------------|
| models.rs | Round-trip serialize/deserialize for all 4 structs, `parse_date` edge cases, `_parse_bool` all variants |
| sku.rs | `parse_sku` valid/invalid inputs, `build_sku` formatting |
| schedule.rs | `holiday_set` date range expansion, `generate_weekly_schedule` with/without holidays, `apply_postpones` single/multiple, `apply_overrides` add/remove, `calculate_progress` elapsed/remaining/end_date |
| storage.rs | Round-trip read/write for every CSV type, atomic write safety, `load_settings` with all entry types |
| config.rs | Path resolution in dev mode |
| docx.rs | Template rendering with test .docx, text extraction, split-run edge case |

### 7.2 Integration tests — all 51 commands

Each test: set up test CSVs → call command → verify JSON response + CSV state on disk.

**State (5):**
- [ ] `load_state` — returns classes, holidays, postpones, settings, app_config, stock_history
- [ ] `set_app_location` — K/L/H/"", rejects invalid
- [ ] `set_tab_order` — saves comma-separated string
- [ ] `set_eps_output_path` — saves path
- [ ] `set_last_review_ts` — saves timestamp

**Classes (6):**
- [ ] `create_class` — validates SKU/date/lesson_total/level, creates record in CSV
- [ ] `update_class` — updates fields including booleans, handles SKU re-parse
- [ ] `delete_class` — removes class + cascades to postpones/overrides
- [ ] `end_class_action` — test "terminate", "merge", "promote" paths separately
- [ ] `save_student_counts` — bulk update
- [ ] `terminate_class_with_last_date` — counts lessons to date, sets status

**Holidays (2):**
- [ ] `add_holiday` — creates with UUID, validates dates
- [ ] `delete_holiday` — removes by ID

**Postpones (4):**
- [ ] `add_postpone` — auto-calculates makeup, handles ended-class reactivation
- [ ] `add_postpone_manual` — validates conflicts + holiday check
- [ ] `get_make_up_date` — returns correct calculated date
- [ ] `delete_postpone` — removes by ID

**Overrides (2):**
- [ ] `add_schedule_override` — add/remove actions, dedup
- [ ] `delete_schedule_override` — removes by ID

**Settings (7):**
- [ ] `add_setting` — teacher/room/level/time with dedup
- [ ] `delete_setting` — removes value
- [ ] `move_setting` — up/down reorder
- [ ] `export_settings_csv` — returns CSV content string
- [ ] `import_settings_csv` — parses, backs up, replaces
- [ ] `set_level_price` — validates level + price
- [ ] `adjust_level_prices` — applies delta to all prices

**Textbooks (7):**
- [ ] `set_textbook` — name + price
- [ ] `delete_textbook` — cascades to stock + level_textbook
- [ ] `set_textbook_stock` — name + count
- [ ] `save_monthly_stock` — YYYY-MM format, snapshot save
- [ ] `get_stock_history` — returns nested dict
- [ ] `set_level_textbook` — level + list of names
- [ ] `set_level_next` — level progression mapping

**Documents (10):**
- [ ] `load_payment_template` — finds .txt or .docx, extracts text
- [ ] `load_makeup_template` — finds specific .docx, extracts text
- [ ] `list_docx_templates` — lists template/print/ .docx files
- [ ] `generate_docx` — renders single-class template, verify output .docx
- [ ] `generate_docx` (class.docx) — dual-class with `_1`/`_2` vars
- [ ] `get_promote_notice_data` — calculates all auto-fill fields
- [ ] `generate_promote_notice` — renders promote_notice.docx
- [ ] `list_message_templates` — lists messages/ dir with categories
- [ ] `load_message_content` — reads .txt or .docx
- [ ] `set_message_category` — saves category mapping

**Calendar (1):**
- [ ] `get_calendar_data` — returns sessions with payment_due flags for date range

**Export/Import (2):**
- [ ] `export_classes_csv` — returns CSV content
- [ ] `import_classes_csv` — parses, backs up, replaces

**File ops (1):**
- [ ] `open_output_folder` — opens directory in OS file manager

**EPS (5):**
- [ ] `load_eps_items` — parses blank template CSV
- [ ] `load_eps_record` — loads records + audit + carry-over calculation
- [ ] `save_eps_record` — computes subtotals, audit status (OK/MISMATCH)
- [ ] `export_eps_csv` — generates HTML report
- [ ] `list_eps_dates_endpoint` — returns sorted date list

### 7.3 Side-by-side comparison
Before decommissioning the Python version:
1. Run both versions against the same `data/` directory
2. Call `load_state` in both, diff the JSON responses
3. Perform CRUD operations in Rust version, verify Python can still read the data (and vice versa)
4. Generate DOCX documents from both versions and compare outputs visually

### 7.4 Manual UI walkthrough
- [ ] App launches with correct icon in taskbar
- [ ] All 12 tabs render correctly
- [ ] Create/edit/delete a class
- [ ] Add/remove holidays, postpones, overrides
- [ ] Calendar week and month views display correctly
- [ ] Generate a DOCX document, verify content
- [ ] Generate a promotion notice
- [ ] EPS audit entry, save, export
- [ ] Import/export settings CSV
- [ ] Import/export classes CSV
- [ ] Fee guide output matches expected format
- [ ] Message templates load and display
- [ ] Textbook stock management works
- [ ] Tab reordering persists across restart
- [ ] Location switching (K/L/H) updates UI correctly

---

## Risks & Mitigations

| Risk | Mitigation |
|------|-----------|
| DOCX split-run problem garbles template output | Test with every existing .docx template; pre-process templates to consolidate runs if needed |
| Subtle date calculation differences (chrono vs Python datetime) | Unit test every schedule function with identical inputs, compare outputs |
| CSV encoding edge cases (BOM, mixed line endings) | Use `encoding_rs` for BOM detection; normalize on read |
| Unicode file paths (Japanese characters in parent dirs) | Use `PathBuf` throughout; test with current directory structure early |
| Settings CSV polymorphic parsing drift | Port parsing logic character-by-character from Python; test with production data |
| WebView2 missing on older Windows 10 | Tauri 2 bundles a WebView2 bootstrapper |

---

## Critical Files Reference

| Purpose | File |
|---------|------|
| All 51 exposed functions (authoritative logic) | `backend/app.py` |
| Data models | `backend/models.py` |
| CSV storage layer | `backend/storage.py` |
| Schedule calculations | `backend/schedule.py` |
| SKU parsing | `backend/sku.py` |
| Path/env config | `backend/config.py` |
| All 55 JS→Python call sites | `frontend/app.js` |
| UI structure | `frontend/index.html` |
| PyInstaller bundle config (resource reference) | `DIJA.spec` |
| App entry point | `main.py` |
