# Changelog

All notable changes to the E-volve SecureAssess Results Automation tool are documented in this file.

---

## [Unreleased] — 2026-04-01

### GUI Redesign (Phase 10)

#### Visual / Branding
- **Switched to light mode** — `ctk.set_appearance_mode("light")` with a clean white/grey palette
- **City & Guilds red brand colour** (`#e30613`) used throughout instead of purple
- **Updated all colour constants** — `CG_PURPLE` → `CG_RED`, `CG_PURPLE_HOVER` → `CG_RED_HOVER`, `CG_PURPLE_LIGHT` → `CG_RED_LIGHT`
- **Regenerated app icon** — white play triangle on C&G red rounded square (old purple icon deleted, regenerates on next launch)
- **Light mode design tokens**: `BG=#F5F5F7`, `SURFACE=#FFFFFF`, `ELEVATED=#EEEEEE`, `BORDER=#DDDDDD`, `TEXT=#1A1A2E`, `SUCCESS=#2E7D32`, `DANGER=#C62828`, `AMBER=#E65100`

#### Layout / UX
- **Removed empty space above tabs** — reduced content padding from `S6` to `S4`, eliminated dead space in all tabs
- **Tab bar redesigned** — app name label removed from left side; tabs now on the left; right side shows `E-volve SecureAssess Results Automation — Unofficial Tool v1.0.0`
- **Footer improved** — height increased from 30→40px, font size from 10→12, footer text colour changed from `TEXT_DIM` to `TEXT_MID` for better readability, button widths increased
- **Tab switching rewritten** — changed from `place`/`place_forget` to `pack`/`pack_forget` to fix sizing and empty-space issues
- **Window maximize fix** — uses `root.after()` scheduling for reliable maximize on Windows; re-maximizes after lock screen is destroyed

#### Bug Fixes
- **Fixed duplicate log handlers** — `self._log_handler` stores reference, old handler removed before adding new one on each automation run
- **Fixed bare `except:` clauses** in `main.py` — changed to `except (ValueError, TypeError):`
- **Fixed `safe_find()` exception chain** in `selenium_utils.py` — now uses `raise ... from e`
- **Fixed `excel_utils.py` column name** — `"E-Certificate"` updated to `"E-Certificate sent"` to match Phase 7.6 rename

### File Path Fixes
- **`config.py` `BASE_DIR` fix** — uses `os.path.dirname(os.path.abspath(__file__))` for module execution, `os.path.dirname(sys.executable)` for PyInstaller
- **`credentials.enc` and year folders** now correctly save to the project root directory, not inside the `evolve_results_automation/` package folder
- **`.gitignore` year folder pattern** — `[12][0-9][0-9][0-9]/` correctly matches root-level year folders like `2026/`

### Dead Code Removal
- **Removed `goto_page()`** from `selenium_utils.py` — dead code, never called
- **Removed `goto_page` import** from `main.py`
- **Removed `load_secure_credentials()` standalone function** from `secure_credentials.py` (still exists in dead `gui_cli.py`)
- **Removed `CURRENT_YEAR` and `LOGS_BASE` dead constants** from `config.py`

### DRY / Code Quality Fixes
- **`BASE_DIR` single source of truth** — `main.py` and `gui_tk.py` now import from `config.py` instead of recomputing
- **Credential file init** — `gui_tk.py` uses `manager.create_empty()` instead of calling private `_derive_key()` method
- **`list_credentials()`** — now returns only `{"username": ...}` dicts instead of full credentials with passwords
- **Error logging** — `secure_credentials.py` error messages now use `logging.error()` instead of `logging.info()`
- **`REPO_URL` constant** added to `gui_tk.py` — "View on GitHub" no longer derives URL from `ISSUES_URL.rsplit()`

### Documentation
- **`AUDIT_REPORT.md`** — fully updated against current codebase; all fixed items marked with ✅; stale line numbers corrected; priority fix list updated
- **`IMPLEMENTATION_ROADMAP.md`** — Phase 8 design system updated to reflect light mode + red palette; architecture description corrected from "grid-based" to "pack-based"; Phase 10 marked as complete
- **`CHANGELOG.md`** — this file created to document all changes
- **`README.md`** — previously updated with correct version (v1.0.0), GUI workflow, correct file paths, and dependencies

---

## [v1.0.0] — 2026-03-30

### Added
- CustomTkinter GUI with tab-based navigation (Dashboard, Accounts, Files, Help)
- Lock screen with AES-256 encrypted master password
- Real-time progress bar tracking 9 automation phases
- Inline toast notifications and completion/error banners
- Files tab for browsing year-based data folders
- Help tab with getting started guide and support links
- Excel lock detection before automation runs
- Custom app icon generated via Pillow
- `__main__.py` for `python -m evolve_results_automation` support

### Changed
- Year-based folder organization (results by exam completion year)
- Full pagination support (scrapes all pages, not just page 1)
- Automatic date filter management (sets filter to previous month start)
- PDF download resume across runs with incremental saves
- ChromeDriver auto-management via Selenium Manager
- Column renames for admin-friendly Excel layout

### Fixed
- `move_columns_to_end` silent failure (return value was discarded)
- `driver.implicitly_wait(6)` misuse (was not a sleep)
- Hash logic DRY violation (3 locations → 1)
- XPath constant DRY violation (2 locations → 1)
- Column schema DRY violation (2 locations → 1)
- Debug HTML file creation removed from `safe_find()`

---
