# Changelog

All notable changes to E-volve SecureAssess Automation are documented in this file.

---

## [v1.3.0] - 2026-04-13

### Added
- **Stop automation** - cooperative `threading.Event` stop with confirmation dialog. Run button shows "Stopping..." and restores on complete/error
- **Lock button** - header row next to Settings, returns to lock screen. Disabled during automation
- **Zero-accounts guidance** - proactive log hint on unlock if no credentials configured
- **Password strength indicator** - 4px colour bar + "Weak"/"Fair"/"Strong" text label on setup screen
- **Disk space warning** - `shutil.disk_usage` check before automation, messagebox if < 100 MB free
- **Excel backup** - `.bak` copy of each year's Excel file before automation starts, verified after write (warns on 0-byte backup)
- **Desktop shortcut** - VBScript `CreateShortcut` helper with `app.ico` icon. Button in Help dialog with messagebox feedback. Shows info message if shortcut already exists
- **Connection check** - two layers: `_preflight_check()` HTTP-200 for no-internet, Selenium login-field presence check for E-volve maintenance
- **Keyboard shortcuts** - Ctrl+S (Settings), Ctrl+H (Help, always available), Ctrl+L (Lock). One-time log hint on startup. Ctrl+S and Ctrl+L unbound when locked
- **Multi-year Quick Open** - custom year dropdown in Quick Open card, auto-discovers existing year folders. Refreshed after automation completes
- **Retry on failure** - `max_attempts=2` per account with fresh browser on retry
- **Persistent settings** - `settings.json` preserves Show browser, Desktop notifications, Minimize to tray, Scheduler, and Start with Windows toggles across sessions
- **Toast notifications** - Windows desktop notifications on automation complete/error via `winotify` (guarded import, toggle in Settings). Hi-res PNG icon extracted from ICO via Pillow, cached to temp dir with mtime invalidation
- **Auto update check** - background GitHub API check on startup, logs if a newer release is available. Fires once per session
- **Excel Analytics sheet** - single compact "Analytics" tab auto-generated in each year's Excel file with 6 KPI cards (Total Exams, Unique Candidates, Extra Time Candidates, Most Popular Exam, Busiest Month, Resit Conversion - all with percentage subtitles), monthly volume chart, exam breakdown horizontal bar chart (matching red), rebook opportunities table, failed resit tracker with colour-coded results, extra time candidates table, and auto-generated key insights. Hidden `_ChartData` sheet for chart data. Brand-consistent styling with light pink stripes and red accents matching the Results tab. Both tabs use red tab colour
- **Results tab styling** - 11pt data font, 11pt bold header, alternating light pink stripes, thin grey borders, red tab colour matching Analytics tab
- **Percent column reordered** - moved next to Result column for better readability
- **System tray** - minimize to tray on close (toggle in Settings), tray icon menu (Open / Open Excel / Open Reports / Run Now / Exit), first-close toast notification (shown once, stored in last_run.json), double-click opens app. Run Now hidden when locked. Tooltip shows lock state and last run info (refreshed after each run). Guarded import - degrades gracefully if `pystray` or `Pillow` missing
- **Scheduler** - daily schedule slot (HH:MM) with enable toggle in Settings, live format validation with red hint, daemon poll thread (every 30s), fires only when app is unlocked and idle, skips with log message if locked or already running. Pre-seeded on startup to prevent re-fire after restart
- **Start with Windows** - toggle in Settings creates/removes `.lnk` shortcut in Windows Startup folder via existing VBScript helper. Switch state reflects actual shortcut presence on open
- **Auto read-only on startup** - app opens directly in read-only mode on subsequent launches (lock screen only on first-time setup). Unlock via header button navigates to full lock screen with password reset option
- **Single-instance lock** - prevents multiple app instances writing to the same data files via OS-level file lock
- **Window title** - shows "Read Only" when locked, reverts to normal on unlock
- **PDF skip tracking** - `pdfs_skipped` counter added to `ProcessingStats`, surfaced in completion notification and log when PDF loop breaks unrecoverably

### Fixed
- **load_existing_results data loss** (`excel_utils.py`) - always loads from "Results" sheet instead of active sheet, preventing silent data wipe when Analytics was the active tab
- **Excel IndexError on malformed rows** (`excel_utils.py`) - `zip(headers, row_values)` instead of `range(len(headers))`
- **Windows path-length failure** (`parsing_utils.py`) - truncate sanitized filename to 200 chars
- **Silent date default** (`main.py`) - log warning on unparseable completed date instead of silently defaulting to current year
- **Dropdown popups gray Windows 11 titlebar** (`gui_tk.py`) - replaced `focus_set()` on `overrideredirect` Toplevel with click-outside-dismiss pattern, preventing window deactivation
- **Window appeared top-left then jumped** (`gui_tk.py`) - `_center_window` called synchronously on root creation
- **Update check silently failing** (`gui_tk.py`) - used non-existent `_msg_queue`; now uses `root.after` + `_log`
- **`_save_last_run` overwriting tray flag** (`gui_tk.py`) - now read-merge-write to preserve `tray_first_close_shown`
- **`_on_close` not stopping automation** (`gui_tk.py`) - now signals `stop_event` and joins thread before exit
- **Resit tracker sort bug** (`excel_utils.py`) - single bad date aborted the entire sort; bad dates now sort to start
- **Workbook file-handle leaks** (`excel_utils.py`) - all `load_workbook` sites wrapped in `try/finally` to prevent `PermissionError` on Windows
- **Version comparison false positive** (`gui_tk.py`) - string tuple comparison replaced with integer tuples
- **Year-folder mkdir side effect** (`config.py`) - read-only path lookups no longer create empty directories
- **Unclosed file handle in tray toast** (`gui_tk.py`) - `_on_close` first-close check now uses `with` statement
- **Double newline in log** (`main.py`) - removed trailing `\n` from "Chrome closed" message
- **Popup overlap** (`gui_tk.py`) - opening one dropdown while another is open no longer orphans the first

### Changed
- **Year dropdown** rebuilt as custom widget (frame+label+arrow) to match account selector style
- **Shortcut hint** logs immediately with startup messages instead of delayed `after(500, ...)`
- **Scale factor comment** (`selenium_utils.py`) - added explanation for headless `0.5`/`0.24` values
- **Settings dialog height** reduced to 463px, main window height increased to 560px
- **Account count label** shows "Locked" in read-only mode
- **Multi-monitor dialog positioning** (`gui_tk.py`) - Settings and Help dialogs center on the parent window's monitor via Windows `MonitorFromPoint` API, clamped to work area bounds
- **CHANGELOG.md** now tracked in git (removed from `.gitignore`)
- Removed dead imports (`glob`, `numbers`, `DataBarRule`, `BASE_DIR`), dead functions (`save_results`, `refresh_all_analytics`, `get_year_folder`), unused parameters, and stale `tray_first_close_shown` default
- Consolidated duplicated dropdown popup logic into generic `_show_popup`/`_close_popup`, shutdown logic into `_graceful_shutdown`, control reset into `_reset_controls`
- Replaced 5 scattered year-folder glob patterns with shared `list_year_folders`/`list_year_excel_files` helpers in `config.py`
- Replaced duplicated Results tab formatting with module-level constants (`_HDR_FONT`, `_HDR_FILL`, `_STRIPE_FILL`, `_THIN_BORDER`, `_DATA_FONT`)
- Moved inline `logging`/`tkinter` imports to top-level
- Renamed `compute_by_exam` to `_compute_by_exam` (internal only)

### Build
- **Removed `--exclude-module PIL/Pillow`** from CI workflow - Pillow is now a real runtime dependency (tray icon, notification icon extraction)
- **Added `--hidden-import pystray._win32`** and `--collect-all pystray` - pystray loads its Windows backend dynamically; PyInstaller missed it without explicit hints
- **Added `--hidden-import winotify`** - same dynamic loading issue
- **Added Pillow codec pruning step** (`build-release.yml`) - removes `_avif`, `_webp`, `_imagingcms`, `_imagingft`, `_imagingtk` PYD files post-build. Only PNG/RGBA needed for tray and notification icons. Saves ~9 MB unzipped

---

## [v1.2.1] - 2026-04-08

### Fixed
- **Log file never rotates between runs** (`logging_utils.py`) - `setup_logger` now closes and removes the old `_FlushHandler` before adding a new one, so each automation run writes to a fresh log file.
- **Stale dropdown after removing account** (`gui_tk.py`) - `_refresh_account_data` now resets the dropdown selection to "All Accounts" if the currently selected account was removed.
- **Chrome stays running on app close** (`gui_tk.py`, `main.py`) - `EvolveAutomation` exposes the active Selenium driver via `self._driver`; `_on_close` calls `driver.quit()` when the user confirms exit during a running automation, preventing orphaned Chrome processes.
- **`remove_credential` return value ignored** (`gui_tk.py`) - `_bind_settings_remove` now checks the boolean return from `remove_credential` and shows "not found" if removal failed, instead of always reporting success.
- **`add_credential` shows wrong error** (`gui_tk.py`) - when `add_credential` returns `False`, the GUI now distinguishes "already exists" from a generic save failure by checking the credential list.
- **Dead `json.JSONDecodeError` catch** (`secure_credentials.py`) - `add_credential` caught `JSONDecodeError` which `decrypt_credentials` never raises (it wraps all errors as `ValueError`). Replaced with `ValueError`.
- **Results label green even with errors** (`gui_tk.py`) - `_on_complete` now shows the results summary in amber when `errors_encountered > 0`, green only when there are zero errors.

### Security
- **Atomic credential file writes** (`secure_credentials.py`) - `_save_credentials` now writes to a temporary file and uses `os.replace()` to atomically swap it in. Prevents credential file corruption if the app crashes mid-write.

### Changed
- **Merged Excel double I/O** (`excel_utils.py`) - `save_results` now applies formatting (autofilter, freeze panes, header styling, column widths) in-memory before a single `wb.save()`, eliminating the separate `autofilter_and_autofit` load-save cycle.
- **Extracted `_build_lock_card` helper** (`gui_tk.py`) - deduplicated the shared card layout (outer frame, centered card, accent stripe) between `_show_lock_screen` and `_show_reset_password`.
- **Extracted `_create_dialog` helper** (`gui_tk.py`) - deduplicated the shared CTkToplevel setup (title, resizable, configure, transient, center, icon, accent stripe) between `_open_settings` and `_open_help`.
- **Extracted `_JS_LAST_CALENDAR` constant** (`selenium_utils.py`) - the repeated `document.querySelectorAll('.dx-calendar'); calendars[calendars.length-1]` JS snippet used in 4 `execute_script` calls is now a single module-level constant.
- **Consistent f-string formatting** (`selenium_utils.py`) - replaced 4 old-style `%` format strings with f-strings.
- **Removed redundant `os.makedirs`** (`gui_tk.py`) - `_save_last_run` no longer calls `os.makedirs` on `BASE_DIR` which always exists.
- **Partial PDF downloads re-downloaded** (`main.py`) - `_download_pdf` now checks `os.path.getsize(save_path) > 0` in addition to `os.path.exists()`, so 0-byte or truncated files from interrupted downloads are re-downloaded instead of treated as valid.

### Build
- **Synced .spec excludes into CI workflow** (`build-release.yml`) - added 6 `--exclude-module` flags for unused stdlib modules (`unittest`, `xmlrpc`, `pydoc`, `ftplib`, `imaplib`, `smtplib`) and `asyncio`/`_asyncio`.
- **Moved Selenium browser cleanup to post-build deletion** (`build-release.yml`) - `firefox`, `edge`, `ie`, `safari` directories are now removed after build instead of via `--exclude-module` (which caused import errors at runtime).
- **Fixed Tcl/Tk pruning paths** (`build-release.yml`) - PyInstaller outputs `_tcl_data`/`_tk_data`, not `tcl`/`tk`. Previous pruning step was a no-op on CI. Corrected directory names so pruning actually executes.
- **Tcl/Tk data pruning** (`build-release.yml`) - CI step now strips message catalogues, timezone data, demos, stock images, and optional packages from the Tcl/Tk bundle.
- **CJK encoding pruning** (`build-release.yml`) - strips 74 unused CJK/legacy encoding files from `_tcl_data/encoding`, keeping only 8 Western encodings (`ascii`, `utf-8`, `cp1250`–`cp1252`, `iso8859-1`, `iso8859-2`, `iso8859-15`). Saves ~1.4 MB unzipped.
- **Build size reduced from 16.2 MB to ~13 MB zipped** (27 MB unzipped) - down from ~31 MB / 16.2 MB in v1.2.0.

---

## [v1.2.0] - 2026-04-08

### Changed
- **Removed `pandas` dependency** - rewrote `excel_utils.py` to use `openpyxl` + plain Python for all Excel operations. All data now flows as `list[dict[str, str]]` instead of DataFrames. Eliminates `pandas` (62 MB) and `numpy` (31 MB) from the build.
- **Removed `requests` dependency** - replaced `requests.get()` in `main.py` with stdlib `urllib.request.urlopen()` for PDF downloads. Eliminates `requests` + 4 transitive deps (2.5 MB).
- **Replaced `cryptography` with `pyaes` + stdlib** - rewrote `secure_credentials.py` to use AES-256-CBC + HMAC-SHA256 via `pyaes` (26 KB) and stdlib `hashlib`/`hmac`. Eliminates `cryptography` (8.6 MB) from the build. **Breaking:** existing `credentials.enc` files must be deleted and accounts re-added.
- **Switched build from `--onefile` to `--onedir`** - releases are now a zip containing a folder with the EXE and `_internal/` directory. Instant startup (no temp extraction), fewer antivirus false positives, smaller download.
- **Excluded unused Selenium binaries** from build - removed macOS/Linux Selenium Manager binaries and devtools protocol definitions (~21 MB of dead weight).
- **Extracted `MATCH_COLS` constant** (`selenium_utils.py`) - column list for row matching moved from inside loop to module-level constant.
- **Extracted `_apply_dialog_icon` helper** (`gui_tk.py`) - deduplicated identical icon setup code from `_open_settings` and `_open_help`.
- **Renamed shadowed variable** (`selenium_utils.py`) - `idx` → `ci` in `parse_results_table` dict comprehension.
- **Removed unnecessary f-string prefix** (`selenium_utils.py:175`).
- **Upgraded encryption from AES-128 to AES-256** - Fernet used AES-128-CBC; new scheme uses AES-256-CBC with separate HMAC-SHA256 authentication and PBKDF2 key derivation (100k iterations).
- **Updated README** - zip-based install instructions, updated dependency list, updated encryption description to AES-256.
- **Fixed release notes formatting** (`RELEASE_NOTES_v1.1.0.md`) - converted indented text to markdown list syntax.
- **Enhanced logging** (`main.py`) - added version + mode header, per-account summaries, and overall run summary with breakdown. Logs flush on every line for crash-safety.
- **GUI branding** (`gui_tk.py`) - login and reset password titles now City & Guilds red; login card spacing tightened; Quick Open buttons all use outlined secondary style; "Reset Everything" button unified to CG_RED; reset page "will NOT be deleted" text changed from green to CG_RED; completion log simplified to single line.
- **Centralized `APP_VER`** - version constant moved to `config.py`, imported by `main.py` and `gui_tk.py` (was hardcoded in two places).
- **Removed dead `DANGER_HOVER` constant** (`gui_tk.py`) - no longer used after "Reset Everything" button unified to standard primary style.
- **Removed unused `idx`** (`selenium_utils.py`) - `enumerate` in `parse_results_table` replaced with plain `for` since index was never used.
- **Updated `.gitignore`** - `dist/` → `dist*/` to match numbered build folders.
- **Updated GitHub Actions workflow** - added `--exclude-module` flags for removed dependencies; added step to strip macOS/Linux Selenium Manager binaries from release.

### Fixed
- **Headless/small-screen mode only scraping ~22 of 50 rows** (`selenium_utils.py`) - Chrome's `--start-maximized` is ignored in headless mode, defaulting to 800×600 viewport. DevExtreme datagrid virtualizes rows outside the viewport (removes them from DOM entirely). Fix: explicit `--window-size=3840,2160` + `--force-device-scale-factor=0.5` for headless mode; non-headless uses `0.5` for screens ≥1080p and `0.24` for smaller screens (e.g. 800×600). Screen resolution detected via Windows API with safe fallback.
- **Excel header styling** (`excel_utils.py`) - added City & Guilds red (`#E30613`) header row with white bold text, frozen top row, and autofilter.
- **GUI login screen** (`gui_tk.py`) - "View saved results only" button now styled in City & Guilds red; "Forgot password" link uses dim grey for intentional de-emphasis.
- **Hash inconsistency between pandas Series and dict** (`excel_utils.py`) - removing pandas eliminates the `"nan"` vs `""` mismatch that could cause duplicate rows when core hash fields were empty. All data is now plain dicts throughout.

### Removed
- **`pandas`** from `requirements.txt` (and `numpy` as transitive dependency)
- **`requests`** from `requirements.txt` (and `urllib3`, `certifi`, `charset_normalizer`, `idna` as transitive dependencies)
- **`cryptography`** from `requirements.txt` (replaced with `pyaes==1.6.1`, 26 KB vs 8.6 MB)

---

## [v1.1.0] - 2026-04-06

### Changed
- **Restructured `main.py`** - broke monolithic `run()` method into 5 focused methods (`run`, `_process_account`, `_scrape_page`, `_process_page_pdfs`, `_download_pdf`)
- **Merged `filter_utils.py` into `selenium_utils.py`** - consolidated all Selenium interactions into a single module; deleted `filter_utils.py`
- **Moved `save_year_to_excel` and `load_all_existing_data` to `excel_utils.py`** - all Excel data management now in one module
- **Centralized platform URLs in `config.py`** - added `RESULTS_URL` and `DOCUMENT_STORE_URL` constants, replaced hardcoded URLs throughout
- **Added `ProcessingStats` dataclass** for cleaner stats tracking across the automation run
- **Removed `pandas` dependency from `main.py`** - replaced DataFrame usage in `_process_page_pdfs` with direct list/dict operations
- **Extracted `COL_INDEX` constant** in `selenium_utils.py` - centralised column-index mapping used by `parse_results_table` and `select_table_row`
- **Simplified `navigate_to_results`** - removed redundant `driver.get()` call; `driver.refresh()` alone is sufficient since the browser is already on the hash-based SPA URL
- **Removed `columns` parameter** from `initialize_excel`, `save_year_to_excel` - uses `COLUMNS` constant directly
- **Standardised all log messages** to Style A (ellipsis for ongoing actions, no trailing punctuation)
- **GUI file helpers** now use `config.py` path functions (`get_excel_file_for_year`, `get_reports_base_for_year`, `get_logs_base_for_year`) instead of manual `os.path.join`
- **README rewritten** for SEO and non-technical audience; technical details collapsed; author link updated to `santonastaso.me`; added Ko-fi support link

### Fixed
- **DataFrame mutation bug** in `save_results` - added `.copy()` to prevent in-place modification of the caller's DataFrame during date formatting
- **PDF regex case sensitivity** - added `re.IGNORECASE` to `extract_pdf_filename_from_html` so uppercase hex UUIDs are matched
- **Missing `return False`** at end of `click_next_page` fallback path - previously returned `None` on failure
- **Glob year-folder filtering** in `_check_excel_locks` - added `re.match(r'^\d{4}$')` check to skip non-year folders (e.g. `__pycache__`)
- **Race condition in `_poll_queue`** - added final drain after automation thread dies to catch "done"/"error" and "log" messages put just before thread death
- **Unclosed workbook** in `autofilter_and_autofit` - added `wb.close()` after save
- **NaN string bug** in `load_all_existing_data` - pandas reads empty cells as `NaN`; added `"nan"` string check and normalises NaN values to `""` when building row dicts
- **Login failure detection** - `login()` now checks for `validation-summary-errors` div after submit and raises with the server error message
- **`driver.quit()` crash** - wrapped in try/except so a failed quit doesn't break the account loop
- **Unchecked `set_date_filter_to_previous_month_start` return** - now logs a warning if the filter may not have been set
- **Connection leak** - added `resp.close()` in finally block for PDF download requests
- **Error context in PDF processing** - logs student name instead of opaque DataFrame index

### Removed
- **`filter_utils.py`** - merged into `selenium_utils.py`
- **`colorama`** reference from README dependencies (was never used)
- **Dead code** - `return True` from `click_candidate_report_button`, unused `self.columns`, unused `pdf_resume_count` assignment
- **Dead GUI state** - `_authenticated`, `_acct_cur_page` (set but never read)
- **Dead constants** - `AUTHOR_HANDLE`, `GITHUB_URL`, `SUCCESS_BG`, `AMBER_HOVER`, `S24`
- **Stale `__pycache__`** - removed `filter_utils.cpython-313.pyc` and `gui.cpython-313.pyc`
- **Cleaned `.gitignore`** - removed entries for deleted internal docs; added `TODO.md`

---

## [v1.0.0] - 2026-04-02

### Added
- **Desktop GUI** built with CustomTkinter - single-page layout, light mode, City & Guilds red branding
- **Lock screen** with AES-256 encrypted master password and read-only mode for returning users
- **Real-time progress bar** and status text during automation
- **Settings dialog** for managing E-volve SecureAssess login accounts (add/remove)
- **Help dialog** with getting started guide, disclaimer, and support links
- **Quick Open buttons** for Excel results, PDF reports, and log folders
- **Multi-account support** with account selector dropdown (run one or all)
- **Excel lock detection** before automation runs (prevents data corruption)
- **High-DPI icon support** via Windows API (`LoadImageW` + `WM_SETICON`)
- **GitHub Actions workflow** for automated `.exe` builds on tag release

### Automation
- Logs into E-volve SecureAssess and navigates to the Results section
- **Date filter** automatically set to 1st of previous month
- **Full pagination** across multi-page result sets with page verification
- **Year-based folder organisation** for Excel, PDF reports, and logs (`YYYY/`)
- **PDF report downloads** with direct link extraction and resume across runs
- **Incremental saves** after each PDF download (crash-safe)
- **Automatic ChromeDriver management** via Selenium Manager

### Security
- AES-256 encrypted credential storage (`credentials.enc`)
- Master password prompted once per session
- No plaintext credentials ever written to disk

---
