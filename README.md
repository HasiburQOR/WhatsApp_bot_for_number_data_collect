# WhatsApp Number Checker -- Web edition

Checks a list of phone numbers (from an Excel file) against WhatsApp Web and
writes back whether each number is on WhatsApp. In full mode it also reads the
display/business name and business links (email, website, Facebook,
Instagram, LinkedIn) from the contact panel.

It drives WhatsApp Web in Chromium through [Playwright](https://playwright.dev/python/),
with human-like pacing: randomized delays, curved randomized mouse movement,
periodic breaks and a session cooldown.

## Files

| File | What it is |
|---|---|
| `checker_web.py` | The program. |
| `launcher_web.py` | Entry point for the exe: logs the environment, checks dependencies, finds the bundled `Browsers` folder, then calls `checker_web.main()`. |
| `WhatsApp_Number_Checker_Web.spec` | PyInstaller build recipe. |
| `build.bat` | One-click build (creates `.venv`, installs `requirements.txt`, builds `dist\WhatsApp_Number_Checker_Web.exe`). |
| `requirements.txt` | Pinned package versions. |
| `faithful_recovery/` | The source as recovered from the original exe, before any fixes. Diff it against `checker_web.py` to see what changed. |

## Running from source

```bash
python -m venv .venv
.venv\Scripts\pip install -r requirements.txt
.venv\Scripts\python launcher_web.py
```

Options (pass after the script name):

```
--mode fast | --mode full (default)   availability only / also name + links
--headless                            invisible browser; QR prints as ASCII
--speed slow | --speed original       slow = more human (default)
--account NAME                        use another saved WhatsApp account
--login / --logout / --list-accounts  account management
path\to\file.xlsx                     explicit Excel file
```

## Speed

`--speed slow` (default) makes everything slower and more human: mouse paths
have twice as many points and ~2x longer pauses, actions pause ~1.8x longer,
4-9 s between numbers, and a 25-60 s break every 12 numbers.
`--speed original` restores the exact timing of the original exe
(1.5-3.5 s gaps, break every 15). Values are in `SLOW_PACING` / `apply_speed()`.

## Portable layout (for the exe)

```
WhatsApp_Number_Checker_Web.exe
Browsers\chromium-1243\...      Playwright's Chromium (not in this repo)
ExcelFiles\your_numbers.xlsx    first row = headers; needs a phone/number column
accounts\                       created automatically (saved logins)
```

**Playwright version matters.** The exe is built with Playwright **1.63.0**,
which expects Chromium revision **1243**. Get the browser with:

```bash
.venv\Scripts\python -m playwright install chromium
```

then copy the `chromium-1243` folder from `%LOCALAPPDATA%\ms-playwright\`
into `Browsers\`. If you change the Playwright version, the revision changes
too.

## How it works (reading order)

1. `main()` -> arguments, startup menu (accounts), then `run()`.
2. `build_context()` -> persistent Chromium profile per account (`accounts\<name>`), so the QR scan is one-time.
3. `process_excel()` -> loops rows, calls `check_number_web()` for each number, saves every 5 rows.
4. `check_number_web()` -> `open_search_box()` -> `paste_digits()` (clipboard paste, verified) -> `results_stable()` -> `match_result()`.
   If the chat-list search cannot see the number (not in contacts), `probe_number_chat()` opens `web.whatsapp.com/send?phone=...` for a definitive answer.
5. `read_open_chat()` -> name and business links from the contact panel (full mode).
6. Human-like input: `human_mouse_move()` (4 randomized curve styles), `micro_break()`, and the `PACING` table.

## History

The original source was lost; it was recovered from the compiled exe's
bytecode and cross-checked function by function (logic, strings and
docstrings; original comments were not recoverable). Fixes since recovery:

* **Wrong-place clicks on the search box.** A hidden element could match the
  search selector first, and the click position was measured once before a
  slow mouse movement, so a layout shift sent the click elsewhere. Now hidden
  matches are skipped, and every click re-measures the element after the mouse
  arrives and checks nothing unrelated covers it. Clicks always use the human
  mouse path.
* **Chrome warning bar** ("unsupported command-line flag"): launched with
  `chromium_sandbox=True` and `--test-type`.
* **Phone number saved as a contact name** for contacts with no name set:
  `is_usable_name()` now rejects bare phone numbers.

* **Every number after the first "not on WhatsApp" failed on some PCs**
  ("paste ... did not land", "typing did not land"): the "isn't on WhatsApp"
  popup stayed open and trapped clicks and keyboard focus. Popups are now
  closed robustly, with a WhatsApp Web reload as the last resort.

Pacing, breaks and cooldowns are unchanged from the original.

## Do not commit

`accounts/` (your WhatsApp login sessions), `ExcelFiles/` / `*.xlsx` (your
data) and `Browsers/` are excluded by `.gitignore`. Check `git status` before
your first commit.
