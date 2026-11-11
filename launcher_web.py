"""
Entry point for the packaged .exe / cross-platform launcher for the
WhatsApp Web edition of the number checker. Works on Windows AND Linux.

(RECOVERED from the compiled exe by reading its bytecode.)

What it does (in order):
  1. Logs the environment to checker_web_error.log (next to the program).
  2. Verifies every dependency imports (so a failure is readable instead of
     a console that flashes and closes).
  3. Makes sure the Playwright Chromium browser exists -- downloads it on
     first run (~150 MB, once per machine).
  4. Hands over to checker_web.main().

Run it from source:
    python launcher_web.py [args]
      args pass straight through to checker_web:
        --mode fast | --mode full (default)
        --headless               (invisible browser; QR prints as ASCII)
        path\\to\\file.xlsx        (explicit Excel file)
"""

import os
import subprocess
import sys
from datetime import datetime


def _base_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def _log_path():
    name = 'checker_web_error.log'
    target = os.path.join(_base_dir(), name)
    try:
        with open(target, 'a'):
            pass
        return target
    except OSError:
        import tempfile
        return os.path.join(tempfile.gettempdir(), name)


LOG = _log_path()


def log(msg=''):
    try:
        print(msg)
    except Exception:
        pass
    try:
        with open(LOG, 'a', encoding='utf-8') as f:
            f.write(msg + '\n')
    except Exception:
        pass


def environment_report():
    import platform
    log('------------------------------------------------------------')
    log(f'Run started : {datetime.now():%Y-%m-%d %H:%M:%S}')
    log(f'OS          : {platform.platform()}')
    log(f'Machine     : {platform.machine()}')
    log(f'Python      : {sys.version}')
    log(f"Frozen      : {getattr(sys, 'frozen', False)}")
    log(f'Exe         : {sys.executable}')
    log('------------------------------------------------------------')


def check_imports():
    ok = True
    for module in ('openpyxl', 'qrcode', 'playwright', 'playwright.sync_api'):
        try:
            __import__(module)
            log(f'  [ok]   {module}')
        except BaseException as e:
            ok = False
            log(f'  [FAIL] {module}  ->  {type(e).__name__}: {e}')
    return ok


def _browser_registry_roots():
    roots = []
    env = os.environ.get('PLAYWRIGHT_BROWSERS_PATH')
    if env and os.path.isdir(env):
        roots.append(env)
    base = _base_dir()
    for name in ('Browsers', 'browsers', 'ms-playwright'):
        d = os.path.join(base, name)
        if os.path.isdir(d) and d not in roots:
            roots.append(d)
    if os.name == 'nt':
        local = os.environ.get('LOCALAPPDATA')
        if local:
            roots.append(os.path.join(local, 'ms-playwright'))
    else:
        roots.append(os.path.join(os.path.expanduser('~'), '.cache', 'ms-playwright'))
    return [r for r in roots if os.path.isdir(r)]


def _find_installed_chromium():
    import glob as _glob
    chrome = 'chrome.exe' if os.name == 'nt' else 'chrome'
    for root in _browser_registry_roots():
        for cr in sorted(_glob.glob(os.path.join(root, 'chromium-*')), reverse=True):
            for sub in ('chrome-win64', 'chrome-win', 'chrome-linux64', 'chrome-linux', ''):
                exe = os.path.join(cr, sub, chrome) if sub else os.path.join(cr, chrome)
                if os.path.isfile(exe):
                    return (exe, root)
    return (None, None)


def ensure_chromium():
    exe, root = _find_installed_chromium()
    if exe:
        os.environ['PLAYWRIGHT_BROWSERS_PATH'] = root
        bundled = os.path.normcase(root).startswith(os.path.normcase(_base_dir()) + os.sep)
        if bundled:
            log(f'Chromium found (bundled with this app -- no download needed): {exe}')
            return True
        log(f'Chromium found: {exe}')
        return True

    try:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as p:
            path = p.chromium.executable_path
            log(f'Playwright expects Chromium at: {path}')
            if path and os.path.isfile(path):
                return True
    except Exception as e:
        log(f'Could not query the Chromium path: {e}')

    log('Chromium not installed yet -- downloading it once (~150 MB)...')
    candidates = []
    if not getattr(sys, 'frozen', False):
        candidates.append([sys.executable, '-m', 'playwright', 'install', 'chromium'])
    exe_name = 'playwright.exe' if os.name == 'nt' else 'playwright'
    meipass = getattr(sys, '_MEIPASS', None)
    if meipass:
        candidates.append([os.path.join(meipass, 'playwright', 'driver', exe_name),
                           'install', 'chromium'])
    try:
        import playwright as _pw
        candidates.append([os.path.join(os.path.dirname(_pw.__file__), 'driver', exe_name),
                           'install', 'chromium'])
    except Exception:
        pass

    for cmd in candidates:
        try:
            if not cmd[0] or not os.path.isfile(cmd[0]):
                log(f'skipping install candidate (missing): {cmd[0]}')
                continue
            log(f"Running: {' '.join(cmd)}")
            result = subprocess.run(cmd)
            if result.returncode == 0:
                exe, root = _find_installed_chromium()
                if root:
                    os.environ['PLAYWRIGHT_BROWSERS_PATH'] = root
                log('Chromium installed.')
                return True
            log(f'install exited with code {result.returncode}')
        except Exception as e:
            log(f'install attempt failed: {e}')

    edge = _find_edge()
    if edge:
        log('Chromium is not available, but Microsoft Edge was found:')
        log(f'  {edge}')
        log('The program will use Microsoft Edge instead -- no download needed.')
        return True
    return False


def _find_edge():
    if os.name != 'nt':
        return None
    for var in ('ProgramFiles(x86)', 'ProgramFiles', 'LOCALAPPDATA'):
        root = os.environ.get(var)
        if not root:
            continue
        exe = os.path.join(root, 'Microsoft', 'Edge', 'Application', 'msedge.exe')
        if os.path.isfile(exe):
            return exe
    return None


def main():
    exit_code = 0
    try:
        environment_report()
        log('Checking dependencies...')
        if not check_imports():
            log('')
            log('*** One or more dependencies failed to load. ***')
            log('The lines marked [FAIL] above are the reason this machine cannot')
            log('run the program. Send this log file to whoever built the exe:')
            log(f'  {LOG}')
            exit_code = 2
        else:
            log('All dependencies loaded.')
            if not ensure_chromium():
                log('No usable browser was found on this PC.')
                log('Fixes, in order of preference:')
                log('  1. Use the portable ZIP build -- it ships Chromium in a')
                log("     'Browsers' folder next to the exe (works offline).")
                log('  2. Connect to the internet once; Chromium (~150 MB) will')
                log('     download automatically on this launch.')
                log('  3. Install Microsoft Edge or Google Chrome -- the program')
                log('     can drive the installed copy automatically.')
                exit_code = 3
            else:
                log('')
                import checker_web
                exit_code = checker_web.main()
    except SystemExit as e:
        exit_code = e.code if isinstance(e.code, int) else 0
    except KeyboardInterrupt:
        log('\nInterrupted.')
        exit_code = 130
    except Exception:
        import traceback
        log('')
        log('--- UNEXPECTED ERROR ---')
        log(traceback.format_exc())
        exit_code = 1
    finally:
        log(f'Exit code: {exit_code}. Full log: {LOG}')
        try:
            input('\nPress Enter to close...')
        except (EOFError, OSError):
            pass
    return exit_code


if __name__ == '__main__':
    sys.exit(main())
