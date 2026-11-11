"""
Cross-platform WhatsApp number checker -- WhatsApp Web edition.

Port of final_v2.py / checker_fast.py from the native Windows WhatsApp
Desktop app (pywinauto driving the REAL mouse/keyboard) to WhatsApp Web in
Chromium, driven through Playwright.

Why this rewrite:
  * Background operation -- Playwright injects every click/keystroke through
    the browser's own input pipeline (Chrome DevTools Protocol), NOT through
    the operating system. Your real mouse and keyboard stay free, the browser
    window doesn't need focus, and with the launch flags below it can even be
    minimized while the bot keeps working.
  * Cross-platform -- identical code on Windows and Linux; on Linux it can
    run headless on a server with no display at all.

Login: one-time QR scan (phone -> Settings -> Linked devices -> Link a
device). The browser profile persists in wa_profile/ next to the program, so
every later run reuses the session -- same "remember me" behaviour as the
desktop app. The web session is a separate linked device, so the desktop
app can stay logged in at the same time (WhatsApp allows 4 linked devices).

The anti-ban pacing of the original tools is ported 1:1 (randomized gaps,
micro-jitter, periodic breaks, session cap + cooldown in full mode), and the
humanised mouse-movement generator was ported too: clicks travel curved,
varied paths via page.mouse.move() instead of teleporting. Numbers are PASTED
into the search box from the clipboard (Ctrl+V) exactly like the desktop bot
-- no per-keystroke timing fingerprint at all; typing with jitter is only the
fallback when no clipboard is available. Each paste is verified against the
field, so several bots can run side by side (one profile each) without a
shared-clipboard race ever checking the wrong number.

Usage:
    python checker_web.py                     # full mode (name + links)
    python checker_web.py --mode fast         # availability only, faster
    python checker_web.py --headless          # invisible window (servers)
    python checker_web.py C:\\path\\file.xlsx  # explicit Excel file
    python checker_web.py --login             # just (re)link a WhatsApp number
    python checker_web.py --account backup1   # use another saved number
    python checker_web.py --logout            # delete the saved session
    python checker_web.py --list-accounts     # show saved accounts
"""

import argparse
import glob
import json
import math
import os
import random
import re
import shutil
import sys
import time
from datetime import datetime

from openpyxl import load_workbook
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

PACING = {
    'full': {
        'delay_min': 1.5, 'delay_max': 3.5,
        'break_every': 15, 'break_min': 12.0, 'break_max': 30.0,
        'session_limit': 500,
        'cooldown_min': 900, 'cooldown_max': 1500,
        'type_min': 0.03, 'type_max': 0.09,
        'settle': 0.5, 'max_wait': 1.5,
    },
    'fast': {
        'delay_min': 0.4, 'delay_max': 1.0,
        'break_every': None, 'session_limit': None,
        'type_min': 0.012, 'type_max': 0.03,
        'settle': 0.2, 'max_wait': 0.7,
    },
}

CHROMIUM_ARGS = [
    '--disable-backgrounding-occluded-windows',
    '--disable-background-timer-throttling',
    '--disable-renderer-backgrounding',
    '--disable-features=CalculateNativeWinOcclusion',
    '--disable-blink-features=AutomationControlled',
    # Chrome shows a yellow "You are using an unsupported command-line flag"
    # bar for some of the flags above (and for Playwright's own defaults).
    # --test-type suppresses that bar; it changes nothing else visible.
    '--test-type',
]


# --- speed profile ----------------------------------------------------------
# 'slow' (default): every pause and every mouse movement takes longer, and the
#                   gaps between numbers / breaks are longer -- more human.
# 'original':       exactly the timing of the original exe.
# apply_speed() sets these at startup (from --speed).
SPEED = 'slow'
ACTION_SCALE = 1.0   # scales humanize(), micro_delay(), micro_break()
MOUSE_SCALE = 1.0    # scales mouse travel and button-press timing
STEP_SCALE = 1.0     # more points per mouse path (slower, smoother curve)

SLOW_PACING = {
    'full': {
        'delay_min': 4.0, 'delay_max': 9.0,          # gap between numbers
        'break_every': 12, 'break_min': 25.0, 'break_max': 60.0,
        'type_min': 0.08, 'type_max': 0.2,            # only if typing fallback
        'settle': 1.0, 'max_wait': 2.5,               # time "reading" results
    },
    'fast': {
        'delay_min': 1.5, 'delay_max': 3.5,
        'type_min': 0.05, 'type_max': 0.12,
        'settle': 0.5, 'max_wait': 1.5,
    },
}


def apply_speed(speed):
    global SPEED, ACTION_SCALE, MOUSE_SCALE, STEP_SCALE
    SPEED = speed
    if speed == 'slow':
        ACTION_SCALE, MOUSE_SCALE, STEP_SCALE = 1.8, 2.2, 2.0
        for mode, overrides in SLOW_PACING.items():
            PACING[mode].update(overrides)
    else:
        ACTION_SCALE, MOUSE_SCALE, STEP_SCALE = 1.0, 1.0, 1.0


def _msleep(low, high):
    """Randomized pause inside mouse movement / clicks, scaled by MOUSE_SCALE."""
    time.sleep(random.uniform(low, high) * MOUSE_SCALE)


def humanize(seconds):
    """Add +/-20-30% jitter so short internal waits aren't perfectly uniform."""
    return seconds * random.uniform(0.8, 1.3) * ACTION_SCALE


def micro_delay(low, high):
    """A small randomized pause between individual actions inside one search."""
    time.sleep(random.uniform(low, high) * ACTION_SCALE)


def micro_break():
    """Short, irregular pause between mouse/keyboard actions (weighted brief)."""
    r = random.random()
    if r < 0.1:
        time.sleep(random.uniform(0.01, 0.05) * ACTION_SCALE)
        return
    if r < 0.8:
        time.sleep(random.uniform(0.06, 0.25) * ACTION_SCALE)
        return
    time.sleep(random.uniform(0.25, 0.6) * ACTION_SCALE)


def interruptible_sleep(seconds, stop_file, chunk=1.0):
    """Sleep `seconds`, checking `stop_file` every `chunk` seconds so a
    multi-minute break can still be cancelled quickly."""
    remaining = seconds
    while remaining > 0:
        if os.path.isfile(stop_file):
            return True
        step = min(chunk, remaining)
        time.sleep(step)
        remaining -= step
    return False


def wait_until(condition_fn, timeout, poll_interval=0.15):
    """Poll condition_fn() until it returns something truthy or timeout
    elapses. Returns whatever condition_fn() returned (or None)."""
    start = time.time()
    while True:
        result = condition_fn()
        if result:
            return result
        if time.time() - start >= timeout:
            return None
        time.sleep(poll_interval)


# --- humanised mouse movement ---------------------------------------------
_MOUSE_POS = [400.0, 300.0]


def _bezier_path(start, target, ctrl, ease_fn, steps):
    sx, sy = start
    tx, ty = target
    cx, cy = ctrl
    # slow speed: more points along the same curve -> slower, smoother travel
    steps = max(1, int(round(steps * STEP_SCALE)))
    for i in range(1, steps + 1):
        t = ease_fn(i / steps)
        x = (1 - t) ** 2 * sx + 2 * (1 - t) * t * cx + t ** 2 * tx
        y = (1 - t) ** 2 * sy + 2 * (1 - t) * t * cy + t ** 2 * ty
        yield (x, y)


def human_mouse_move(page, target_x, target_y):
    """Move the (virtual) mouse to (target_x, target_y) the way a person's
    hand does -- never the same way twice. Each call randomly picks a
    movement style (smooth curve, quick near-straight dart, curve with
    mid-path hesitation, or slight overshoot-then-correct) so the motion
    isn't one fixed, fingerprintable pattern repeated on every click."""
    start_x, start_y = _MOUSE_POS
    target_x, target_y = float(target_x), float(target_y)
    distance = math.hypot(target_x - start_x, target_y - start_y)
    if distance < 3:
        _MOUSE_POS[0], _MOUSE_POS[1] = target_x, target_y
        return

    perp_x, perp_y = -(target_y - start_y), (target_x - start_x)
    perp_len = math.hypot(perp_x, perp_y) or 1
    mid_x = (start_x + target_x) / 2
    mid_y = (start_y + target_y) / 2

    def bow_ctrl(bow_frac):
        bow = distance * bow_frac * random.choice([-1, 1])
        return (mid_x + perp_x / perp_len * bow, mid_y + perp_y / perp_len * bow)

    style = random.choices(['smooth', 'direct', 'hesitant', 'overshoot'],
                           weights=[40, 25, 20, 15])[0]
    try:
        if style == 'direct':
            steps = max(4, min(12, int(distance / 28)))
            ctrl = bow_ctrl(random.uniform(0.02, 0.06))
            ease = lambda t: 1 - (1 - t) ** 2
            for x, y in _bezier_path((start_x, start_y), (target_x, target_y), ctrl, ease, steps):
                page.mouse.move(int(x), int(y))
                _msleep(0.003, 0.01)
        elif style == 'hesitant':
            steps = max(8, min(24, int(distance / 15)))
            ctrl = bow_ctrl(random.uniform(0.1, 0.22))
            ease = lambda t: 1 - (1 - t) ** 2
            pause_at = random.randint(int(steps * 0.35), int(steps * 0.65))
            for i, (x, y) in enumerate(
                    _bezier_path((start_x, start_y), (target_x, target_y), ctrl, ease, steps), start=1):
                page.mouse.move(int(x), int(y))
                _msleep(0.005, 0.014)
                if i == pause_at:
                    _msleep(0.08, 0.25)
        elif style == 'overshoot':
            steps = max(7, min(18, int(distance / 20)))
            ctrl = bow_ctrl(random.uniform(0.06, 0.16))
            overshoot_frac = random.uniform(1.05, 1.12)
            ox = start_x + (target_x - start_x) * overshoot_frac
            oy = start_y + (target_y - start_y) * overshoot_frac
            ease = lambda t: 1 - (1 - t) ** 2
            for x, y in _bezier_path((start_x, start_y), (ox, oy), ctrl, ease, steps):
                page.mouse.move(int(x), int(y))
                _msleep(0.004, 0.012)
            _msleep(0.03, 0.09)
            for x, y in _bezier_path((ox, oy), (target_x, target_y),
                                     ((ox + target_x) / 2, (oy + target_y) / 2),
                                     lambda t: t, max(3, steps // 4)):
                page.mouse.move(int(x), int(y))
                _msleep(0.004, 0.01)
        else:  # smooth
            steps = max(6, min(22, int(distance / 18)))
            ctrl = bow_ctrl(random.uniform(0.05, 0.18))
            ease = lambda t: 1 - (1 - t) ** 2
            for x, y in _bezier_path((start_x, start_y), (target_x, target_y), ctrl, ease, steps):
                page.mouse.move(int(x), int(y))
                _msleep(0.004, 0.016)
    except Exception:
        pass
    try:
        page.mouse.move(int(target_x), int(target_y))
    except Exception:
        pass
    _MOUSE_POS[0], _MOUSE_POS[1] = target_x, target_y


def human_click(page, x, y):
    """Move to (x, y) on a human-ish path, then click-and-release with a
    natural press duration. All of it is trusted input inside Chromium."""
    human_mouse_move(page, x, y)
    _msleep(0.04, 0.18)
    page.mouse.down()
    _msleep(0.045, 0.1)
    page.mouse.up()


def element_center(element):
    """Viewport-relative center of an element, or None if it has no box."""
    box = element.bounding_box()
    if not box:
        return None
    return (box['x'] + box['width'] / 2, box['y'] + box['height'] / 2)


def _is_hit(element, x, y):
    """True when the topmost element at viewport point (x, y) is `element`
    (or inside / around it) -- i.e. a click there would really land on it."""
    # WhatsApp's search box sits under a decorative icon/button inside the same
    # search container, so "topmost element" can legitimately be a sibling
    # button. Treat anything inside the element's own few-levels-up container
    # as "the same control"; only a genuinely unrelated overlay (modal, popup)
    # counts as covered.
    try:
        return bool(element.evaluate(
            "(el, p) => { const t = document.elementFromPoint(p[0], p[1]); "
            "if (!t) return false; "
            "if (t === el || t.contains(el)) return true; "
            # the chat header's title sits in layered spans/buttons: any point
            # inside the same <header> is a click on the header
            "const h = el.closest('header'); if (h && h.contains(t)) return true; "
            "let a = el; "
            "for (let i = 0; i < 4 && a; i++, a = a.parentElement) { "
            "if (a.contains(t)) return true; } return false; }",
            [x, y]))
    except Exception:
        return True  # cannot tell -- do not block the click


def _is_focused(element):
    try:
        return bool(element.evaluate(
            "el => el === document.activeElement || el.contains(document.activeElement)"))
    except Exception:
        return False


def human_click_element(page, element, dbg=None, what='element'):
    """Click an element along a human path.

    FIX (the "clicks the wrong place" bug): the old version measured the
    element ONCE and then travelled to those coordinates over 0.1-0.5 s. If
    WhatsApp re-laid-out meanwhile (chat list re-rendering, a banner or panel
    sliding in, window resize) the click landed on whatever had moved there.
    Now the element is re-measured after the mouse arrives; if it moved we
    retarget, and it checks (after 3 tries) that nothing unrelated covers the
    click point. The press is ALWAYS made with the human mouse path; the only
    use of Playwright's instant element.click() is when the element has no
    bounding box at all."""
    log_fn = dbg if dbg else (lambda m: None)
    now = None
    for attempt in range(3):
        center = element_center(element)
        if center is None:
            break
        human_mouse_move(page, center[0], center[1])
        now = element_center(element)
        if now is None:
            break
        if math.hypot(now[0] - center[0], now[1] - center[1]) > 3:
            log_fn(f'{what} moved while the mouse travelled -- retargeting ({attempt + 1}/3)')
            continue
        if not _is_hit(element, now[0], now[1]):
            log_fn(f'{what} is covered at the click point -- waiting and retrying ({attempt + 1}/3)')
            time.sleep(0.3)
            continue
        break
    else:
        now = element_center(element)

    if now is None:
        # Only when the element has no box at all (detached / hidden).
        log_fn(f'no bounding box for {what} -- using DOM click fallback')
        element.click()
        return
    # Always finish with a human-paced press at the element's current centre
    # (never Playwright's instant click, which has no mouse movement at all).
    human_mouse_move(page, now[0], now[1])
    _msleep(0.04, 0.18)
    page.mouse.down()
    _msleep(0.045, 0.1)
    page.mouse.up()


SEARCH_BOX_SELECTORS = [
    'input[data-tab="3"]',
    'div[data-testid="chat-list-search-container"] input',
    'input[aria-label*="Search"]',
    'input[placeholder*="Search"]',
    'div[contenteditable="true"][data-tab="3"]',
    '#side div[contenteditable="true"]',
    'div[contenteditable="true"][title*="earch"]',
]

RESULT_ITEM_SELECTORS = [
    '#pane-side [data-testid^="list-item-"]',
    '#pane-side div[role="row"]',
    '#side div[role="listitem"]',
    '#side div[data-id]',
    '#pane-side div[data-id]',
]

NO_RESULT_TEXTS = [
    'no chats, contacts or messages found',
    'no results found',
    'nothing found',
    'no chats found',
    'no contacts found',
]

LOGGED_IN_SELECTORS = ['#side', 'input[data-tab="3"]', 'div[data-testid="chat-list"]']


def find_first(page, selectors, timeout_ms=0, visible_only=False):
    """First element matching any selector, or None. With visible_only=True,
    hidden matches are skipped (a hidden first match used to shadow the real
    visible search box)."""
    for sel in selectors:
        try:
            if visible_only:
                for cand in page.query_selector_all(sel):
                    if cand.is_visible():
                        return cand
                continue
            if timeout_ms:
                el = page.wait_for_selector(sel, timeout=timeout_ms, state='attached')
            else:
                el = page.query_selector(sel)
            if el:
                return el
        except Exception:
            continue
    return None


def is_logged_in(page):
    return find_first(page, LOGGED_IN_SELECTORS) is not None


def qr_payload(page):
    """The raw QR payload string WhatsApp Web encodes, or None. The QR
    container div's data-ref attribute carries the exact string the QR
    image encodes -- enough to re-render it ourselves."""
    try:
        el = page.query_selector('div[data-ref]')
        if el:
            return el.get_attribute('data-ref')
        return None
    except Exception:
        return None


def print_qr_ascii(payload):
    """Render the QR payload as scannable ASCII art in the terminal (used in
    headless mode, where there is no window to point a phone at)."""
    if not payload:
        return
    try:
        import qrcode
        qr = qrcode.QRCode(border=1)
        qr.add_data(payload)
        qr.print_ascii(invert=True)
    except Exception:
        print("(Install 'qrcode' to draw it:  pip install qrcode)")
        print(f'QR payload (feed it to any QR generator to scan): {payload}')


def ensure_login(page, headless, timeout=300, dbg=print):
    """Block until WhatsApp Web is logged in.

    First run: shows the QR (in the browser window when headed, as ASCII art
    in the terminal when headless) and waits for the phone to scan it. Later
    runs: the persisted profile is already logged in and this returns
    immediately. Also used mid-run as the session-health check."""
    for _ in range(15):
        if is_logged_in(page):
            return True
        time.sleep(1)
    if is_logged_in(page):
        return True

    dbg('')
    dbg('============================================================')
    dbg('Not logged in yet -- WhatsApp Web is showing a QR code.')
    if headless:
        dbg('Running headless: scan the ASCII QR below with your phone')
        dbg('(WhatsApp > Settings > Linked devices > Link a device).')
        dbg('')
        print_qr_ascii(qr_payload(page))
    else:
        dbg('Scan the QR code in the browser window with your phone:')
        dbg('   WhatsApp > Settings > Linked devices > Link a device')
    dbg('The bot starts automatically once the scan completes.')
    dbg('============================================================')
    dbg('')

    start = time.time()
    last_payload = qr_payload(page)
    while time.time() - start < timeout:
        if is_logged_in(page):
            dbg('Login detected -- continuing.')
            return True
        time.sleep(2)
        payload = qr_payload(page)
        if headless and payload and payload != last_payload:
            dbg('(QR refreshed -- scan the new one below)')
            print_qr_ascii(payload)
            last_payload = payload
    return False


DIALOG_SELECTORS = [
    'div[role="dialog"]',
    'div[role="alertdialog"]',
    'div[data-animate-modal-popup="true"]',
    '[data-testid="popup-contents"]',
]
DIALOG_BUTTON_WORDS = ('ok', 'close', 'got it', 'done', 'dismiss')


def _open_dialog(page, must_contain=None):
    for sel in DIALOG_SELECTORS:
        try:
            for el in page.query_selector_all(sel):
                if not el.is_visible():
                    continue
                if must_contain and must_contain not in (el.inner_text() or '').lower():
                    continue
                return el
        except Exception:
            continue
    return None


def dismiss_dialogs(page, dbg, attempts=3):
    """Close any modal popup (e.g. WhatsApp's "The phone number ... isn't on
    WhatsApp" notice). While a popup is open it swallows every click and
    keystroke, so the search box can never receive the number.

    FIX: the original only looked for an exact "OK" button and otherwise sent
    Escape; on WhatsApp builds where that did not close the popup, every
    later number failed ("paste ... did not land", "typing did not land").
    Returns True when no popup is left open."""
    for attempt in range(attempts):
        dlg = _open_dialog(page)
        if not dlg:
            return True
        dbg(f'popup is open -- closing it ({attempt + 1}/{attempts})')
        target = None
        try:
            buttons = [b for b in dlg.query_selector_all('button, div[role="button"]')
                       if b.is_visible()]
            for b in buttons:
                txt = (b.inner_text() or '').strip().lower()
                if txt in DIALOG_BUTTON_WORDS:
                    target = b
                    break
            if target is None and len(buttons) == 1:
                target = buttons[0]
        except Exception:
            target = None
        try:
            if target is not None:
                human_click_element(page, target, dbg, 'popup button')
            else:
                page.keyboard.press('Escape')
        except Exception:
            pass
        time.sleep(humanize(0.5))
    still_open = _open_dialog(page) is not None
    if still_open:
        dbg('popup would not close')
    return not still_open


def describe_focus(page):
    """Short description of the element that currently has keyboard focus
    (for debug lines when typed/pasted text does not reach the search box)."""
    try:
        return page.evaluate(
            "() => { const e = document.activeElement; if (!e) return 'nothing'; "
            "const bits = [e.tagName.toLowerCase()]; "
            "for (const a of ['role', 'aria-label', 'title', 'data-tab']) "
            "{ const v = e.getAttribute(a); if (v) bits.push(a + '=' + JSON.stringify(v.slice(0, 40))); } "
            "if (e.closest('[role=dialog]')) bits.push('(inside a popup)'); "
            "return bits.join(' '); }")
    except Exception as e:
        return f'unknown ({e})'


def open_search_box(page, dbg):
    """Dismiss any open panel, then focus the chat-list search box (clicking
    it twice like the desktop bot did, in case it was collapsed). Returns the
    search box element or None."""
    try:
        page.keyboard.press('Escape')
        time.sleep(humanize(0.3))
        micro_break()
    except Exception as e:
        dbg(f'initial ESC failed: {e}')

    # Recovery ladder, cheapest first. Step 0 is the normal path and is
    # exactly the original behaviour; the later steps only run when the search
    # box did not take focus (e.g. a popup is swallowing clicks/keystrokes).
    box = None
    for step in range(3):
        if step == 1:
            dbg(f'search box not focused (focus is on: {describe_focus(page)}) -- checking for a popup')
            dismiss_dialogs(page, dbg)
        elif step == 2:
            # Last resort: reload the app -- no popup survives a reload.
            dbg('still no focus -- reloading WhatsApp Web')
            try:
                page.goto('https://web.whatsapp.com/', wait_until='domcontentloaded', timeout=60000)
                find_first(page, SEARCH_BOX_SELECTORS, timeout_ms=30000)
            except Exception as e:
                dbg(f'reload failed: {e}')

        box = None
        for attempt in range(8):
            box = find_first(page, SEARCH_BOX_SELECTORS, visible_only=True)
            if box:
                break
            dbg(f'search box not found, retry {attempt + 1}/8')
            time.sleep(0.3)
        if not box:
            continue

        try:
            box.scroll_into_view_if_needed()
        except Exception:
            pass

        # Same rhythm as the original exe: two clicks with a pause between
        # them (in case the box was collapsed). Each click re-measures the
        # box, so a layout shift no longer sends a click to the wrong place.
        micro_break()
        human_click_element(page, box, dbg, 'search box')
        time.sleep(humanize(0.3))
        human_click_element(page, box, dbg, 'search box')
        time.sleep(humanize(0.2))
        if _is_focused(box):
            return box
        if _open_dialog(page):
            # A popup is open: it traps focus, so focusing the box directly
            # would not stick. Close it first (next step of the ladder).
            continue
        try:
            box.focus()
            time.sleep(0.3)  # let any focus trap run before trusting it
            if _is_focused(box):
                dbg('search box focused directly after the clicks missed')
                return box
        except Exception:
            pass
    return box


# --- clipboard (Windows only) ------------------------------------------------
if sys.platform == 'win32':
    import ctypes

CF_UNICODETEXT = 13
GMEM_MOVEABLE = 2


def _clipboard_funcs():
    """user32/kernel32 clipboard functions with restype/argtypes set so 64-bit
    handles are not truncated, or None off-Windows / on failure."""
    if sys.platform != 'win32':
        return None
    try:
        u = ctypes.windll.user32
        k = ctypes.windll.kernel32
        u.OpenClipboard.argtypes = [ctypes.c_void_p]
        u.GetClipboardData.restype = ctypes.c_void_p
        u.GetClipboardData.argtypes = [ctypes.c_uint]
        u.SetClipboardData.restype = ctypes.c_void_p
        u.SetClipboardData.argtypes = [ctypes.c_uint, ctypes.c_void_p]
        k.GlobalAlloc.restype = ctypes.c_void_p
        k.GlobalAlloc.argtypes = [ctypes.c_uint, ctypes.c_size_t]
        k.GlobalLock.restype = ctypes.c_void_p
        k.GlobalLock.argtypes = [ctypes.c_void_p]
        k.GlobalUnlock.argtypes = [ctypes.c_void_p]
        k.GlobalFree.argtypes = [ctypes.c_void_p]
        return u, k
    except Exception:
        return None


def _open_clipboard(user32, attempts=5):
    """The Windows clipboard is a single shared resource -- OpenClipboard()
    routinely fails with 'Access is denied' for a few milliseconds while some
    other process (clipboard managers, RDP, another bot instance...) holds
    the lock. Retry with a short backoff instead of treating that as fatal
    (final_v2.py behaviour)."""
    for attempt in range(attempts):
        if user32.OpenClipboard(None):
            return True
        time.sleep(0.05 * (attempt + 1))
    return False


def get_clipboard_text():
    """Current clipboard text, or None if there isn't any/it fails."""
    fk = _clipboard_funcs()
    if not fk:
        return None
    user32, kernel32 = fk
    try:
        if not _open_clipboard(user32):
            return None
        try:
            if not user32.IsClipboardFormatAvailable(CF_UNICODETEXT):
                return None
            h = user32.GetClipboardData(CF_UNICODETEXT)
            if not h:
                return None
            ptr = kernel32.GlobalLock(ctypes.c_void_p(h))
            if not ptr:
                return None
            try:
                return ctypes.cast(ptr, ctypes.c_wchar_p).value
            finally:
                kernel32.GlobalUnlock(ctypes.c_void_p(h))
        finally:
            user32.CloseClipboard()
    except Exception:
        return None


def set_clipboard_text(text):
    """Put text on the clipboard; True on success."""
    fk = _clipboard_funcs()
    if not fk:
        return False
    user32, kernel32 = fk
    try:
        if not _open_clipboard(user32):
            return False
        h = None
        try:
            user32.EmptyClipboard()
            n = (len(text) + 1) * ctypes.sizeof(ctypes.c_wchar)
            h = kernel32.GlobalAlloc(GMEM_MOVEABLE, n)
            if not h:
                return False
            ptr = kernel32.GlobalLock(ctypes.c_void_p(h))
            if not ptr:
                return False
            try:
                ctypes.memmove(ptr, ctypes.create_unicode_buffer(text), n)
            finally:
                kernel32.GlobalUnlock(ctypes.c_void_p(h))
            if not user32.SetClipboardData(CF_UNICODETEXT, ctypes.c_void_p(h)):
                kernel32.GlobalFree(ctypes.c_void_p(h))
                h = None
                return False
            h = None  # the clipboard owns the memory now
            return True
        finally:
            user32.CloseClipboard()
            if h:
                kernel32.GlobalFree(ctypes.c_void_p(h))
    except Exception:
        return False


def focused_field_text(page):
    """Text of whatever element currently has focus (the search box we just
    pasted into). Used to VERIFY the paste actually landed."""
    try:
        return page.evaluate(
            "() => { const el = document.activeElement; if (!el) return ''; "
            "const t = (el.value !== undefined && el.value !== null) ? el.value : "
            "(el.innerText || el.textContent || ''); return String(t).trim(); }")
    except Exception:
        return ''


def paste_digits(page, digits, p, dbg):
    """Clear the search field and PASTE the digits, exactly like the desktop
    bot (final_v2.py): clipboard + Ctrl+V -- no keystroke sequence at all.

    Paths tried, most human-faithful first:
      1. system clipboard + Ctrl+V     -> a genuine, trusted paste event
      2. navigator.clipboard + Ctrl+V  -> same, through the browser clipboard
         (needs the tab focused; covers headless / foreground runs)
      3. keyboard.insertText           -> one single insert event, still zero
         per-key timing to fingerprint
      4. typing with per-keystroke jitter -> final_v2.py's own fallback

    Every attempt is verified against the field, so two bots running side by
    side (which share one OS clipboard) can never check the wrong number: a
    clipboard race fails verification and is simply retried."""
    page.keyboard.press('Control+a')
    time.sleep(humanize(0.1))
    page.keyboard.press('Backspace')
    time.sleep(humanize(0.2))
    micro_break()

    previous = get_clipboard_text()
    try:
        for attempt in range(3):
            if not set_clipboard_text(digits):
                break
            micro_break()
            page.keyboard.press('Control+v')
            time.sleep(humanize(0.25))
            if focused_field_text(page) == digits:
                dbg('pasted from system clipboard')
                return True
            dbg(f'paste attempt {attempt + 1} did not land (clipboard race?) -- retrying')
    finally:
        if previous is not None:
            try:
                set_clipboard_text(previous)
            except Exception:
                pass

    try:
        page.evaluate('t => navigator.clipboard.writeText(t)', digits)
        micro_break()
        page.keyboard.press('Control+v')
        time.sleep(humanize(0.25))
        if focused_field_text(page) == digits:
            dbg('pasted via navigator.clipboard')
            return True
    except Exception:
        pass

    try:
        page.keyboard.insert_text(digits)
        time.sleep(humanize(0.25))
        if focused_field_text(page) == digits:
            dbg('inserted as one paste-like event')
            return True
    except Exception:
        pass

    dbg('paste unavailable -- falling back to typing with jitter')
    page.keyboard.press('Control+a')
    page.keyboard.press('Backspace')
    time.sleep(humanize(0.2))
    for ch in digits:
        page.keyboard.type(ch)
        time.sleep(random.uniform(p['type_min'], p['type_max']))
    time.sleep(humanize(0.15))
    if focused_field_text(page) == digits:
        return True
    dbg('typing did not land either -- the search pane cannot be trusted')
    dbg(f'keyboard focus is on: {describe_focus(page)}')
    return False


def collect_result_items(page):
    """(element, text) for every visible search-result row, de-duplicated by
    text. Multiple selectors are tried because WhatsApp reshuffles its row
    markup between releases."""
    items = []
    seen_texts = set()
    try:
        pane = page.query_selector('#pane-side')
        side = page.query_selector('#side')
        scopes = [s for s in (pane, side) if s] or [page]
        for scope in scopes:
            for sel in RESULT_ITEM_SELECTORS:
                try:
                    for el in scope.query_selector_all(sel):
                        try:
                            text = el.inner_text() or ''
                        except Exception:
                            continue
                        if not text:
                            continue
                        if text in seen_texts:
                            continue
                        seen_texts.add(text)
                        items.append((el, text))
                except Exception:
                    continue
    except Exception:
        pass
    return items


def results_stable(page, p):
    """Wait until the search results stop changing: two identical consecutive
    snapshots AFTER a mandatory settle time. Right after typing, the list
    still shows the OLD (unfiltered) chats for a beat, and two back-to-back
    polls of that unchanged old list would look 'stable' -- exiting there
    would search the wrong snapshot (the same trap the desktop bot guarded
    against)."""
    time.sleep(humanize(p['settle']))
    last = None
    hits = 0
    start = time.time()
    while True:
        texts = tuple(t for _, t in collect_result_items(page))
        if texts == last:
            hits += 1
            if hits >= 2:
                return collect_result_items(page)
        else:
            hits = 0
        last = texts
        if time.time() - start >= max(0.2, p['max_wait'] - p['settle']):
            return collect_result_items(page)
        time.sleep(0.1)


EMAIL_RE = re.compile(r'[\w.+-]+@[\w-]+\.[a-z]{2,}', re.I)
URL_RE = re.compile(r'(https?://[^\s]+|www\.[^\s]+|[\w-]+\.[a-z]{2,}(?:/[^\s]*)?)', re.I)


def classify_link(text):
    """Sort a piece of profile text (or a raw href) into a link category, if
    it is one. Ported verbatim from final_v2.py."""
    t = (text or '').strip()
    if not t:
        return (None, None)
    low = t.lower()
    if 'facebook.com' in low or 'fb.com' in low:
        return ('facebook', t)
    if 'instagram.com' in low:
        return ('instagram', t)
    if 'linkedin.com' in low:
        return ('linkedin', t)
    m = EMAIL_RE.search(t)
    if m:
        return ('email', m.group(0))
    if URL_RE.search(t):
        return ('website', t)
    return (None, None)


PANEL_BLACKLIST = {
    'about', 'groups in common', 'open', 'select messages', 'other business',
    'starred messages', 'contact info', 'search', 'chats', 'report', 'available',
    'delete chat', 'business tools', 'block', 'media, links and docs',
    'about and phone number',
}


_PHONE_LIKE_RE = re.compile(r'^[\s+\-()\d.~]+$')


def is_usable_name(candidate):
    c = (candidate or '').strip()
    # FIX: a bare phone number (e.g. the logged-in account's own number, which
    # the contact panel also lists) is not a name -- it was being written into
    # the Name column for contacts that have no name set.
    if _PHONE_LIKE_RE.match(c) and len(re.sub(r'\D', '', c)) >= 6:
        return False
    return (bool(c) and c.lower() not in PANEL_BLACKLIST
            and 'group' not in c.lower() and 'hour' not in c.lower())


RESULT_PLACEHOLDERS = {'messages', 'communities', 'pinned chat', 'channels', 'chats', 'contacts'}


def match_result(items, digits, dbg):
    """Find the result row for `digits`. Same two-stage heuristic as the
    desktop bot: match by digits anywhere in the row text first; if nothing
    matched, accept any non-placeholder row (verified business rows show
    only the business name -- and with a pure-digit query WhatsApp's own
    search already filtered the list to real matches)."""
    for el, text in items:
        item_digits = re.sub(r'[^\d]', '', text)
        if not digits or digits not in item_digits:
            continue
        dbg(f'matched row by digits: {text[:60]!r}')
        return (el, text)
    for el, text in items:
        low = text.strip().lower()
        if not low or low in RESULT_PLACEHOLDERS:
            continue
        if any(n in low for n in NO_RESULT_TEXTS):
            continue
        dbg(f'matched row by fallback (business name only): {text[:60]!r}')
        return (el, text)
    return (None, None)


def probe_number_chat(page, digits, dbg, timeout=20.0):
    """Definitive availability check for numbers the search box cannot see.

    WhatsApp Web's chat-list search only covers chats, contacts and messages,
    so a REGISTERED number that is not in the account's contacts shows no row
    at all -- indistinguishable from a number that isn't on WhatsApp. (The
    desktop app's search box did a server-side number lookup, which is why
    final_v2.py could trust a miss; the web search cannot.) WhatsApp's own
    send page gives a guaranteed verdict inside the same logged-in session:

        web.whatsapp.com/send?phone=<digits>
          chat opens (header + message bar)    -> number IS on WhatsApp
          'The number ... isn't on WhatsApp.'  -> number is NOT

    Runs IN the main page by navigating to that URL. A second app tab is NOT
    used on purpose: WhatsApp Web keeps one active tab per session and would
    downgrade the main tab to a 'WhatsApp Web is open in another tab' screen,
    which the session check would misread as a logout. The full app loads as
    part of this navigation and stays loaded afterwards, so the next number's
    search flow starts from a normal, logged-in app.

    Returns (verdict, name):
      verdict True  -- registered; the chat is left OPEN in the page;
                       `name` is the chat title when readable (None when the
                       title is just the number itself).
      verdict False -- WhatsApp itself said the number isn't on WhatsApp;
                       the notice dialog is dismissed before returning.
      verdict None  -- inconclusive (timeout / network); caller decides.
    """
    try:
        page.goto('https://web.whatsapp.com/send?phone=' + digits,
                  wait_until='domcontentloaded', timeout=60000)
        deadline = time.time() + timeout
        while time.time() < deadline:
            body = ''
            try:
                body = page.evaluate("() => document.body ? document.body.innerText : ''")
            except Exception:
                continue
            low = body.lower()
            if ("isn't on whatsapp" in low or 'is not on whatsapp' in low
                    or 'shared via url is invalid' in low):
                dbg("probe: WhatsApp says this number isn't on WhatsApp")
                try:
                    ok = find_first(page, ('div[role="button"]:has-text("OK")',
                                           'button:has-text("OK")'), visible_only=True)
                    if ok:
                        human_click_element(page, ok, dbg, 'OK button')
                    else:
                        page.keyboard.press('Escape')
                    time.sleep(humanize(0.4))
                except Exception:
                    pass
                # make sure the notice is really gone before the next number
                time.sleep(0.3)
                if _open_dialog(page, must_contain='whatsapp'):
                    dismiss_dialogs(page, dbg)
                return (False, None)
            title = None
            for sel in ('#main header span[dir="auto"]', '#main header [title]'):
                try:
                    el = page.query_selector(sel)
                    if not el:
                        continue
                    title = (el.inner_text() or '').strip() or None
                except Exception:
                    title = None
                if title:
                    break
            if title and page.query_selector('#main footer'):
                name = None
                if is_usable_name(title) and re.sub(r'\D', '', title) != digits:
                    name = title
                dbg(f'probe: chat opened -- number is on WhatsApp (title {title!r})')
                return (True, name)
            time.sleep(0.5)
        dbg('probe: no verdict within timeout -- inconclusive')
        return (None, None)
    except Exception as e:
        dbg(f'probe failed: {e}')
        return (None, None)


def read_open_chat(page, digits, dbg, suggest=None):
    """The chat for `digits` is (about to be) open in `page`: read the display
    name and harvest business links from the contact-info panel -- the same
    pass the desktop bot did after clicking a search result. `suggest` is a
    name read earlier (e.g. the probe tab's chat title), used when the header
    read fails. Returns (name, links_dict); never raises, and never flips an
    already-made found verdict."""
    panel_links = {'email': None, 'website': None, 'facebook': None,
                   'instagram': None, 'linkedin': None}
    name = None
    try:
        header = wait_until(
            lambda: page.query_selector('#main header')
            or page.query_selector('header[data-testid="conversation-header"]'),
            timeout=6.0)
        if not header:
            dbg('chat did not open (no header) -- keeping the verdict anyway')
            return (suggest, panel_links)
        micro_delay(0.2, 0.5)

        header_title = None
        for sel in ('#main header span[dir="auto"]',
                    'header[data-testid="conversation-header"] span[dir="auto"]',
                    '#main header [title]',
                    'header span[title]'):
            try:
                header_title = page.query_selector(sel)
                if header_title:
                    break
            except Exception:
                continue

        if header_title:
            try:
                candidate = (header_title.inner_text() or '').strip()
                if is_usable_name(candidate) and re.sub(r'[^\d]', '', candidate) != digits:
                    name = candidate
            except Exception:
                pass
        if name is None:
            name = suggest

        try:
            click_target = header_title if header_title else header
            human_click_element(page, click_target, dbg, 'chat header')
            time.sleep(humanize(0.7))

            candidates = []
            try:
                for a in page.query_selector_all('a[href]'):
                    try:
                        candidates.append(a.get_attribute('href'))
                    except Exception:
                        continue
            except Exception:
                pass
            try:
                for span in page.query_selector_all('span[dir="auto"]'):
                    try:
                        txt = span.inner_text()
                    except Exception:
                        continue
                    if not txt or not txt.strip():
                        continue
                    candidates.append(txt)
            except Exception:
                pass

            if name is None:
                flat = []
                for span in page.query_selector_all('span[dir="auto"]'):
                    try:
                        t = (span.inner_text() or '').strip()
                    except Exception:
                        continue
                    if not t:
                        continue
                    flat.append(t)
                for i, t in enumerate(flat):
                    if re.sub(r'[^\d]', '', t) != digits:
                        continue
                    if i + 1 < len(flat) and flat[i + 1].startswith('~'):
                        cand = flat[i + 1].lstrip('~').strip()
                        if is_usable_name(cand):
                            name = cand
                    if name is None and i > 0 and is_usable_name(flat[i - 1]):
                        name = flat[i - 1]
                    break

            for txt in candidates:
                category, value = classify_link(txt)
                if not category:
                    continue
                if panel_links[category]:
                    continue
                panel_links[category] = value
                dbg(f'found {category}: {value!r}')
        finally:
            try:
                page.keyboard.press('Escape')
                time.sleep(humanize(0.2))
                page.keyboard.press('Escape')
                time.sleep(humanize(0.2))
            except Exception:
                pass
        return (name, panel_links)
    except Exception as e:
        dbg(f'profile read failed: {e}')
        return (name or suggest, panel_links)


def check_number_web(page, number, mode='full', debug=True):
    """Search `number` on WhatsApp Web and report availability; in 'full'
    mode also open the chat + contact panel to read the name and business
    links.

    Returns (found, name, links):
      found -- True if a matching result row showed up.
      name  -- display/business name when one could be read, else None.
      links -- dict with email/website/facebook/instagram/linkedin when found
               and mode == 'full'; {} otherwise.

    Never raises -- an error reads as not-found with a debug line (the same
    contract as the desktop version)."""
    def dbg(msg):
        if debug:
            print(f'      [debug] {msg}')

    p = PACING[mode]
    try:
        if not open_search_box(page, dbg):
            dbg('could not open the search box')
            return (False, None, {})

        digits = re.sub(r'[^\d]', '', number)
        pre_texts = tuple(t for _, t in collect_result_items(page))
        if not paste_digits(page, digits, p, dbg):
            dbg('search text never landed in the box')
        items = results_stable(page, p)
        dbg(f'{len(items)} result rows after search')

        search_miss = False
        row = None

        field_now = focused_field_text(page)
        if re.sub(r'[^\d]', '', field_now or '') != digits:
            dbg('search field does not hold this number -- using the probe')
            search_miss = True
        if not search_miss and page.query_selector(
                'div[data-testid="search-no-chats-or-contacts-container"]'):
            dbg("pane shows the 'no results' notice")
            search_miss = True
        if not search_miss:
            side_text = ''
            side = page.query_selector('#side')
            try:
                if side:
                    side_text = (side.inner_text() or '').lower()
            except Exception:
                pass
            if any(n in side_text for n in NO_RESULT_TEXTS):
                dbg("pane shows a 'no results' message")
                search_miss = True
        if not search_miss:
            if pre_texts and tuple(t for _, t in items) == pre_texts:
                dbg('result list unchanged from before the search -- the query never filtered; using the probe')
                search_miss = True
        if not search_miss:
            row, row_text = match_result(items, digits, dbg)
            if row is None:
                dbg('no matching row in search results')
                search_miss = True

        if search_miss:
            try:
                page.keyboard.press('Escape')
                time.sleep(humanize(0.3))
            except Exception:
                pass
            verdict, probe_name = probe_number_chat(page, digits, dbg)
            if verdict is True:
                dbg('registered, but invisible to search -- verified via chat probe')
                if mode == 'fast':
                    return (True, None, {})
                name, links = read_open_chat(page, digits, dbg, suggest=probe_name)
                return (True, name, links)
            if verdict is False:
                return (False, None, {})
            dbg('probe inconclusive -- reporting NO')
            return (False, None, {})

        if mode == 'fast':
            return (True, None, {})

        micro_delay(0.15, 0.45)
        human_click_element(page, row, dbg, 'result row')
        name, links = read_open_chat(page, digits, dbg)
        return (True, name, links)
    except Exception as e:
        dbg(f'EXCEPTION: {e}')
        return (False, None, {})


def find_column(ws, keywords):
    """Find a header column (row 1) whose text contains any of `keywords`."""
    for cell in ws[1]:
        value = str(cell.value).strip().lower() if cell.value is not None else ''
        if any(k in value for k in keywords):
            return cell.column
    return None


def process_excel(page, filepath, mode='full', headless=False):
    p = PACING[mode]
    print(f'\nOpening: {filepath}')
    wb = load_workbook(filepath)
    ws = wb.active

    phone_col = find_column(ws, ['phone', 'number', 'mobile'])
    if phone_col is None:
        print('ERROR: could not find a phone/number column in the header row (row 1). Aborting.')
        return

    remarks_col = find_column(ws, ['remark'])
    if remarks_col is None:
        remarks_col = ws.max_column + 1
        ws.cell(row=1, column=remarks_col, value='Remarks')

    name_col = None
    link_cols = {}
    if mode == 'full':
        name_col = find_column(ws, ['name'])
        if name_col is None:
            name_col = ws.max_column + 1
            ws.cell(row=1, column=name_col, value='Name')
        for key, header, keywords in (
                ('email', 'E-mail', ['email', 'e-mail']),
                ('facebook', 'Facebook', ['facebook']),
                ('instagram', 'Instagram', ['instagram']),
                ('linkedin', 'LinkedIn', ['linkedin']),
                ('website', 'Website', ['website'])):
            col = find_column(ws, keywords)
            if col is None:
                col = ws.max_column + 1
                ws.cell(row=1, column=col, value=header)
            link_cols[key] = col
        print(f'Phone column: {phone_col}, Remarks column: {remarks_col}, Name column: {name_col}')
        print(f'Link columns: {link_cols}')
    else:
        print(f'Phone column: {phone_col}, Remarks column: {remarks_col}')

    stop_file = os.path.join(app_dir(), 'STOP.txt')
    try:
        if os.path.isfile(stop_file):
            os.remove(stop_file)
    except Exception:
        pass

    print('\nTo stop early: create a file named STOP.txt in the program folder')
    print(f'  ({stop_file})')
    print("The bot finishes the number it's on, saves, and exits cleanly. Ctrl+C also works.\n")
    pacing_bits = f"{p['delay_min']:.1f}-{p['delay_max']:.1f}s random gap between numbers"
    if p.get('break_every'):
        pacing_bits += (
            f", a longer break every {p['break_every']} checks, and a "
            f"{p['cooldown_min'] / 60:.0f}-{p['cooldown_max'] / 60:.0f} min cooldown every "
            f"{p['session_limit']} checks (auto-resumes -- safe to leave overnight)")
    print(f'Pacing: {pacing_bits}. This pacing is deliberate, to avoid looking like bulk automation.\n')

    total = 0
    checked = 0
    session_checked = 0
    found_count = 0
    SAVE_EVERY = 5
    unsaved_rows = 0
    stopped_early = False

    def flush():
        nonlocal unsaved_rows
        if unsaved_rows:
            wb.save(filepath)
            unsaved_rows = 0

    try:
        for row in range(2, ws.max_row + 1):
            if os.path.isfile(stop_file):
                print('STOP.txt detected -- stopping early.')
                stopped_early = True
                try:
                    os.remove(stop_file)
                except Exception:
                    pass
                break

            if not is_logged_in(page):
                side_back = False
                try:
                    side_back = bool(page.wait_for_selector('#side', timeout=3000))
                except Exception:
                    side_back = False
                if not side_back:
                    print('\nSession lost (QR visible again) -- pausing.')
                    if not ensure_login(page, headless=headless, timeout=600):
                        print('Still logged out after 10 minutes -- saving and exiting.')
                        print('')
                        print('If this number was banned/flagged, switch to a')
                        print('DIFFERENT WhatsApp number:')
                        print('  1. On the phone: WhatsApp > Settings > Linked devices')
                        print('     > unlink this PC (frees a device slot).')
                        print('  2. Start the program again, menu option 2:')
                        print("     'Switch / add account' -> type a new name.")
                        print("  3. Scan the new QR with the OTHER phone's WhatsApp.")
                        print('     Each account keeps its own saved session, so the')
                        print('     old one stays untouched (delete it via option 4).')
                        stopped_early = True
                        break
                    print('Session restored -- continuing.\n')

            raw_number = ws.cell(row=row, column=phone_col).value
            if raw_number is None or not str(raw_number).strip():
                continue
            total += 1
            digits = re.sub(r'[^\d]', '', str(raw_number))
            if not digits:
                continue

            existing_remark = ws.cell(row=row, column=remarks_col).value
            if existing_remark is not None and str(existing_remark).strip():
                print(f'[row {row}] {raw_number} -> already checked ({existing_remark}), skipping')
                continue

            print(f'[row {row}] Checking {raw_number} ({digits})...')
            found, name, links = check_number_web(page, digits, mode)
            checked += 1
            session_checked += 1

            ws.cell(row=row, column=remarks_col, value='YES' if found else 'NO')
            if mode == 'full':
                if name:
                    ws.cell(row=row, column=name_col, value=name)
                for key, col in link_cols.items():
                    value = links.get(key)
                    if value:
                        ws.cell(row=row, column=col, value=value)
            unsaved_rows += 1

            if found:
                found_count += 1
                if mode == 'full':
                    print('  -> YES' + (f'  (name: {name})' if name else '  (no name set)'))
                else:
                    print('  -> YES')
            else:
                print('  -> NO')

            if unsaved_rows >= SAVE_EVERY:
                flush()

            if p.get('session_limit') and session_checked >= p['session_limit']:
                flush()
                cooldown = random.uniform(p['cooldown_min'], p['cooldown_max'])
                resume_at = time.strftime('%H:%M:%S', time.localtime(time.time() + cooldown))
                print(
                    f"\nReached the {p['session_limit']}-number session cap. Cooling down for ~"
                    f"{cooldown / 60:.0f} minutes (resuming around {resume_at}). "
                    f"Safe to leave running; create STOP.txt to cancel the wait.")
                session_checked = 0
                if interruptible_sleep(cooldown, stop_file):
                    continue
                print('Cooldown finished, resuming...\n')
                continue

            gap = random.uniform(p['delay_min'], p['delay_max'])
            if interruptible_sleep(gap, stop_file):
                continue

            if not p.get('break_every') or checked % p['break_every'] != 0:
                continue
            break_len = random.uniform(p['break_min'], p['break_max'])
            print(f'  ...taking a {break_len:.0f}s break (checked {checked} so far)...')
            if interruptible_sleep(break_len, stop_file):
                continue
    finally:
        flush()

    print('==================================================')
    if stopped_early:
        print(f'Stopped early. Checked {checked} numbers before stopping. {found_count} found on WhatsApp.')
    else:
        print(f'Done! Checked {checked} numbers ({total} rows with a number). {found_count} found on WhatsApp.')
    print(f'Results saved to: {filepath}')


# --- accounts -----------------------------------------------------------------
def app_dir():
    """Folder the program lives in -- the .exe's folder when frozen."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def excel_folder():
    folder = os.path.join(app_dir(), 'ExcelFiles')
    try:
        os.makedirs(folder, exist_ok=True)
    except Exception:
        pass
    return folder


LEGACY_PROFILE_DIRNAME = 'wa_profile'
DEFAULT_ACCOUNT = 'Main'


def accounts_root():
    return os.path.join(app_dir(), 'accounts')


def sanitize_account_name(name):
    """Make an account name safe to use as a folder name."""
    cleaned = re.sub(r'[^\w \-]', '_', str(name or '').strip())
    cleaned = re.sub(r'\s+', ' ', cleaned).strip(' ._')
    return cleaned[:40] or DEFAULT_ACCOUNT


def account_meta_path(name):
    return os.path.join(accounts_root(), sanitize_account_name(name) + '.json')


def load_account_meta(name):
    path = account_meta_path(name)
    try:
        if os.path.isfile(path):
            with open(path, 'r', encoding='utf-8') as f:
                meta = json.load(f)
            if isinstance(meta, dict):
                return meta
        return None
    except Exception:
        return None


def save_account_meta(name, **updates):
    """Create/merge the account's <name>.json note."""
    name = sanitize_account_name(name)
    os.makedirs(accounts_root(), exist_ok=True)
    meta = load_account_meta(name) or {
        'name': name,
        'created': datetime.now().isoformat(timespec='seconds'),
    }
    meta.update(updates)
    try:
        with open(account_meta_path(name), 'w', encoding='utf-8') as f:
            json.dump(meta, f, indent=2)
    except OSError as e:
        print(f'WARNING: could not save the account note: {e}')
    return meta


def account_names():
    """Every known account: metadata notes + bare profile folders + the old
    single-profile layout (which counts as the 'Main' account)."""
    names = []
    root = accounts_root()
    if os.path.isdir(root):
        for entry in os.listdir(root):
            full = os.path.join(root, entry)
            if entry.lower().endswith('.json') and os.path.isfile(full):
                names.append(entry[:-5])
            elif os.path.isdir(full) and not entry.startswith('.'):
                names.append(entry)
    legacy = os.path.join(app_dir(), LEGACY_PROFILE_DIRNAME)
    if os.path.isdir(legacy) and DEFAULT_ACCOUNT not in names:
        names.append(DEFAULT_ACCOUNT)
    return sorted(set(names))


def account_profile_dir(name):
    """Browser-profile folder the account's WhatsApp session lives in."""
    name = sanitize_account_name(name)
    meta = load_account_meta(name)
    if meta and meta.get('profile_dir') and os.path.isdir(meta['profile_dir']):
        return meta['profile_dir']
    legacy = os.path.join(app_dir(), LEGACY_PROFILE_DIRNAME)
    if name == DEFAULT_ACCOUNT and os.path.isdir(legacy):
        save_account_meta(name, profile_dir=legacy)
        return legacy
    return os.path.join(accounts_root(), name)


def default_account_name():
    names = account_names()
    if not names:
        return DEFAULT_ACCOUNT
    if DEFAULT_ACCOUNT in names:
        return DEFAULT_ACCOUNT

    def last_used(n):
        meta = load_account_meta(n) or {}
        return meta.get('last_login') or meta.get('created') or ''
    return max(names, key=last_used)


def describe_account(name):
    """One-line summary for menus and logs, e.g. 'Main (+234..., last login ...)'."""
    meta = load_account_meta(sanitize_account_name(name)) or {}
    who = meta.get('phone') or meta.get('last_seen_name')
    bits = [f'logged in as {who}'] if who else []
    if meta.get('last_login'):
        bits.append(f"last login {str(meta['last_login'])[:16].replace('T', ' ')}")
    if not bits:
        return f'{sanitize_account_name(name)} (never logged in)'
    return f"{sanitize_account_name(name)} ({', '.join(bits)})"


def create_account(name):
    name = sanitize_account_name(name)
    if load_account_meta(name):
        print(f"Account '{name}' already exists.")
        return name
    save_account_meta(name)
    print(f"Account '{name}' created.")
    return name


def delete_account(name):
    """Delete an account's saved session (profile folder) and its note."""
    name = sanitize_account_name(name)
    meta = load_account_meta(name) or {}
    profile = meta.get('profile_dir')
    if not profile:
        legacy = os.path.join(app_dir(), LEGACY_PROFILE_DIRNAME)
        if name == DEFAULT_ACCOUNT and os.path.isdir(legacy):
            profile = legacy
    if not profile:
        profile = os.path.join(accounts_root(), name)
    if os.path.isdir(profile):
        shutil.rmtree(profile, ignore_errors=True)
    try:
        os.remove(account_meta_path(name))
    except OSError:
        return


def read_logged_in_identity(page):
    """Best-effort answer to 'which WhatsApp number is this session?' --
    from localStorage's user id, or the profile button's label. Returns
    something like '+2348012345678' or a display name, or None. Purely
    informational: never blocks the run when the UI doesn't cooperate."""
    try:
        raw = page.evaluate(
            "() => { try { return localStorage.getItem('last-wid') || ''; } "
            "catch (e) { return ''; } }")
        m = re.search(r'"?(\d{6,20})', str(raw or ''))
        if m:
            return '+' + m.group(1)
    except Exception:
        pass
    generic = ('profile', 'profil', 'perfil', 'search', 'menu', 'chats', 'chat', 'status',
               'new chat', 'settings', 'community', 'communities', 'starred',
               'linked devices', 'log out', 'keyboard shortcuts', 'notifications', 'default')
    for sel in ('div[data-testid="menu-bar-item_profile"]', 'div[aria-label][role="button"]'):
        try:
            el = page.query_selector(sel)
            label = (el.get_attribute('aria-label') or '').strip() if el else ''
            low = label.lower()
            if label and not any(word in low for word in generic):
                return label
        except Exception:
            continue
    return None


def record_login(name, page=None):
    """Stamp the account note after a successful login."""
    who = read_logged_in_identity(page) if page else None
    updates = {'last_login': datetime.now().isoformat(timespec='seconds')}
    if who:
        if who.startswith('+'):
            updates['phone'] = who
        else:
            updates['last_seen_name'] = who
    save_account_meta(name, **updates)
    return who


def print_accounts():
    names = account_names()
    if not names:
        print("No saved accounts yet -- the first run creates 'Main'.")
        return
    current = default_account_name()
    print('Saved WhatsApp accounts:')
    for n in names:
        marker = '   <- default' if n == current else ''
        print('  - ' + f'{describe_account(n)}{marker}')
    print(f'Account profiles live in: {accounts_root()}')


def _is_interactive():
    try:
        return sys.stdin.isatty() and sys.stdout.isatty()
    except Exception:
        return False


def _menu_pick_or_add_account(current):
    """Menu option 2: list accounts / create a new one by typing its name."""
    names = account_names()
    print('\nSaved accounts:')
    for i, n in enumerate(names, 1):
        marker = '   (current)' if n == current else ''
        print('  ' + f'{i}. {describe_account(n)}{marker}')
    print('  ' + f'{len(names) + 1}. < add a NEW account >')
    try:
        choice = input(
            f"Pick 1-{len(names) + 1}, or type a new account name (Enter = keep '{current}'): ").strip()
    except (EOFError, OSError):
        return current
    if not choice:
        return current
    if choice.isdigit() and 1 <= int(choice) <= len(names):
        picked = names[int(choice) - 1]
        account_profile_dir(picked)
        return picked
    new = create_account(choice)
    if new != current:
        print(f"\n'{new}' starts with an empty WhatsApp session -- a QR code")
        print("will appear on its first start. Scan it with THAT phone's")
        print('WhatsApp (phone > Settings > Linked devices > Link a device).')
    return new


def startup_menu(args):
    """Interactive menu for double-click launches. Returns (action, args)
    with action one of 'run' | 'login' | 'quit'; may set args.account."""
    if not account_names():
        create_account(DEFAULT_ACCOUNT)
    current = sanitize_account_name(args.account or default_account_name())
    while True:
        print('\n====================================================')
        print('  WhatsApp Number Checker -- Web')
        print('====================================================')
        print('  Current account: ' + describe_account(current))
        print()
        print('  1. Start checking now')
        print('  2. Switch / add account (use a second WhatsApp number)')
        print('  3. Log in / relink current account (scan the QR)')
        print("  4. Log out & delete current account's saved session")
        print('  5. Exit')
        try:
            choice = input('Choose [1-5, Enter = start]: ').strip().lower()
        except (EOFError, OSError):
            return ('quit', args)

        if choice in ('', '1', 's', 'start'):
            args.account = current
            return ('run', args)
        if choice == '2':
            current = _menu_pick_or_add_account(current)
            continue
        if choice == '3':
            args.account = current
            return ('login', args)
        if choice == '4':
            print(f"\nThis deletes the saved WhatsApp session of '{current}'.")
            print('(The phone number itself is unaffected -- you are just')
            print(" unlinking this PC's saved login.)")
            try:
                ans = input(f"Really delete '{current}'? [y/N] ").strip().lower()
            except (EOFError, OSError):
                ans = 'n'
            if ans in ('y', 'yes'):
                delete_account(current)
                print(f"Session of '{current}' removed. Tip: also unlink this")
                print('PC on the phone (WhatsApp > Settings > Linked devices).')
                remaining = account_names()
                if remaining:
                    current = remaining[0]
                else:
                    create_account(DEFAULT_ACCOUNT)
                    current = DEFAULT_ACCOUNT
            continue
        if choice in ('5', 'q', 'quit', 'exit'):
            return ('quit', args)
        print('Please answer 1-5.')


def do_login(args):
    """Open the browser, show the QR, wait for the scan, save the session.
    Does not touch any Excel file -- this is the 'just log in' command."""
    name = sanitize_account_name(args.account or default_account_name())
    profile = account_profile_dir(name)
    print(f'Account : {name}')
    print(f'Profile : {profile}')
    print('Starting the browser for the WhatsApp QR login...\n')
    with sync_playwright() as p:
        context = build_context(p, args.headless, profile,
                                browser=getattr(args, 'browser', 'auto'))
        page = context.pages[0] if context.pages else context.new_page()
        page.set_default_timeout(5000)
        try:
            page.goto('https://web.whatsapp.com/', wait_until='domcontentloaded', timeout=60000)
        except PlaywrightTimeoutError:
            print('WARNING: web.whatsapp.com was slow to load -- continuing anyway.')
        ok = ensure_login(page, args.headless, timeout=600)
        who = record_login(name, page) if ok else None
        try:
            context.close()
        except Exception:
            pass
    if not ok:
        print('Login did not complete within 10 minutes. Run again and scan the QR.')
        return 1
    print("\nLogin saved for account '" + name + "'."
          + (f'  Logged in as: {who}' if who else ''))
    print('From now on this account starts without a QR scan.')
    return 0


def do_logout(args):
    """Delete an account's saved session (after a confirmation)."""
    name = sanitize_account_name(args.account or default_account_name())
    print('Account : ' + describe_account(name))
    if not args.yes:
        try:
            ans = input(f"Delete the saved session of '{name}'? [y/N] ").strip().lower()
        except (EOFError, OSError):
            ans = 'n'
        if ans not in ('y', 'yes'):
            print('Cancelled -- nothing was deleted.')
            return 0
    delete_account(name)
    print(f"Session of '{name}' removed. Tip: also unlink this PC on the")
    print('phone (WhatsApp > Settings > Linked devices) to free a device slot.')
    return 0


def pick_excel_file(explicit=None):
    """Resolve which .xlsx to process: CLI arg > single file in ExcelFiles/ >
    numbered menu > pasted path. Mirrors the original picker's behaviour."""
    if explicit:
        for cand in (explicit,
                     os.path.join(app_dir(), explicit),
                     os.path.join(excel_folder(), os.path.basename(explicit))):
            if os.path.isfile(cand):
                return os.path.abspath(cand)
    folder = excel_folder()
    candidates = [f for f in glob.glob(os.path.join(folder, '*.xlsx'))
                  if not os.path.basename(f).startswith('~$')]
    if len(candidates) == 1:
        print(f'Using Excel file: {candidates[0]}')
        return candidates[0]
    if len(candidates) > 1:
        print(f'Multiple .xlsx files found in {folder}:')
        for i, c in enumerate(candidates, 1):
            print('  ' + f'{i}. {os.path.basename(c)}')
        choice = input('Enter the number of the file to use (or paste a full path): ').strip()
        if choice.isdigit() and 1 <= int(choice) <= len(candidates):
            return candidates[int(choice) - 1]
        return choice.strip('"').strip("'")
    print(f'No .xlsx file found in: {folder}')
    print('Put your Excel file in that folder and run this program again,')
    print('or paste the full path to it below.')
    path = input('Full path to the Excel file (or press Enter to quit): ').strip().strip('"').strip("'")
    return path


def _console_safe(text):
    """Make any string printable on legacy cp1252/cp437 Windows consoles --
    Playwright errors contain Unicode box-drawing characters that would
    otherwise crash print() mid-fallback."""
    enc = getattr(sys.stdout, 'encoding', None) or 'ascii'
    try:
        text.encode(enc)
        return text
    except (UnicodeEncodeError, LookupError):
        return text.encode(enc, 'replace').decode(enc, 'replace')


def build_context(p, headless, profile_dir, browser='auto'):
    """Persistent Chromium context. The profile in `profile_dir` is what
    keeps the WhatsApp session alive between runs (one-time QR scan).

    `browser` picks the engine: 'auto' tries Playwright's own Chromium
    first, then falls back to the machine's installed Microsoft Edge /
    Google Chrome -- so the program still starts on a PC where Chromium was
    never downloaded (offline machines, blocked downloads, ...)."""
    if browser in ('msedge', 'chrome'):
        order = [(browser, 'system Microsoft Edge' if browser == 'msedge' else 'system Google Chrome')]
    else:
        order = [(None, 'Chromium')]
        if browser == 'auto':
            order += [('msedge', 'system Microsoft Edge'), ('chrome', 'system Google Chrome')]

    last_error = None
    for channel, label in order:
        kwargs = dict(user_data_dir=profile_dir, headless=headless, args=CHROMIUM_ARGS)
        if channel:
            kwargs['channel'] = channel
        if headless:
            kwargs['viewport'] = {'width': 1280, 'height': 800}
        else:
            kwargs['no_viewport'] = True
        # Playwright passes --no-sandbox by default, which makes Chrome show the
        # "unsupported command-line flag: --no-sandbox" warning bar.
        # chromium_sandbox=True removes the flag; if a PC cannot start with the
        # sandbox, retry once without it rather than failing.
        for sandbox in (True, False):
            try:
                context = p.chromium.launch_persistent_context(
                    chromium_sandbox=sandbox, **kwargs)
                print(f'[browser] using {label}' + ('' if sandbox else ' (no sandbox)'))
                _after_launch(context, headless)
                return context
            except Exception as e:
                last_error = e
                print(f'[browser] {label} could not start'
                      + (' with the sandbox' if sandbox else '') + ': '
                      + _console_safe(' '.join(str(e).split())[:200]))
    raise last_error


def _after_launch(context, headless):
    """Post-launch setup shared by every engine choice."""
    try:
        context.grant_permissions(['clipboard-read', 'clipboard-write'],
                                  origin='https://web.whatsapp.com')
    except Exception:
        pass
    try:
        if headless:
            page = context.pages[0] if context.pages else context.new_page()
            raw_ua = page.evaluate('navigator.userAgent')
            fixed = raw_ua
            if 'HeadlessChrome' in fixed:
                fixed = fixed.replace('HeadlessChrome', 'Chrome')
            if 'HeadlessEdg/' in fixed:
                fixed = fixed.replace('HeadlessEdg/', 'Edg/')
            if fixed != raw_ua:
                cdp = context.new_cdp_session(page)
                cdp.send('Emulation.setUserAgentOverride', {'userAgent': fixed})
                context._ua_fix_cdp = cdp
    except Exception:
        return


def run(args):
    mode = args.mode
    print('\n==================================================')
    print(f"  WhatsApp Number Checker -- Web ({'fast' if mode == 'fast' else 'full'})")
    print(f'  Speed: {SPEED}' + ('  (slower, more human; use --speed original for the old timing)'
                                if SPEED == 'slow' else ''))
    print('==================================================\n')

    filepath = pick_excel_file(args.excel)
    if not filepath or not os.path.isfile(filepath):
        print(f'ERROR: file not found: {filepath!r}')
        print(f'Put your Excel file in: {excel_folder()}')
        return 1

    account = sanitize_account_name(args.account or default_account_name())
    if args.profile:
        profile_dir = args.profile
        print(f'Account  : {account} (custom profile override)')
    else:
        profile_dir = account_profile_dir(account)
        print('Account  : ' + describe_account(account))
    print(f'Profile  : {profile_dir}')
    print('Launching browser...')

    with sync_playwright() as p:
        context = build_context(p, args.headless, profile_dir,
                                browser=getattr(args, 'browser', 'auto'))
        try:
            page = context.pages[0] if context.pages else context.new_page()
            page.set_default_timeout(5000)
            try:
                page.goto('https://web.whatsapp.com/', wait_until='domcontentloaded', timeout=60000)
            except PlaywrightTimeoutError:
                print('WARNING: web.whatsapp.com was slow to load -- continuing anyway.')

            if not ensure_login(page, args.headless, timeout=300):
                print('Login timed out after 5 minutes. Run again and scan the QR.')
                try:
                    context.close()
                except Exception:
                    pass
                return 1
            record_login(account, page)

            print('Waiting for the chat list to finish loading...')
            ready = False
            for _ in range(30):
                if find_first(page, SEARCH_BOX_SELECTORS):
                    ready = True
                    break
                time.sleep(2)
            if not ready:
                print('WARNING: the search box did not appear within 60s (history still syncing?). Continuing anyway...')

            process_excel(page, filepath, mode, headless=args.headless)
        finally:
            try:
                context.close()
            except Exception:
                pass
    return 0


def build_arg_parser():
    ap = argparse.ArgumentParser(
        description='WhatsApp number checker (WhatsApp Web, background-safe, cross-platform).')
    ap.add_argument('excel', nargs='?', default=None,
                    help='path to the .xlsx file (default: pick from ExcelFiles/ folder)')
    ap.add_argument('--mode', choices=['fast', 'full'], default='full',
                    help="'fast' = availability only; 'full' = also name + business links (default)")
    ap.add_argument('--headless', action='store_true',
                    help='run Chromium invisibly (for servers; QR prints as ASCII art)')
    ap.add_argument('--profile', default=None,
                    help="browser profile folder override (default: the current account's folder under accounts\\)")
    ap.add_argument('--account', default=None, metavar='NAME',
                    help='use this saved WhatsApp account/profile (e.g. Main, backup1); a new name starts a fresh QR login')
    ap.add_argument('--browser', choices=['auto', 'chromium', 'msedge', 'chrome'], default='auto',
                    help='browser engine: auto = Chromium (bundled/downloaded), then the installed Microsoft Edge / Google Chrome')
    ap.add_argument('--login', action='store_true',
                    help='only log this account in (show the QR, save the session) -- no Excel checking')
    ap.add_argument('--logout', action='store_true',
                    help='delete the saved session of --account (default account if omitted); add --yes to skip the confirmation')
    ap.add_argument('--list-accounts', action='store_true',
                    help='list the saved WhatsApp accounts and exit')
    ap.add_argument('--yes', action='store_true',
                    help="assume 'yes' at confirmation prompts (e.g. --logout)")
    ap.add_argument('--speed', choices=['slow', 'original'], default='slow',
                    help="'slow' (default) = slower, more human mouse movement and pauses; "
                         "'original' = exact timing of the original exe")
    ap.add_argument('--no-menu', action='store_true',
                    help='skip the interactive startup menu')
    return ap


def main(argv=None):
    args = build_arg_parser().parse_args(argv)
    apply_speed(args.speed)
    if args.list_accounts:
        print_accounts()
        return 0
    if args.logout:
        return do_logout(args)

    action = None
    if not (args.no_menu or args.excel or args.account or args.profile
            or args.login or not _is_interactive()):
        action, args = startup_menu(args)
    if action == 'quit':
        print('Bye.')
        return 0
    if action == 'login' or args.login:
        return do_login(args)
    return run(args)


if __name__ == '__main__':
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print('\nInterrupted.')
    except Exception:
        import traceback
        print('')
        print('--- UNEXPECTED ERROR ---')
        traceback.print_exc()
    finally:
        print('')
        try:
            input('Press Enter to close...')
        except (EOFError, OSError):
            pass
