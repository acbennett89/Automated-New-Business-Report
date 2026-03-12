from __future__ import annotations

import json
import os
import re
import subprocess
import webbrowser
from datetime import datetime
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR.parent if SCRIPT_DIR.name.casefold() == "scripts" else SCRIPT_DIR
WORKING_FILES_DIR = PROJECT_ROOT / "Working Files"

LOGIN_URL = "https://bignition.thewedge.net/Account/Login?ReturnUrl=%2f"
DASHBOARD_URL = "https://bignition.thewedge.net/Home/Dashboard"
AGENCY_REPORTS_URL = "https://bignition.thewedge.net/Reports/AgencyReports"
OPPS_BY_PRODUCER_URL = "https://bignition.thewedge.net/Reports/ViewAgencyReport?id=5"
STORAGE_STATE_PATH = PROJECT_ROOT / "config" / "storage_state.json"
BIGNITION_CREDENTIALS_PATH = PROJECT_ROOT / "config" / "bignition_credentials.json"


def log_step(message: str) -> None:
    print(f"[{datetime.now():%H:%M:%S}] {message}")


def find_chrome() -> Path | None:
    candidates = [
        Path(os.environ.get("PROGRAMFILES", "")) / "Google/Chrome/Application/chrome.exe",
        Path(os.environ.get("PROGRAMFILES(X86)", "")) / "Google/Chrome/Application/chrome.exe",
        Path(os.environ.get("LOCALAPPDATA", "")) / "Google/Chrome/Application/chrome.exe",
    ]
    for path in candidates:
        if path.is_file():
            return path
    return None


def launch_browser(p, headless: bool):
    log_step(f"Launching {'headless' if headless else 'headed'} Chromium for Bignition.")
    if headless:
        return p.chromium.launch(headless=True)
    try:
        return p.chromium.launch(channel="chrome", headless=False)
    except Exception:
        return p.chromium.launch(headless=False)


def load_bignition_credentials(path: Path) -> tuple[str, str] | None:
    if not path.is_file():
        return None
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return None

    username = str(raw.get("username", "")).strip()
    password = str(raw.get("password", "")).strip()
    if not username or not password:
        return None
    return username, password


def _first_visible_locator(page, selectors: list[str]):
    for selector in selectors:
        try:
            loc = page.locator(selector)
            if loc.count() > 0 and loc.first.is_visible():
                return loc.first
        except Exception:
            continue
    return None


def try_submit_bignition_login(page, username: str, password: str) -> bool:
    username_selectors = [
        'input#UserName',
        'input#Username',
        'input[name="UserName"]',
        'input[name="Username"]',
        'input[name*="email" i]',
        'input[id*="email" i]',
        'input[name*="user" i]',
        'input[id*="user" i]',
        'input[type="email"]',
        'form input[type="text"]',
    ]
    password_selectors = [
        'input#Password',
        'input#password',
        'input[name="Password"]',
        'input[name="password"]',
        'input[type="password"]',
    ]

    user_input = _first_visible_locator(page, username_selectors)
    pass_input = _first_visible_locator(page, password_selectors)
    if user_input is None or pass_input is None:
        log_step("Bignition login form fields were not found for auto-fill.")
        return False

    try:
        log_step("Filling Bignition username/password.")
        user_input.click(timeout=5_000)
        user_input.fill(username)
        pass_input.click(timeout=5_000)
        pass_input.fill(password)
    except Exception:
        log_step("Failed filling Bignition login fields.")
        return False

    submit_selectors = [
        'button[type="submit"]:has-text("Login")',
        'button:has-text("Login")',
        'input[type="submit"]',
    ]
    submit_button = _first_visible_locator(page, submit_selectors)
    if submit_button is not None:
        try:
            log_step("Submitting Bignition login form.")
            submit_button.click(timeout=10_000)
            return True
        except Exception:
            pass

    try:
        log_step("Submitting Bignition login via Enter key.")
        pass_input.press("Enter")
        return True
    except Exception:
        return False


def login_and_save_state(p, login_url: str, storage_state_path: Path) -> bool:
    credentials = load_bignition_credentials(BIGNITION_CREDENTIALS_PATH)
    if credentials is not None:
        # Preferred: headless login when saved credentials exist.
        browser = launch_browser(p, headless=True)
        context = browser.new_context()
        page = context.new_page()
        log_step("Opening Bignition login page (headless).")
        page.goto(login_url, wait_until="domcontentloaded")

        log_step("Attempting Bignition auto-login with saved credentials.")
        did_submit = try_submit_bignition_login(page, credentials[0], credentials[1])
        if not did_submit:
            context.close()
            browser.close()
            log_step("Saved Bignition credentials exist but could not be submitted.")
            return False

        try:
            log_step("Waiting for Bignition dashboard after credential submit (120s).")
            page.wait_for_url(re.compile(r"/Home/Dashboard"), timeout=120_000)
            log_step("Bignition auto-login reached dashboard.")
        except Exception:
            context.close()
            browser.close()
            log_step("Bignition auto-login timed out waiting for dashboard.")
            return False

        storage_state_path.parent.mkdir(parents=True, exist_ok=True)
        context.storage_state(path=str(storage_state_path))
        log_step(f"Saved Bignition storage state to: {storage_state_path}")
        context.close()
        browser.close()
        return True

    # No saved credentials: keep interactive fallback.
    browser = launch_browser(p, headless=False)
    context = browser.new_context()
    page = context.new_page()
    log_step("Opening Bignition login page (interactive fallback).")
    page.goto(login_url, wait_until="domcontentloaded")
    log_step("No saved Bignition credentials found; waiting for manual login.")
    try:
        page.wait_for_url(re.compile(r"/Home/Dashboard"), timeout=180_000)
        log_step("Manual login reached dashboard.")
    except Exception:
        print("After completing login (and any MFA), press Enter to continue.")
        input()

    storage_state_path.parent.mkdir(parents=True, exist_ok=True)
    context.storage_state(path=str(storage_state_path))
    log_step(f"Saved Bignition storage state to: {storage_state_path}")
    context.close()
    browser.close()
    return True


def download_report_headless(p, storage_state_path: Path) -> bool:
    if not storage_state_path.is_file():
        log_step("No Bignition storage_state file found for headless run.")
        return False

    browser = launch_browser(p, headless=True)
    context = browser.new_context(accept_downloads=True, storage_state=str(storage_state_path))
    page = context.new_page()
    log_step("Navigating to Bignition dashboard using saved session.")
    page.goto(DASHBOARD_URL, wait_until="domcontentloaded")
    if "/Account/Login" in page.url:
        log_step("Saved Bignition session is invalid (redirected to login).")
        context.close()
        browser.close()
        return False

    try:
        log_step("Opening Reports -> Agency Reports.")
        page.get_by_role("link", name="Reports").click()
        page.get_by_role("link", name=re.compile(r"Agency Reports", re.I)).click()
    except Exception:
        try:
            log_step("Fallback text navigation to Agency Reports.")
            page.locator('a:has-text("Reports")').first.click()
            page.locator('a:has-text("Agency Reports")').first.click()
        except Exception:
            log_step("Direct URL fallback for Agency Reports.")
            page.goto(AGENCY_REPORTS_URL, wait_until="domcontentloaded")

    page.wait_for_url(re.compile(r"/Reports/AgencyReports"), timeout=60_000)
    log_step("Opened Agency Reports.")

    downloads_path = WORKING_FILES_DIR
    downloads_path.mkdir(parents=True, exist_ok=True)
    log_step("Requesting Opportunities by Producer report download.")
    with page.expect_download(timeout=120_000) as download_info:
        try:
            page.locator('a[href="/Reports/ViewAgencyReport?id=5"]').first.click()
        except Exception:
            try:
                page.get_by_role("link", name=re.compile(r"Opportunities by Producer", re.I)).click()
            except Exception:
                try:
                    page.locator('a:has-text("Opportunities by Producer")').first.click()
                except Exception:
                    log_step("Direct URL fallback for Opportunities by Producer report.")
                    page.goto(OPPS_BY_PRODUCER_URL, wait_until="domcontentloaded")

    download = download_info.value
    target_path = downloads_path / download.suggested_filename
    download.save_as(target_path)
    log_step(f"Downloaded Bignition report to: {target_path}")

    context.close()
    browser.close()
    return True


def open_login_screen_playwright(url: str, storage_state_path: Path) -> bool:
    try:
        from playwright.sync_api import sync_playwright
    except Exception as exc:
        log_step(f"Playwright not available: {exc}")
        return False

    with sync_playwright() as p:
        if storage_state_path.is_file():
            log_step("Using existing Bignition storage state (headless).")
            if download_report_headless(p, storage_state_path):
                return True
            log_step("Stored Bignition session invalid. Re-authentication required.")

        if not login_and_save_state(p, url, storage_state_path):
            return False

        log_step("Retrying report download after refreshing storage state.")
        if not download_report_headless(p, storage_state_path):
            return False

        log_step("Bignition report workflow complete.")
        return True


def open_login_screen_fallback(url: str) -> None:
    chrome_path = find_chrome()
    if chrome_path:
        subprocess.Popen([str(chrome_path), url], close_fds=True)
        log_step(f"Opened Chrome fallback: {url}")
        return

    webbrowser.open(url)
    log_step(f"Opened default browser fallback: {url}")


def main() -> int:
    os.chdir(PROJECT_ROOT)
    log_step("Starting Bignition workflow.")
    ok = open_login_screen_playwright(LOGIN_URL, STORAGE_STATE_PATH)
    if not ok:
        log_step("Bignition workflow failed.")
        return 1
    log_step("Bignition workflow finished successfully.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

