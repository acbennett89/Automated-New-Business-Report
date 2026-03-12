from __future__ import annotations

import argparse
import json
from datetime import datetime
import re
import time
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR.parent if SCRIPT_DIR.name.casefold() == "scripts" else SCRIPT_DIR
WORKING_FILES_DIR = PROJECT_ROOT / "Working Files"

EPIC_URL = "https://insu621.appliedepic.com/#/"
DATABASE_NAME = "INSU621_PROD"
ENTERPRISE_ID = "INSU621"
STORAGE_STATE_PATH = PROJECT_ROOT / "config" / "epic_storage_state.json"
EPIC_CREDENTIALS_PATH = PROJECT_ROOT / "config" / "epic_credentials.json"
REPORT_NAME = "Automated Report - Production for Bignition Comparison"
DIAGNOSTICS_DIR = WORKING_FILES_DIR / "Diagnostics"


def log_step(message: str) -> None:
    print(f"[{datetime.now():%H:%M:%S}] {message}", flush=True)


def seconds_since(started_at: float) -> str:
    return f"{time.perf_counter() - started_at:.1f}s"


def launch_browser(p, headless: bool):
    if headless:
        return p.chromium.launch(headless=True)
    try:
        return p.chromium.launch(channel="chrome", headless=False)
    except Exception:
        return p.chromium.launch(headless=False)


def is_visible(locator) -> bool:
    try:
        return locator.count() > 0 and locator.first.is_visible()
    except Exception:
        return False


def wait_visible(page, locator, timeout_ms: int) -> bool:
    end = time.time() + timeout_ms / 1000.0
    while time.time() < end:
        if is_visible(locator):
            return True
        try:
            page.wait_for_timeout(250)
        except Exception:
            pass
    return False


def manual_step_or_fail(headless: bool, fail_log: str, prompt: str) -> bool:
    if headless:
        log_step(fail_log)
        return False
    input(prompt)
    return True


def describe_page(page) -> str:
    try:
        url = page.url
    except Exception:
        url = "<unknown>"
    try:
        title = page.title()
    except Exception:
        title = "<unavailable>"
    return f"URL={url} | Title={title}"


def save_diagnostic(page, name: str) -> None:
    DIAGNOSTICS_DIR.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_name = re.sub(r"[^A-Za-z0-9._-]+", "_", name).strip("_") or "epic"
    screenshot_path = DIAGNOSTICS_DIR / f"{stamp}_{safe_name}.png"
    meta_path = DIAGNOSTICS_DIR / f"{stamp}_{safe_name}.txt"

    meta_lines = [describe_page(page)]
    try:
        meta_lines.append(f"Viewport={page.viewport_size}")
    except Exception:
        pass

    try:
        page.screenshot(path=str(screenshot_path), full_page=True)
        meta_lines.append(f"Screenshot={screenshot_path}")
    except Exception as exc:
        meta_lines.append(f"ScreenshotError={exc}")

    meta_path.write_text("\n".join(meta_lines), encoding="utf-8")
    log_step(f"Diagnostic saved: {meta_path}")


def load_epic_credentials(path: Path) -> tuple[str, str] | None:
    if not path.is_file():
        return None
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return None

    usercode = str(raw.get("usercode", "")).strip()
    password = str(raw.get("password", "")).strip()
    if not usercode or not password:
        return None
    return usercode, password


def submit_usercode_password_if_present(page, usercode: str, password: str, timeout_ms: int = 20_000) -> bool:
    user_input = page.locator('input#usercode, input[name="usercode"], input[id="usercode"]')
    password_input = page.locator('input#password, input[name="password"], input[id="password"], input[type="password"]')
    if not wait_visible(page, user_input, timeout_ms):
        return False
    if not is_visible(password_input):
        return False

    try:
        log_step("Login step: filling usercode and password.")
        user_input.first.click(timeout=5_000)
        user_input.first.fill(usercode)
        password_input.first.click(timeout=5_000)
        password_input.first.fill(password)
    except Exception:
        log_step("Login step: failed filling usercode/password fields.")
        return False

    login_submit = page.locator(
        'button[type="submit"]:has-text("Login"), button:has-text("Login"), button:has(div:has-text("Login"))'
    )
    try:
        login_submit.first.click(timeout=10_000)
        log_step("Login step: submitted usercode/password form.")
        return True
    except Exception:
        log_step("Login step: failed clicking usercode/password Login button.")
        return False


def click_login_and_wait(page, credentials: tuple[str, str] | None = None, headless: bool = False):
    started = time.perf_counter()
    log_step("Login step: locating Login button.")
    login_button = page.locator("div.message button:has-text('Login')")
    if not wait_visible(page, login_button, 15_000):
        login_button = page.get_by_role("button", name=re.compile(r"login", re.I))
    if not is_visible(login_button):
        log_step("Login step: Login button not visible.")
        return False

    click_attempted = False
    popup_page = None
    if credentials is not None:
        try:
            log_step("Login step: clicking Login and waiting for popup/new window (12s).")
            with page.context.expect_page(timeout=12_000) as popup_info:
                login_button.first.click(timeout=5_000)
                click_attempted = True
            popup_page = popup_info.value
            log_step("Login step: popup/new window detected.")
        except Exception:
            popup_page = None

    if not click_attempted:
        try:
            log_step("Login step: clicking Login.")
            login_button.first.click(timeout=5_000)
            click_attempted = True
        except Exception:
            log_step("Login step: Login click failed.")
            return False

    if credentials is not None:
        usercode, password = credentials
        if popup_page is not None:
            try:
                popup_page.wait_for_load_state("domcontentloaded", timeout=20_000)
            except Exception:
                pass
            log_step(f"Login step: popup state -> {describe_page(popup_page)}")
            did_submit_popup = submit_usercode_password_if_present(popup_page, usercode, password, timeout_ms=30_000)
            if did_submit_popup:
                try:
                    log_step("Login step: waiting for popup/new window to close (45s).")
                    popup_page.wait_for_event("close", timeout=45_000)
                except Exception:
                    pass
            else:
                log_step("Login step: could not auto-submit credentials in popup/new window.")
        else:
            # Fallback for environments where the login form stays in the same page.
            submit_usercode_password_if_present(page, usercode, password, timeout_ms=15_000)

    try:
        log_step("Login step: waiting for post-login state (Continue, Database, or Home sidebar) (60s).")
        page.wait_for_selector(
            "button:has-text('Continue'), [data-automation-id='cboDatabase'], a[data-automation-id='sidebar-button-3 level-1']",
            timeout=60_000,
        )
        log_step(f"Login step complete in {seconds_since(started)}. {describe_page(page)}")
    except Exception:
        save_diagnostic(page, "login_timeout")
        if not manual_step_or_fail(
            headless,
            "Login step timed out in headless mode.",
            "Login step timed out; complete login manually, then press Enter.",
        ):
            return False
    return True


def fill_enterprise_id(page, enterprise_id: str) -> bool:
    selectors = [
        '[data-automation-id="txtEnterprise"] input[type="text"]',
        'input[name*="Enterprise" i]',
        'input[id*="Enterprise" i]',
        'input[placeholder*="Enterprise" i]',
        'input[aria-label*="Enterprise" i]',
    ]
    for selector in selectors:
        try:
            locator = page.locator(selector)
            if locator.count() > 0 and locator.first.is_visible():
                locator.first.fill(enterprise_id)
                return True
        except Exception:
            continue

    try:
        page.get_by_label(re.compile(r"enterprise", re.I)).fill(enterprise_id)
        return True
    except Exception:
        pass

    try:
        dialog = page.get_by_role("dialog")
        if dialog.count() > 0:
            input_in_dialog = dialog.first.locator("input[type='text']").first
            if input_in_dialog.is_visible():
                input_in_dialog.fill(enterprise_id)
                return True
    except Exception:
        pass

    try:
        block = page.locator('div:has-text("Enterprise ID")').first
        input_near = block.locator("input[type='text']").first
        if input_near.is_visible():
            input_near.fill(enterprise_id)
            return True
    except Exception:
        pass

    try:
        visible_text_inputs = page.locator("input[type='text']")
        if visible_text_inputs.count() == 1 and visible_text_inputs.first.is_visible():
            visible_text_inputs.first.fill(enterprise_id)
            return True
    except Exception:
        pass

    return False


def select_database(page, database_name: str) -> bool:
    started = time.perf_counter()
    log_step(f"Database step: selecting database '{database_name}'.")
    try:
        log_step("Database step: waiting for database combo (120s).")
        page.wait_for_selector('[data-automation-id="cboDatabase"]', timeout=120_000)
    except Exception:
        log_step("Database step: database combo not found.")
        return False

    combo = page.locator('[data-automation-id="cboDatabase"]')
    if combo.count() == 0:
        return False

    try:
        combo.locator(".drop-btn").first.click()
    except Exception:
        try:
            combo.click()
        except Exception:
            pass

    try:
        log_step("Database step: waiting for database rows (30s).")
        page.wait_for_selector('div[data-automation-id^="cboDatabase body-row"]', timeout=30_000)
    except Exception:
        pass

    try:
        rows = page.locator('div[data-automation-id^="cboDatabase body-row"]')
        target = rows.filter(has=page.locator("span.text", has_text=database_name))
        if target.count() > 0:
            target.first.click()
            log_step(f"Database step complete in {seconds_since(started)}.")
            return True
    except Exception:
        pass

    try:
        page.get_by_text(database_name, exact=True).click()
        log_step(f"Database step complete via text match in {seconds_since(started)}.")
        return True
    except Exception:
        log_step("Database step failed to select database.")
        return False


def click_continue(page) -> None:
    try:
        page.get_by_role("button", name=re.compile(r"continue", re.I)).click()
        return
    except Exception:
        pass

    try:
        page.locator("button:has-text('Continue')").first.click()
    except Exception:
        pass


def open_reports_tab(context, page):
    started = time.perf_counter()
    log_step("Reports step: opening Reports/Marketing.")

    nav_button = page.locator('a[data-automation-id="sidebar-button-3 level-1"]').first

    # Primary path: Chromium commonly navigates in the same tab.
    try:
        log_step("Reports step: trying same-tab navigation.")
        nav_button.click(timeout=8_000)
        log_step("Reports step: waiting fixed 3s before report lookup.")
        page.wait_for_timeout(3_000)
        log_step(f"Reports step complete (same tab) in {seconds_since(started)}. URL={page.url}")
        return page
    except Exception:
        pass

    # Fallback path: environments that still open in a new tab.
    try:
        log_step("Reports step: trying new-tab fallback.")
        with context.expect_page(timeout=8_000) as page_info:
            nav_button.click(timeout=8_000)
        reports_page = page_info.value
        try:
            log_step("Reports step: waiting fixed 3s in new tab before report lookup.")
            reports_page.wait_for_timeout(3_000)
        except Exception:
            reports_page.wait_for_load_state("domcontentloaded")
        reports_page.bring_to_front()
        log_step(f"Reports step complete (new tab) in {seconds_since(started)}. URL={reports_page.url}")
        return reports_page
    except Exception:
        log_step(f"Reports step failed after {seconds_since(started)}. Manual navigation needed.")
        return page


def pick_attached_epic_page(context):
    try:
        pages = [p for p in context.pages if not p.is_closed()]
    except Exception:
        pages = []
    if not pages:
        return None
    # Chromium EPIC flow may keep URL stable; use most recently opened tab.
    return pages[-1]


def select_my_reports(page) -> bool:
    started = time.perf_counter()
    log_step("My Reports step: selecting My Reports in left sidebar.")
    def normalize(value: str) -> str:
        return re.sub(r"\s+", " ", value or "").strip().casefold()

    # Give EPIC sidebar time to hydrate after entering Reports/Marketing.
    try:
        log_step("My Reports step: waiting for sidebar links to load (20s).")
        page.wait_for_selector("a.sidebar-button", timeout=20_000)
    except Exception:
        pass

    for attempt in range(1, 21):
        log_step(f"My Reports step: attempt {attempt}/20.")
        try:
            if page.is_closed():
                log_step("My Reports step: page closed.")
                return False
        except Exception:
            log_step("My Reports step: page state unavailable.")
            return False

        target = None

        # Strict first choice from your provided HTML.
        try:
            strict = page.locator('a[data-automation-id="sidebar-button-1 level-1"]')
            strict_count = strict.count()
        except Exception:
            strict_count = 0
        for i in range(strict_count):
            candidate = strict.nth(i)
            try:
                if not candidate.is_visible():
                    continue
                txt = candidate.inner_text(timeout=1_000)
            except Exception:
                continue
            if normalize(txt) == "my reports":
                target = candidate
                break

        # Fallback to other sidebar links with exact visible text.
        if target is None:
            try:
                links = page.locator("a.sidebar-button")
                count = links.count()
            except Exception:
                count = 0
            for i in range(count):
                link = links.nth(i)
                try:
                    if not link.is_visible():
                        continue
                    text = link.inner_text(timeout=1_000)
                except Exception:
                    continue
                if normalize(text) == "my reports":
                    target = link
                    break

        if target is None:
            try:
                page.wait_for_timeout(500)
            except Exception:
                return False
            continue

        try:
            target.scroll_into_view_if_needed()
        except Exception:
            pass
        try:
            target.click(timeout=5_000)
        except Exception:
            try:
                target.click(force=True, timeout=5_000)
            except Exception:
                try:
                    page.wait_for_timeout(250)
                except Exception:
                    return False
                continue

        # Wait for the reports virtual list to be available after selecting My Reports.
        try:
            page.wait_for_selector('div[data-automation-id^="vlvwReports body-row item-"]', timeout=30_000)
            log_step(f"My Reports step complete in {seconds_since(started)}.")
            return True
        except Exception:
            try:
                page.wait_for_timeout(500)
            except Exception:
                return False

    log_step(f"My Reports step failed after {seconds_since(started)}.")
    return False


def open_report_by_name(page, report_name: str) -> bool:
    started = time.perf_counter()
    log_step(f"Report Open step: opening '{report_name}'.")
    def normalize(value: str) -> str:
        return re.sub(r"\s+", " ", value or "").strip().casefold()

    def report_opened() -> bool:
        try:
            return page.locator("text=Modify Criteria").count() > 0
        except Exception:
            return False

    def try_open_row(row) -> bool:
        try:
            row.scroll_into_view_if_needed()
        except Exception:
            pass
        try:
            row.click(timeout=5_000)
        except Exception:
            pass
        page.wait_for_timeout(150)

        click_targets = [
            row,
            row.locator("div.body-cell.first").first,
            row.locator("span.text").first,
        ]
        for target in click_targets:
            try:
                target.dblclick(delay=120, timeout=8_000)
            except Exception:
                continue
            try:
                page.wait_for_selector("text=Modify Criteria", timeout=10_000)
                return True
            except Exception:
                continue
        return report_opened()

    try:
        log_step("Report Open step: waiting for reports rows (60s).")
        page.wait_for_selector('div[data-automation-id^="vlvwReports body-row item-"]', timeout=60_000)
    except Exception:
        log_step("Report Open step: reports rows not found.")
        return False

    target_name = normalize(report_name)

    # Fast-path for this specific report row in your environment.
    try:
        row_11 = page.locator('div[data-automation-id="vlvwReports body-row item-11"]').first
        if row_11.count() > 0:
            row_11_text = row_11.locator("div.body-cell.first span.text").first.inner_text(timeout=2_500)
            if normalize(row_11_text) == target_name and try_open_row(row_11):
                log_step(f"Report Open step complete via row item-11 in {seconds_since(started)}.")
                return True
    except Exception:
        pass

    # The reports list is virtualized, so iterate visible rows and scroll until found.
    for scan_pass in range(1, 41):
        if scan_pass == 1 or scan_pass % 5 == 0:
            log_step(f"Report Open step: scan pass {scan_pass}/40.")
        try:
            page.wait_for_timeout(250)
        except Exception:
            return False

        rows = page.locator('div[data-automation-id^="vlvwReports body-row item-"]')
        try:
            row_count = rows.count()
        except Exception:
            row_count = 0

        for idx in range(row_count):
            row = rows.nth(idx)
            try:
                text = row.locator("div.body-cell.first span.text").first.inner_text(timeout=2_500)
            except Exception:
                continue
            if normalize(text) != target_name:
                continue
            if try_open_row(row):
                log_step(f"Report Open step complete in {seconds_since(started)} (scan pass {scan_pass}).")
                return True

        # Fallback direct match in case row iteration misses the target due rendering.
        try:
            label = page.get_by_text(report_name, exact=True).first
            if label.count() > 0 and label.is_visible():
                row = label.locator("xpath=ancestor::div[contains(@class,'body-row')]").first
                if row.count() > 0 and try_open_row(row):
                    log_step(f"Report Open step complete via text fallback in {seconds_since(started)}.")
                    return True
        except Exception:
            pass

        # Your sample showed this report as item-11; use it as a strict fallback.
        try:
            row_11 = page.locator('div[data-automation-id="vlvwReports body-row item-11"]').first
            if row_11.count() > 0:
                row_11_text = row_11.locator("div.body-cell.first span.text").first.inner_text(timeout=2_500)
                if normalize(row_11_text) == target_name and try_open_row(row_11):
                    log_step(f"Report Open step complete via row-11 fallback in {seconds_since(started)}.")
                    return True
        except Exception:
            pass

        try:
            page.locator("div.body-rows.has-header").first.evaluate(
                "(el) => { el.scrollTop = Math.min(el.scrollTop + 280, el.scrollHeight); }"
            )
        except Exception:
            try:
                page.mouse.wheel(0, 400)
            except Exception:
                pass
        try:
            page.wait_for_timeout(450)
        except Exception:
            return False

    log_step(f"Report Open step failed after {seconds_since(started)}.")
    return False


def set_accounting_month_value(container, month: str, year: int) -> bool:
    started = time.perf_counter()
    log_step(f"Criteria step: setting month='{month}', year={year}.")
    month_index_map = {
        "january": 0,
        "february": 1,
        "march": 2,
        "april": 3,
        "may": 4,
        "june": 5,
        "july": 6,
        "august": 7,
        "september": 8,
        "october": 9,
        "november": 10,
        "december": 11,
    }
    target_index = month_index_map.get(month.strip().casefold())
    if target_index is None:
        log_step("Criteria step: invalid month name.")
        return False

    month_selected = False
    try:
        combo = container.locator("asi-combo-box").first
        drop_btn = combo.locator(".drop-btn").first
        drop_btn.click(timeout=5_000)
        page = container.page
        page.wait_for_timeout(150)

        rows_scroller = combo.locator("div.body-rows.has-header.in-drop-down").first
        try:
            rows_scroller.evaluate("(el, idx) => { el.scrollTop = Math.max(0, (idx - 1) * 26); }", target_index)
        except Exception:
            pass

        option_row = combo.locator(f"div.body-row[index='{target_index}']").first
        option_row.scroll_into_view_if_needed()
        option_row.click(timeout=5_000)
        month_selected = True
        log_step(f"Criteria step: selected month '{month}' from dropdown.")
    except Exception:
        try:
            # Fallback: type month name and commit.
            month_input = container.locator("asi-combo-box input[focusname='month']").first
            month_input.click(timeout=5_000)
            month_input.press("Control+A")
            month_input.type(month, delay=20)
            month_input.press("Enter")
            month_selected = True
            log_step(f"Criteria step: selected month '{month}' via typed fallback.")
        except Exception:
            log_step("Criteria step: failed selecting month.")
            return False

    if not month_selected:
        return False

    try:
        year_input = container.locator("asi-integer-edit input").first
        year_input.click(timeout=5_000)
        year_input.press("Control+A")
        year_input.fill(str(year))
        year_input.press("Enter")
    except Exception:
        log_step("Criteria step: failed setting year input.")
        return False
    log_step(f"Criteria step: month/year set in {seconds_since(started)}.")
    return True


def update_accounting_month_criteria(page, from_year: int, to_year: int) -> bool:
    started = time.perf_counter()
    log_step(f"Criteria step: updating Accounting Month ({from_year} -> {to_year}).")
    try:
        criteria_row = page.locator('div[data-automation-id^="vlvwCriteria body-row item-"]').filter(
            has=page.locator("span.text", has_text="Accounting Month")
        )
        if criteria_row.count() == 0:
            # Fallback from your provided structure.
            criteria_row = page.locator('div[data-automation-id="vlvwCriteria body-row item-5"]')
            if criteria_row.count() == 0:
                log_step("Criteria step: Accounting Month row not found.")
                return False
        criteria_row.first.click()
        page.wait_for_timeout(100)
        criteria_row.first.dblclick()
    except Exception:
        log_step("Criteria step: failed opening Accounting Month row.")
        return False

    months = page.locator("asi-accounting-month")
    visible_month_containers = []
    end = time.time() + 10
    log_step("Criteria step: waiting for month controls (10s).")
    while time.time() < end and len(visible_month_containers) < 2:
        visible_month_containers = []
        try:
            count = months.count()
        except Exception:
            count = 0
        for idx in range(count):
            m = months.nth(idx)
            try:
                if m.is_visible():
                    visible_month_containers.append(m)
            except Exception:
                continue
        if len(visible_month_containers) < 2:
            page.wait_for_timeout(250)

    if len(visible_month_containers) < 2:
        log_step("Criteria step: month controls not available.")
        return False

    from_ok = set_accounting_month_value(visible_month_containers[0], "January", from_year)
    to_ok = set_accounting_month_value(visible_month_containers[1], "December", to_year)
    if from_ok and to_ok:
        log_step(f"Criteria step complete in {seconds_since(started)}.")
    else:
        log_step(f"Criteria step failed in {seconds_since(started)}.")
    return from_ok and to_ok


def generate_report_and_download(page, download_dir: Path) -> bool:
    started = time.perf_counter()
    log_step("Generate step: opening Actions -> Generate Report.")
    actions_button = page.locator("div.main-button[title^='Actions']")
    if actions_button.count() == 0:
        actions_button = page.locator("div.main-button").filter(
            has=page.locator("span.text", has_text="Actions")
        )
    if actions_button.count() == 0:
        log_step("Generate step: Actions button not found.")
        return False

    download_dir.mkdir(parents=True, exist_ok=True)

    log_step("Generate step: waiting for download after Generate Report (600s).")
    with page.expect_download(timeout=600_000) as download_info:
        actions_button.first.click()
        option = page.locator('li[data-automation-id="dropdown-menu-item-103"]')
        if option.count() == 0:
            option = page.locator("li.dropdown-menu-item").filter(
                has=page.locator("span.text", has_text="Generate Report")
            )
        if option.count() == 0:
            log_step("Generate step: Generate Report menu item not found.")
            return False
        option.first.click()

    download = download_info.value
    target_path = download_dir / download.suggested_filename
    download.save_as(str(target_path))
    log_step(f"Generate step complete in {seconds_since(started)}. Downloaded to: {target_path}")
    return True


def logout_epic(page) -> bool:
    started = time.perf_counter()
    log_step("Logout step: clicking Logout button.")

    logout_button = page.locator("div.main-button[title^='Logout']")
    if logout_button.count() == 0:
        logout_button = page.locator("div.logout div.main-button").filter(
            has=page.locator("span.text", has_text="Logout")
        )
    if logout_button.count() == 0:
        log_step("Logout step: Logout button not found.")
        return False

    try:
        logout_button.first.click(timeout=8_000)
    except Exception:
        log_step("Logout step: failed clicking Logout button.")
        return False

    try:
        log_step("Logout step: waiting for Logout confirmation dialog (30s).")
        logout_dialog = page.locator("div.message-box").filter(
            has=page.locator("div.title", has_text=re.compile(r"^Logout$", re.I))
        ).first
        logout_dialog.wait_for(state="visible", timeout=30_000)
        yes_button = logout_dialog.locator('button[data-automation-id="Yes"]').first
        yes_button.click(timeout=8_000)
        log_step("Logout step: selected 'Yes' to close all Epic windows.")
    except Exception:
        log_step("Logout step: Logout confirmation dialog handling failed.")
        return False

    try:
        log_step("Logout step: waiting for Save Changes dialog (30s).")
        save_dialog = page.locator("div.message-box").filter(
            has=page.locator("div.title", has_text=re.compile(r"^Save Changes$", re.I))
        ).first
        save_dialog.wait_for(state="visible", timeout=30_000)
        no_button = save_dialog.locator('button[data-automation-id="No"]').first
        no_button.click(timeout=8_000)
        log_step("Logout step: selected 'No' for Save Changes.")
    except Exception:
        log_step("Logout step: Save Changes dialog handling failed.")
        return False

    try:
        log_step("Logout step: waiting for logout completion marker (30s).")
        page.wait_for_selector("div.message:has-text('You are logged out'), div.message a:has-text('Login here')", timeout=30_000)
    except Exception:
        # EPIC may close tabs/windows quickly after logout; this is not fatal.
        pass

    log_step(f"Logout step complete in {seconds_since(started)}.")
    return True


def run_epic_iteration(
    p,
    storage_state_path: Path,
    step: str,
    attach_open_browser: bool = False,
    cdp_url: str = "http://127.0.0.1:9222",
) -> bool:
    started = time.perf_counter()
    log_step(f"Iterative flow start: step='{step}', attach_open_browser={attach_open_browser}.")
    credentials = load_epic_credentials(EPIC_CREDENTIALS_PATH)
    if credentials is not None:
        log_step("Iterative flow: loaded saved EPIC credentials.")
    else:
        log_step("Iterative flow: no saved EPIC credentials found.")
    attached = False
    browser = None
    context = None
    page = None

    if attach_open_browser:
        try:
            log_step(f"Iterative flow: connecting over CDP at {cdp_url}.")
            browser = p.chromium.connect_over_cdp(cdp_url)
            attached = True
        except Exception as exc:
            log_step(f"Iterative flow: CDP attach failed: {exc}")
            log_step("Iterative flow: start Chrome with --remote-debugging-port=9222, then retry.")
            return False

        if not browser.contexts:
            log_step("Iterative flow: attached, but no browser contexts/pages found.")
            return False
        context = browser.contexts[0]
        page = pick_attached_epic_page(context)
        if page is None:
            page = context.new_page()
        try:
            page.bring_to_front()
        except Exception:
            pass
    else:
        log_step("Iterative flow: launching local browser context.")
        browser = launch_browser(p, headless=False)
        context_kwargs = {"accept_downloads": True}
        if storage_state_path.is_file():
            context_kwargs["storage_state"] = str(storage_state_path)
        context = browser.new_context(**context_kwargs)
        page = context.new_page()
        log_step(f"Iterative flow: navigating to {EPIC_URL}.")
        page.goto(EPIC_URL, wait_until="domcontentloaded")

    if page is None:
        log_step("Iterative flow: no usable page found.")
        return False

    if credentials is not None:
        try:
            submit_usercode_password_if_present(page, credentials[0], credentials[1], timeout_ms=2_000)
        except Exception:
            pass

    login_button = page.locator('button:has-text("Login")')
    if is_visible(login_button):
        log_step("Iterative flow: session not authenticated in active browser page.")
        if credentials is not None:
            click_login_and_wait(page, credentials=credentials)
        if is_visible(login_button):
            if attached:
                input("Log in in that browser, then press Enter to continue.")
            else:
                input("Run full flow first to log in, then press Enter to close.")
                context.close()
                browser.close()
                return False

    def has_any_rows(target_page, selector: str) -> bool:
        try:
            return (not target_page.is_closed()) and target_page.locator(selector).count() > 0
        except Exception:
            return False

    reports_page = page
    try:
        in_report_program = "program=Report" in (page.url or "")
    except Exception:
        in_report_program = False
    has_report_rows = has_any_rows(page, 'div[data-automation-id^="vlvwReports body-row item-"]')
    has_criteria_rows = has_any_rows(page, 'div[data-automation-id^="vlvwCriteria body-row item-"]')
    log_step(f"Iterative flow: state in_report_program={in_report_program}, has_report_rows={has_report_rows}, has_criteria_rows={has_criteria_rows}.")

    # If already on criteria screen for this report, do not force a My Reports click.
    if not (step in {"criteria", "generate"} and has_criteria_rows):
        if not in_report_program and not has_report_rows:
            log_step("Iterative flow: opening Reports/Marketing.")
            reports_page = open_reports_tab(context, page) or page
            try:
                if reports_page.is_closed():
                    replacement = pick_attached_epic_page(context)
                    if replacement is not None:
                        reports_page = replacement
            except Exception:
                pass

            # If still not on reports list, ask for manual navigation and continue safely.
            reports_ok = has_any_rows(reports_page, 'div[data-automation-id^="vlvwReports body-row item-"]')
            if not reports_ok:
                log_step("Iterative flow: could not verify Reports/Marketing list.")
                input("Open Reports/Marketing -> My Reports manually, then press Enter.")
                replacement = pick_attached_epic_page(context)
                if replacement is not None:
                    reports_page = replacement

    if step in {"open", "criteria", "generate"} and not has_criteria_rows:
        if not open_report_by_name(reports_page, REPORT_NAME):
            log_step(f"Iterative flow: auto-open report '{REPORT_NAME}' failed.")
            input("Open the report manually, then press Enter to continue.")

    if step in {"criteria", "generate"}:
        current_year = datetime.now().year
        if not update_accounting_month_criteria(reports_page, from_year=current_year - 1, to_year=current_year):
            log_step("Iterative flow: auto-set Accounting Month failed.")
            input("Set Accounting Month manually, then press Enter to continue.")

    if step == "generate":
        download_dir = WORKING_FILES_DIR
        if not generate_report_and_download(reports_page, download_dir):
            log_step("Iterative flow: auto-generate/download failed.")
            input("Run Actions -> Generate Report manually, then press Enter.")

    log_step(f"Iterative flow complete in {seconds_since(started)}.")
    input(f"Iterative step '{step}' complete. Press Enter to close...")
    if not attached:
        context.close()
        browser.close()
    return True


def run_epic_flow(p, storage_state_path: Path, headless: bool, allow_login: bool) -> bool:
    started = time.perf_counter()
    log_step(f"Full flow start: headless={headless}, allow_login={allow_login}.")
    browser = launch_browser(p, headless=headless)
    context_kwargs = {}
    if storage_state_path.is_file():
        context_kwargs["storage_state"] = str(storage_state_path)
    context_kwargs["accept_downloads"] = True

    context = browser.new_context(**context_kwargs)
    page = context.new_page()
    log_step(f"Full flow: navigating to {EPIC_URL}.")
    page.goto(EPIC_URL, wait_until="domcontentloaded")
    log_step(f"Full flow: initial page state -> {describe_page(page)}")
    credentials = load_epic_credentials(EPIC_CREDENTIALS_PATH)
    if credentials is not None:
        log_step("Full flow: loaded saved EPIC credentials.")
    else:
        log_step("Full flow: no saved EPIC credentials found.")

    enterprise_input = page.locator(
        '[data-automation-id="txtEnterprise"] input[type="text"], '
        'div:has-text("Enterprise ID") input[type="text"], '
        'div[role="dialog"] input[type="text"], '
        'input[aria-label*="Enterprise" i]'
    )
    login_button = page.locator('button:has-text("Login")')
    database_combo = page.locator('[data-automation-id="cboDatabase"]')
    reports_sidebar_button = page.locator('a[data-automation-id="sidebar-button-3 level-1"]')

    log_step("Full flow: checking enterprise/login screen state.")
    if not is_visible(login_button) and wait_visible(page, enterprise_input, 8_000):
        if not allow_login:
            context.close()
            browser.close()
            return False
        if fill_enterprise_id(page, ENTERPRISE_ID):
            log_step("Full flow: enterprise ID filled; clicking Continue.")
            click_continue(page)
            log_step("Full flow: waiting for Login button after enterprise submit (30s).")
            wait_visible(page, login_button, 30_000)
            log_step(f"Full flow: post-enterprise state -> {describe_page(page)}")

    if credentials is not None:
        try:
            submit_usercode_password_if_present(page, credentials[0], credentials[1], timeout_ms=2_000)
        except Exception:
            pass

    if is_visible(login_button):
        if not allow_login:
            context.close()
            browser.close()
            return False
        log_step("Full flow: login required.")
        if not click_login_and_wait(page, credentials=credentials, headless=headless):
            context.close()
            browser.close()
            return False

    log_step("Full flow: determining whether database selection is required.")
    if wait_visible(page, reports_sidebar_button, 8_000):
        log_step(f"Full flow: already on home page; database selection skipped. {describe_page(page)}")
    else:
        log_step("Full flow: waiting for database combo (120s).")
        has_db_combo = wait_visible(page, database_combo, 120_000)
        if not has_db_combo:
            log_step("Full flow: database combo not found yet; extending wait by 60s.")
            has_db_combo = wait_visible(page, database_combo, 60_000)

        if has_db_combo:
            selected = False
            for db_attempt in range(1, 4):
                log_step(f"Full flow: database select attempt {db_attempt}/3.")
                if select_database(page, DATABASE_NAME):
                    selected = True
                    break
                try:
                    page.wait_for_timeout(1000)
                except Exception:
                    pass

            if not selected:
                log_step(f"Full flow: auto-select database '{DATABASE_NAME}' failed.")
                if not manual_step_or_fail(
                    headless,
                    "Full flow: database selection failed in headless mode.",
                    "After selecting the database, press Enter to continue.",
                ):
                    context.close()
                    browser.close()
                    return False

            log_step("Full flow: clicking Continue after database selection.")
            click_continue(page)
        else:
            # Some usercodes bypass database selection and go straight home.
            if wait_visible(page, reports_sidebar_button, 30_000):
                log_step(f"Full flow: reached home page without database selection. {describe_page(page)}")
            else:
                save_diagnostic(page, "login_database_state_not_detected")
                log_step("Full flow: neither database combo nor home sidebar detected.")
                if not manual_step_or_fail(
                    headless,
                    "Full flow: login/database landing state not detected in headless mode.",
                    "Complete login/database steps manually, then press Enter to continue.",
                ):
                    context.close()
                    browser.close()
                    return False

    try:
        log_step("Full flow: waiting for sidebar Reports button (120s).")
        page.wait_for_selector('a[data-automation-id="sidebar-button-3 level-1"]', timeout=120_000)
        log_step(f"Full flow: sidebar detected. {describe_page(page)}")
    except Exception:
        save_diagnostic(page, "sidebar_not_visible")
        log_step("Full flow: sidebar not ready; waiting for manual confirmation.")
        if not manual_step_or_fail(
            headless,
            "Full flow: sidebar did not appear in headless mode.",
            "When the sidebar is visible, press Enter to continue.",
        ):
            context.close()
            browser.close()
            return False

    storage_state_path.parent.mkdir(parents=True, exist_ok=True)
    context.storage_state(path=str(storage_state_path))
    log_step(f"Full flow: saved storage state to {storage_state_path}.")

    reports_page = open_reports_tab(context, page)
    if reports_page is None:
        reports_page = page
    log_step(f"Full flow: reports page candidate -> {describe_page(reports_page)}")

    if not open_report_by_name(reports_page, REPORT_NAME):
        save_diagnostic(reports_page, "open_report_failed")
        log_step(f"Full flow: auto-open report '{REPORT_NAME}' failed.")
        if not manual_step_or_fail(
            headless,
            f"Full flow: auto-open report '{REPORT_NAME}' failed in headless mode.",
            "Open the report manually, then press Enter to continue.",
        ):
            context.close()
            browser.close()
            return False

    current_year = datetime.now().year
    if not update_accounting_month_criteria(reports_page, from_year=current_year - 1, to_year=current_year):
        save_diagnostic(reports_page, "accounting_month_failed")
        log_step("Full flow: auto-set Accounting Month failed.")
        if not manual_step_or_fail(
            headless,
            "Full flow: auto-set Accounting Month failed in headless mode.",
            "Set Accounting Month to January last year through December this year, then press Enter.",
        ):
            context.close()
            browser.close()
            return False

    download_dir = WORKING_FILES_DIR
    if not generate_report_and_download(reports_page, download_dir):
        save_diagnostic(reports_page, "generate_download_failed")
        log_step("Full flow: auto-generate/download failed.")
        if not manual_step_or_fail(
            headless,
            "Full flow: auto-generate/download failed in headless mode.",
            "Run Actions -> Generate Report manually, then press Enter.",
        ):
            context.close()
            browser.close()
            return False

    if not logout_epic(reports_page):
        save_diagnostic(reports_page, "logout_failed")
        log_step("Full flow: auto-logout failed.")
        if not manual_step_or_fail(
            headless,
            "Full flow: auto-logout failed in headless mode.",
            "Log out manually, then press Enter to continue.",
        ):
            context.close()
            browser.close()
            return False

    log_step(f"Full flow complete in {seconds_since(started)}.")
    if headless:
        context.close()
        browser.close()
    else:
        log_step("EPIC flow complete.")
        context.close()
        browser.close()
    return True


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--iter",
        dest="iter_step",
        choices=["reports", "open", "criteria", "generate"],
        help="Run only part of the EPIC flow using saved session state.",
    )
    parser.add_argument(
        "--attach-open-browser",
        action="store_true",
        help="Attach iterative run to an already-open Chrome via CDP (remote debugging port).",
    )
    parser.add_argument(
        "--cdp-url",
        default="http://127.0.0.1:9222",
        help="CDP URL for --attach-open-browser (default: http://127.0.0.1:9222).",
    )
    parser.add_argument(
        "--headed",
        action="store_true",
        help="Run full EPIC flow in a visible browser (default is headless).",
    )
    args = parser.parse_args()

    try:
        from playwright.sync_api import sync_playwright
    except Exception as exc:
        log_step(f"Startup failed: Playwright not available: {exc}")
        return 1

    with sync_playwright() as p:
        if args.iter_step:
            log_step(f"Running EPIC iterative step: {args.iter_step}")
            return 0 if run_epic_iteration(
                p,
                STORAGE_STATE_PATH,
                args.iter_step,
                attach_open_browser=args.attach_open_browser,
                cdp_url=args.cdp_url,
            ) else 1

        run_headless = not args.headed
        log_step(f"Running full EPIC flow ({'headless' if run_headless else 'headed'}).")
        if run_epic_flow(p, STORAGE_STATE_PATH, headless=run_headless, allow_login=True):
            return 0
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
