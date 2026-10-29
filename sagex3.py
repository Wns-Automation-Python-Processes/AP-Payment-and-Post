import os
import re
import json
import time
import sys
import openpyxl
from datetime import datetime

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException

# Visual execution speed control (in seconds)
STEP_DELAY = 0.0

CURRENT_DRIVER = None


def _app_dir():
    """Directory for runtime files (logs, output workbooks).

    When frozen into an executable, this is the folder NEXT TO the .exe so the
    user's files are easy to find. When running as a script, it is the script's
    own folder.
    """
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def _resource_dir():
    """Directory for bundled read-only resources (assets, templates).

    PyInstaller unpacks these to sys._MEIPASS at runtime.
    """
    if getattr(sys, "frozen", False):
        return getattr(sys, "_MEIPASS", _app_dir())
    return os.path.dirname(os.path.abspath(__file__))


BASE_DIR = _app_dir()
RESOURCE_DIR = _resource_dir()
DEFAULT_FILE_NAME = "InPut_Parameter_File.xlsx"
file_path = os.path.join(BASE_DIR, DEFAULT_FILE_NAME)

LOG_DIR = os.path.join(BASE_DIR, "logs")
os.makedirs(LOG_DIR, exist_ok=True)
STEPS_LOG = os.path.join(LOG_DIR, "steps.log")
PENDING_JSON = os.path.join(LOG_DIR, "pending_pay_numbers.json")

CURRENT_STEP = None

# Optional GUI hooks (set by sagex3_ui.py). When None, the script behaves as a
# normal command line tool that prints to stdout.
LOG_HOOK = None
STOP_FLAG = None
HEADLESS = False


def _emit(text):
    """Sends a line to the GUI hook if one is attached, otherwise prints it."""
    if LOG_HOOK is not None:
        try:
            LOG_HOOK(text)
            return
        except Exception:
            pass
    print(text)


def step(name):
    global CURRENT_STEP
    CURRENT_STEP = name
    _emit(f"[STEP] {name}")


def log_step(name, status, detail=""):
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"{stamp} - {status} - {name}" + (f" -> {detail}" if detail else "")
    with open(STEPS_LOG, "a", encoding="utf-8") as fh:
        fh.write(line + "\n")


def read_parameters(path):
    """Reads credentials and payment settings from the SAGEX3_Credentials sheet."""
    wb = openpyxl.load_workbook(path, data_only=False)
    ws = wb["SAGEX3_Credentials"]
    params = {
        "username": str(ws.cell(row=2, column=1).value or "").strip(),
        "password": str(ws.cell(row=2, column=2).value or "").strip(),
        "company_code": str(ws.cell(row=2, column=3).value or "").strip(),
        "bank_code": str(ws.cell(row=2, column=4).value or "").strip(),
    }
    raw_date = ws.cell(row=2, column=5).value
    if isinstance(raw_date, datetime):
        params["payment_date"] = raw_date.strftime("%d/%m/%Y")
    else:
        params["payment_date"] = str(raw_date or "").strip()
    wb.close()
    return params


def read_accounts(path):
    """Returns the Account_Details rows (from row 2) as a list of dicts.
    Columns: A=Vendor Code, B=Vendor Name, C=PayNumber, D=Supplier Email."""
    wb = openpyxl.load_workbook(path, data_only=False)
    ws = wb["Account_Details"]
    rows = []
    for row in range(2, ws.max_row + 1):
        supplier = str(ws.cell(row=row, column=1).value or "").strip()
        vendor_name = str(ws.cell(row=row, column=2).value or "").strip()
        pay_number = str(ws.cell(row=row, column=3).value or "").strip()
        email = str(ws.cell(row=row, column=4).value or "").strip()
        if supplier or vendor_name or pay_number or email:
            rows.append({"row": row, "supplier": supplier, "vendor_name": vendor_name,
                         "pay_number": pay_number, "email": email})
    wb.close()
    return rows


def write_pay_number_to_excel(path, row_number, pay_number, supplier=""):
    """Writes the pay number to Account_Details Column C for a specific row.
    Retries a few times in case Excel has the file locked; if still locked,
    saves the value to pending_pay_numbers.json so nothing is lost."""
    _emit(f"Updating {path} row {row_number} with PayNumber: {pay_number}...")
    last_err = None
    for attempt in range(1, 4):
        try:
            wb = openpyxl.load_workbook(path)
            ws = wb["Account_Details"]
            ws.cell(row=row_number, column=3).value = pay_number
            wb.save(path)
            wb.close()
            _emit(f"Successfully recorded '{pay_number}' in Account_Details Row {row_number}, Column C.")
            log_step("Write PayNumber to Excel", "OK", f"row {row_number}: {pay_number}")
            return True
        except PermissionError as e:
            last_err = e
            _emit(f"[RETRY] Excel is locked (is it open in Excel?). Attempt {attempt}/3 in 3s...")
            time.sleep(3)

    entry = {"timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
             "row": row_number, "supplier": supplier, "pay_number": pay_number}
    data = []
    if os.path.exists(PENDING_JSON):
        try:
            with open(PENDING_JSON, encoding="utf-8") as fh:
                data = json.load(fh)
        except Exception:
            data = []
    data.append(entry)
    with open(PENDING_JSON, "w", encoding="utf-8") as fh:
        json.dump(data, fh, indent=2)
    _emit(f"[WARN] Could not write to Excel after 3 tries. Saved PayNumber to {PENDING_JSON} "
          f"so it can be applied once the file is closed. ({last_err})")
    log_step("Write PayNumber to Excel", "FAIL", f"row {row_number}: locked, saved to {PENDING_JSON}")
    return False


def highlight_element(driver, element):
    """Draws a visible red border with a soft yellow background to highlight current action."""
    try:
        driver.execute_script(
            "arguments[0].style.border='3px solid red'; arguments[0].style.backgroundColor='rgba(255, 255, 0, 0.3)';",
            element
        )
    except Exception:
        pass


def unhighlight_element(driver, element):
    """Removes highlighting after the action is completed."""
    try:
        driver.execute_script(
            "arguments[0].style.border=''; arguments[0].style.backgroundColor='';",
            element
        )
    except Exception:
        pass


# JavaScript that returns true while ANY Sage X3 loading indicator is active:
# overlay masks, busy cursor (body class / s-busy), aria-busy regions, the wait
# bar, and in-flight jQuery AJAX. Used to block clicks/inputs until the screen
# has fully settled.
IS_BUSY_JS = r"""
const isVisible = (el) => {
    if (!el) return false;
    const s = window.getComputedStyle(el);
    if (s.display === 'none' || s.visibility === 'hidden') return false;
    if (s.opacity !== '' && parseFloat(s.opacity) === 0) return false;
    const r = el.getBoundingClientRect();
    return r.width > 0 && r.height > 0;
};
try {
    // 1) Body/HTML marks Sage as busy (drives the loading cursor).
    const bodyCls = (document.body ? document.body.className : '') + ' ' +
                    (document.documentElement ? document.documentElement.className : '');
    if (/\b(s-busy|s-loading|loading|is-loading|waiting|busy)\b/.test(bodyCls)) return true;

    // 2) Known loading overlays / masks / spinners.
    const overlaySel = [
        '.s-loading', '.s-loading-mask', '.loading-mask', '.s-busy',
        '.s-wait', '.s-waiting', '.s-progress-mask', '.s-progressbar',
        '.s-mask', '.splash', '.s-overlay-mask', '.s-vignette',
        '[aria-busy="true"]'
    ];
    for (const sel of overlaySel) {
        for (const el of document.querySelectorAll(sel)) {
            if (isVisible(el)) return true;
        }
    }

    // 3) Page-level busy cursor flags.
    const cur = window.getComputedStyle(document.body || document.documentElement).cursor;
    if (cur && /progress|wait/i.test(cur)) return true;

    // 4) In-flight AJAX (Sage uses jQuery/ExtJS under the hood).
    if (window.jQuery && jQuery.active && jQuery.active > 0) return true;

    return false;
} catch (e) {
    return false;
}
"""

# Generic overlay XPaths kept as a fallback for older/other Sage builds.
LOADER_XPATHS = [
    "//div[contains(@class, 's-loading')]",
    "//div[contains(@class, 'loading-mask')]",
    "//div[contains(@class, 's-busy')]",
    "//div[contains(@class, 's-wait')]",
    "//div[contains(@class, 's-progress-mask')]",
    "//*[@aria-busy='true']",
]


def _page_is_busy(driver):
    """Returns True if any Sage X3 loading indicator is currently active."""
    try:
        if driver.execute_script("return " + IS_BUSY_JS):
            return True
    except Exception:
        pass
    # Fallback: any visible legacy loader element.
    for xpath in LOADER_XPATHS:
        try:
            els = driver.find_elements(By.XPATH, xpath)
        except Exception:
            continue
        for el in els:
            try:
                if el.is_displayed():
                    return True
            except Exception:
                continue
    return False


def wait_for_sage_loader(driver, timeout=30):
    """Blocks until ALL Sage X3 loading indicators have cleared.

    Covers the busy cursor, overlay masks, aria-busy regions, the wait/progress
    bar and in-flight AJAX. Always waits at least one poll cycle so it also
    catches loaders that appear a moment AFTER the action that triggered them."""
    deadline = time.time() + timeout
    # Give the UI a brief moment to actually raise a loader in response to the
    # previous action before we start testing it.
    time.sleep(0.4)
    while time.time() < deadline:
        try:
            if not _page_is_busy(driver):
                return True
        except Exception:
            pass
        time.sleep(0.3)
    _emit("[WARN] Loader did not clear within timeout; continuing.")
    status = time.time() - (deadline - timeout)
    return False


def wait_until_idle(driver, wait=None, timeout=20, step_name="", quiet=0.5):
    """Waits for the loader AND then confirms the page stays idle for `quiet`
    seconds - so we never click during a brief gap between two loads."""
    wait_for_sage_loader(driver, timeout=timeout)
    last_busy = time.time()
    deadline = time.time() + timeout
    while time.time() < deadline:
        busy = False
        try:
            busy = _page_is_busy(driver)
        except Exception:
            busy = False
        if busy:
            last_busy = time.time()
        elif time.time() - last_busy >= quiet:
            return True
        time.sleep(0.2)
    return False


def wait_for_page_ready(driver, timeout=30):
    """Waits until the screen fully loads after a transition: document ready,
    then all Sage loaders (including the busy cursor) cleared and idle."""
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )
    except Exception:
        pass
    wait_until_idle(driver, None, timeout=timeout)
    if STEP_DELAY > 0:
        time.sleep(STEP_DELAY)


def wait_and_click(driver, wait, by, value, step_name="click", record=True):
    """Waits for ALL loading to finish, ensures visibility, scrolls, highlights,
    pauses, clicks, then waits for the resulting load to finish."""
    try:
        wait_until_idle(driver, wait, step_name=step_name)

        element = wait.until(EC.element_to_be_clickable((by, value)))
        try:
            driver.execute_script("arguments[0].scrollIntoView({block: 'center', inline: 'nearest'});", element)
        except Exception:
            pass

        try:
            highlight_element(driver, element)
        except Exception:
            pass
        if STEP_DELAY > 0:
            time.sleep(STEP_DELAY)

        retries = 3
        while retries > 0:
            try:
                wait_until_idle(driver, wait, step_name=step_name)
                element.click()
                break
            except (StaleElementReferenceException, Exception):
                retries -= 1
                time.sleep(0.2)
                if retries == 0:
                    try:
                        driver.execute_script("arguments[0].click();", element)
                    except Exception:
                        raise

        # Wait for the click's resulting load to fully settle before returning.
        wait_until_idle(driver, wait, step_name=step_name)
        if STEP_DELAY > 0:
            time.sleep(STEP_DELAY)
        try:
            unhighlight_element(driver, element)
        except Exception:
            pass
        log_step(step_name, "OK")
        return element
    except Exception as e:
        log_step(step_name, "FAIL" if record else "RETRY", f"{type(e).__name__}: {e}")
        raise


def wait_and_fill(driver, wait, by, value, text_to_send, step_name="fill"):
    """Waits for ALL loading to finish, then highlights, clears and inputs text.
    Waits again for any load triggered by committing the value."""
    try:
        wait_until_idle(driver, wait, step_name=step_name)

        element = wait.until(EC.element_to_be_clickable((by, value)))
        try:
            driver.execute_script("arguments[0].scrollIntoView({block: 'center', inline: 'nearest'});", element)
        except Exception:
            pass

        try:
            highlight_element(driver, element)
        except Exception:
            pass
        if STEP_DELAY > 0:
            time.sleep(STEP_DELAY)

        element.click()
        element.send_keys(Keys.CONTROL, "a")
        if STEP_DELAY > 0:
            time.sleep(0.1)
        element.send_keys(Keys.BACKSPACE)
        if STEP_DELAY > 0:
            time.sleep(0.1)

        if text_to_send:
            element.send_keys(text_to_send)

        if STEP_DELAY > 0:
            time.sleep(STEP_DELAY)
        element.send_keys(Keys.TAB)

        # Committing a field can trigger a lookup/validation load - wait it out.
        wait_until_idle(driver, wait, step_name=step_name)
        try:
            unhighlight_element(driver, element)
        except Exception:
            pass
        log_step(step_name, "OK")
        return element
    except Exception as e:
        log_step(step_name, "FAIL", f"{type(e).__name__}: {e}")
        raise


def wait_click_and_down_arrow(driver, wait, by, value, step_name="select via DOWN arrow"):
    """Waits for ALL loading to finish, clicks, hits DOWN arrow, waits again."""
    try:
        wait_until_idle(driver, wait, step_name=step_name)

        wait.until(EC.visibility_of_element_located((by, value)))
        element = wait.until(EC.element_to_be_clickable((by, value)))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center', inline: 'nearest'});", element)

        highlight_element(driver, element)
        time.sleep(STEP_DELAY)

        element.click()
        time.sleep(0.8)
        element.send_keys(Keys.ARROW_DOWN)

        wait_until_idle(driver, wait, step_name=step_name)
        time.sleep(STEP_DELAY)
        unhighlight_element(driver, element)
        log_step(step_name, "OK")
        return element
    except Exception as e:
        log_step(step_name, "FAIL", f"{type(e).__name__}: {e}")
        raise


def click_first_match(driver, wait, candidates, step_name):
    """Tries each xpath in order (row 1, then row 2...); stops at the first that works."""
    errors = []
    for xpath in candidates:
        try:
            wait_and_click(driver, wait, By.XPATH, xpath, step_name=step_name, record=False)
            return xpath
        except Exception as e:
            errors.append(f"{xpath} -> {type(e).__name__}: {e}")
            _emit(f"[RETRY] {step_name}: candidate failed, moving to next -> {xpath}")
            time.sleep(STEP_DELAY)
    err = TimeoutError("All candidates failed: " + " | ".join(errors))
    log_step(step_name, "FAIL", str(err))
    raise err


ZAPAY_RE = re.compile(r"Creation of ZAPAY\s*[:\-]?\s*([A-Za-z0-9]+)", re.IGNORECASE)

ALL_INPUT_VALUES_JS = """
const clean = (s) => (s || '').replace(/\\s+/g, ' ').trim();
return Array.from(document.querySelectorAll('input, textarea'))
    .filter(el => (el.value || '').trim() !== '')
    .map(el => ({ id: el.id || '', value: clean(el.value) }));
"""


def scan_for_pay_number(driver, source=""):
    """Searches log rows, cells, input values and body text for the PayNumber."""
    for xpath, label in (
        ("//tr[contains(., 'Creation of ZAPAY')]", "log row"),
        ("//*[contains(text(), 'Creation of ZAPAY')]", "cell"),
    ):
        try:
            els = driver.find_elements(By.XPATH, xpath)
        except Exception:
            els = []
        for i, el in enumerate(els, 1):
            try:
                text = el.text
            except Exception:
                _emit(f"[RETRY] {label} {i} went stale, moving to {label} {i + 1}...")
                continue
            m = ZAPAY_RE.search(text or "")
            if m:
                pay = m.group(1).strip()
                _emit(f"Extracted PayNumber via {label} {i} {source}: {pay}")
                log_step("Extract ZAPAY from log", "OK", f"{label} {i} {source}: {pay}")
                return pay
            _emit(f"[RETRY] {label} {i} has no PayNumber -> '{(text or '')[:110]}', moving to {label} {i + 1}...")

    try:
        inputs = driver.execute_script(ALL_INPUT_VALUES_JS) or []
    except Exception:
        inputs = []
    log_lines = [i for i in inputs if str(i.get("id", "")).startswith("4-")]
    preview = " || ".join(i["value"][:90] for i in log_lines[:6])
    _emit(f"[LOG] {len(log_lines)} non-empty log line(s) {source}: {preview}")
    for item in inputs:
        m = ZAPAY_RE.search(item.get("value", ""))
        if m:
            pay = m.group(1).strip()
            _emit(f"Extracted PayNumber via input {item.get('id')} {source}: {pay}")
            log_step("Extract ZAPAY from log", "OK", f"input {item.get('id')} {source}: {pay}")
            return pay

    try:
        body_text = driver.find_element(By.TAG_NAME, "body").text or ""
    except Exception:
        body_text = ""
    m = ZAPAY_RE.search(body_text)
    if m:
        pay = m.group(1).strip()
        _emit(f"Extracted PayNumber via body dump {source}: {pay}")
        log_step("Extract ZAPAY from log", "OK", f"body {source}: {pay}")
        return pay
    return None


def extract_pay_number(driver, timeout=60):
    """Waits for the Log Reading screen, then clicks through the log views
    (header -> Next error/warning -> result rows 1..12) until a PayNumber shows up."""
    _emit("Waiting for Log Reading report to generate...")
    wait_for_page_ready(driver)

    header_xpath = "//*[contains(., 'Log Reading') or contains(., 'Creation of ZAPAY')]"
    found = False
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            if driver.find_elements(By.XPATH, header_xpath):
                found = True
                break
        except Exception:
            pass
        wait_for_sage_loader(driver)
        time.sleep(1)

    if not found:
        raise TimeoutError(f"Log Reading screen never appeared within {timeout}s.")

    time.sleep(1.5)
    wait_for_sage_loader(driver)
    pay_number = scan_for_pay_number(driver, "(first look)")
    if pay_number:
        return pay_number

    # Messages stay hidden until the log view is interacted with - click through it
    attempts = [
        ("log header cell",
         "//div[contains(@class,'s-grid-table-head')]//div[contains(@class,'s-grid-title-cell-label')]"),
        ("Next error link", "//a[contains(normalize-space(.), 'Next error')]"),
        ("Next warning link", "//a[contains(normalize-space(.), 'Next warning')]"),
    ]
    attempts += [
        (f"result row {n}", f"//td[contains(@class,'s-grid-row-index')][normalize-space()='{n}']/parent::tr")
        for n in range(1, 8)
    ]

    for label, xpath in attempts:
        _emit(f"[LOG] Trying to reveal messages: {label}...")
        try:
            els = driver.find_elements(By.XPATH, xpath)
            if not els:
                _emit(f"[LOG] {label} not present, skipping.")
                continue
            try:
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", els[0])
            except Exception:
                pass
            if STEP_DELAY > 0:
                time.sleep(STEP_DELAY)
            try:
                highlight_element(driver, els[0])
            except Exception:
                pass
            els[0].click()
        except Exception as e:
            _emit(f"[LOG] {label} could not be clicked ({type(e).__name__}: {e}).")
            continue

        time.sleep(1.5)
        wait_for_sage_loader(driver)
        pay_number = scan_for_pay_number(driver, f"(after {label})")
        if pay_number:
            return pay_number

    raise ValueError(
        "Log Reading opened but 'Creation of ZAPAY' was never found after clicking through "
        "header, Next error/warning and result rows 1-12."
    )


def open_payment_screen(driver, wait):
    """Navigates the Sage sitemap to the A/P-A-R Payment function."""
    wait_and_click(driver, wait, By.XPATH, '//*[@id="navigationPageHeaderLink"]/I', step_name="Open sitemap")
    wait_and_click(driver, wait, By.XPATH, '//*[@id="header_sitemap_mod_7_bar"]',
                   step_name="Open A/P-A-R accounting module")
    wait_and_click(driver, wait, By.XPATH, '//*[@id="header_sitemap_mod_7_sub_1_mn_117"]',
                   step_name="Open Payment menu")
    wait_click_and_down_arrow(driver, wait, By.XPATH, '//*[@id="3-48-input"]',
                              step_name="Select Payment By (DOWN arrow)")
    wait_and_click(driver, wait, By.XPATH, '//*[@id="3-49-input"]', step_name="Click Payment Type field")
    wait_click_and_down_arrow(driver, wait, By.XPATH, '//*[@id="3-49-input"]',
                              step_name="Select Payment Type (DOWN arrow)")


def login_to_sage(driver, wait, cfg):
    """Login to Sage X3."""
    login_url = "https://sagex3v12.tiauto.co.za/"
    step(f"Navigate to {login_url}")
    driver.get(login_url)
    wait_for_page_ready(driver)

    step("Fill Login (username)")
    wait_and_fill(driver, wait, By.XPATH, '//*[@id="login"]', cfg["username"],
                  step_name="Fill Login (username)")

    step("Fill Password")
    wait_and_fill(driver, wait, By.XPATH, '//*[@id="password"]', cfg["password"],
                  step_name="Fill Password")

    step("Click Sign In button")
    wait_and_click(driver, wait, By.XPATH, '//*[@id="go-basic"]', step_name="Click Sign In button")
    wait_for_page_ready(driver)


def logout_from_sage(driver, wait):
    """Logout from Sage X3."""
    try:
        step("Logging out")
        # Try common logout patterns
        logout_candidates = [
            '//*[@id="header-user-menu"]',
            '//*[@id="user-menu"]',
            "//a[contains(@href, 'logout')]",
            "//button[contains(text(), 'Logout')]",
            "//a[contains(text(), 'Log out')]",
            "//a[contains(text(), 'Sign out')]",
        ]
        for xpath in logout_candidates:
            try:
                els = driver.find_elements(By.XPATH, xpath)
                if els and els[0].is_displayed():
                    els[0].click()
                    time.sleep(1)
                    break
            except Exception:
                continue
        # Try clicking logout in menu
        logout_menu_candidates = [
            "//a[contains(text(), 'Logout')]",
            "//a[contains(text(), 'Log out')]",
            "//a[contains(text(), 'Sign out')]",
            "//div[contains(text(), 'Logout')]",
        ]
        for xpath in logout_menu_candidates:
            try:
                els = driver.find_elements(By.XPATH, xpath)
                if els and els[0].is_displayed():
                    els[0].click()
                    time.sleep(1)
                    break
            except Exception:
                continue
        # Force navigate to login or clear session
        driver.delete_all_cookies()
        driver.get("https://sagex3v12.tiauto.co.za/")
        wait_for_page_ready(driver)
        time.sleep(2)
        log_step("Logout", "OK")
    except Exception as e:
        log_step("Logout", "FAIL", f"{type(e).__name__}: {e}")
        try:
            driver.delete_all_cookies()
            driver.get("https://sagex3v12.tiauto.co.za/")
        except Exception:
            pass


def process_row(driver, wait, row, cfg, excel_path):
    """Runs the full payment flow on the currently open Payment screen for one row."""
    row_number = row["row"]
    supplier_code = row["supplier"]
    company_code = cfg["company_code"]
    bank_code = cfg["bank_code"]
    payment_date = cfg["payment_date"]

    step(f"Row {row_number} - Fill Company Code: {company_code}")
    wait_and_click(driver, wait, By.XPATH, '//*[@id="3-50-input"]', step_name="Click Company Code field")
    wait_and_fill(driver, wait, By.XPATH, '//*[@id="3-50-input"]', company_code,
                  step_name="Fill Company Code")

    step(f"Row {row_number} - Clear Remittance User field")
    remittance_xpath = '/html/body/div/article/div[3]/div/div/div/div/article/div[2]/article/div[1]/div[2]/div/div/div/section/div/div/div/section/div/div/div[1]/section/div/div/div/section/div/div/div[6]/div/div/div/input'
    wait_and_click(driver, wait, By.XPATH, remittance_xpath, step_name="Click Remittance User field")
    wait_and_fill(driver, wait, By.XPATH, remittance_xpath, "", step_name="Clear Remittance User")

    step(f"Row {row_number} - Fill Supplier Code: {supplier_code}")
    supplier_xpath = '/html/body/div/article/div[3]/div/div/div/div/article/div[2]/article/div[1]/div[2]/div/div/div/section/div/div/div/section/div/div/div[1]/section/div/div/div/section/div/div/div[9]/div/div/div[1]/input'
    wait_and_click(driver, wait, By.XPATH, supplier_xpath, step_name="Click Supplier Code field")
    wait_and_fill(driver, wait, By.XPATH, supplier_xpath, supplier_code, step_name="Fill Supplier Code")

    step(f"Row {row_number} - Fill Bank Code: {bank_code}")
    bank_xpath = '/html/body/div/article/div[3]/div/div/div/div/article/div[2]/article/div[1]/div[2]/div/div/div/section/div/div/div/section/div/div/div[2]/section/div/div/div/section/div/div/div[1]/div/div/div[1]/input'
    wait_and_click(driver, wait, By.XPATH, bank_xpath, step_name="Click Bank Code field")
    wait_and_fill(driver, wait, By.XPATH, bank_xpath, bank_code, step_name="Fill Bank Code")

    step(f"Row {row_number} - Fill Payment Date: {payment_date}")
    date_xpath = '/html/body/div/article/div[3]/div/div/div/div/article/div[2]/article/div[1]/div[2]/div/div/div/section/div/div/div/section/div/div/div[2]/section/div/div/div/section/div/div/div[2]/div/div/div/input'
    wait_and_click(driver, wait, By.XPATH, date_xpath, step_name="Click Payment Date field")
    wait_and_fill(driver, wait, By.XPATH, date_xpath, payment_date, step_name="Fill Payment Date")

    step(f"Row {row_number} - Click Refresh button")
    refresh_btn_xpath = '/html/body/div/article/div[3]/div/div/div/div/article/div[2]/article/div[1]/header/div/div[2]/div[1]/a[3]'
    wait_and_click(driver, wait, By.XPATH, refresh_btn_xpath, step_name="Click Refresh button")
    wait_for_page_ready(driver)

    step(f"Row {row_number} - Click [Pay the Rest] link")
    pay_link_candidates = [
        f'//*[@id="layout-slot-3-71"]/div/article/div[2]/div[1]/div/div[1]/div[2]/table/tr[{r}]/td[2]/a'
        for r in range(1, 4)
    ]
    pay_link_candidates += [
        f'//*[@id="layout-slot-3-71"]//table/tbody/tr[{r}]/td[2]/a' for r in range(1, 4)
    ]
    click_first_match(driver, wait, pay_link_candidates, step_name="Click [Pay the Rest] link")

    time.sleep(1)
    wait_for_page_ready(driver)

    step(f"Row {row_number} - Click [Pay the Rest] option")
    pay_the_rest_candidates = [
        '/html/body/div/article/div[3]/div[2]/div/a[4]',
        '/html/body/div/article/div[3]/div[2]/div/a[3]',
        '/html/body/div/article/div[3]/div[2]/div/a[2]',
        '/html/body/div/article/div[3]/div[2]/div/a[1]',
        "//*[contains(@class,'menu')]//a[contains(normalize-space(), 'Pay the rest') or contains(normalize-space(), 'Pay the Rest')]",
    ]
    click_first_match(driver, wait, pay_the_rest_candidates, step_name="Click [Pay the Rest] option")

    time.sleep(0.8)
    wait_for_page_ready(driver)

    step(f"Row {row_number} - Click PAY button")
    pay_btn_xpath = '/html/body/div/article/div[3]/div/div/div/div/article/div[2]/article/div[1]/header/div/div[2]/div[1]/a[2]'
    wait_and_click(driver, wait, By.XPATH, pay_btn_xpath, step_name="Click PAY button")

    step(f"Row {row_number} - Click Yes on confirmation popup")
    pay_popup_candidates = [
        '/html/body/div/article/div[3]/div[2]/div/article/div[2]/a[1]',
        "//*[contains(@class,'popup') or contains(@class,'modal')]//a[normalize-space()='Yes' or normalize-space()='YES']",
        '/html/body/div/article/div[3]/div[2]/div/article/div[2]/a[2]',
        '/html/body/div/article/div[3]/div[2]/div/a[1]',
    ]
    click_first_match(driver, wait, pay_popup_candidates, step_name="Click Yes on confirmation popup")

    step(f"Row {row_number} - Wait for result screen to load")
    wait_for_page_ready(driver, timeout=40)

    step(f"Row {row_number} - Wait for Log Reading & extract PayNumber")
    pay_number = extract_pay_number(driver, timeout=60)
    _emit(f"Generated PayNumber: {pay_number}")

    step(f"Row {row_number} - Save PayNumber")
    write_pay_number_to_excel(excel_path, row_number, pay_number, supplier=supplier_code)

    step(f"Row {row_number} - Exit")
    try:
        wait_and_click(driver, wait, By.XPATH,
                       '/html/body/div/article/div[1]/div/article/div[1]/header/div[1]/div[2]/div[3]/a/div',
                       step_name="Exit")
    except Exception as e:
        log_step("Exit", "FAIL", f"{type(e).__name__}: {e}")

    step(f"Row {row_number} - Cancel/Close")
    try:
        wait_and_click(driver, wait, By.XPATH,
                       '/html/body/div/article/div[3]/div/div/div/div/article/div[2]/article/div[1]/header/div/div[2]/a/I',
                       step_name="Cancel (X)")
    except Exception as e:
        log_step("Cancel (X)", "FAIL", f"{type(e).__name__}: {e}")
    wait_for_page_ready(driver)


def run(excel_path=None, headless=None):
    global file_path, HEADLESS
    if excel_path:
        file_path = excel_path
    if headless is None:
        headless = HEADLESS
    HEADLESS = bool(headless)

    cfg = read_parameters(file_path)
    _emit(f"Loaded credentials for user: {cfg['username']}")
    _emit(f"Target Parameters -> Company: {cfg['company_code']}, "
          f"Bank: {cfg['bank_code']}, Date: {cfg['payment_date']}")

    accounts = read_accounts(file_path)
    if not accounts:
        _emit("[INFO] No data rows found in Account_Details.")
        return {"processed": 0, "failed": 0, "skipped": 0, "stopped": False}

    todo = [r for r in accounts if not r["pay_number"]]
    skipped = [r for r in accounts if r["pay_number"]]
    for r in skipped:
        _emit(f"[SKIP] Row {r['row']} (supplier={r['supplier']}, vendor={r['vendor_name']}) "
              f"already has PayNumber '{r['pay_number']}' -> skipping.")
        log_step("Skip row (already paid)", "OK",
                 f"row {r['row']} supplier={r['supplier']} pay_number={r['pay_number']}")

    if not todo:
        _emit("[INFO] Every row already has a PayNumber. Nothing to process.")
        return {"processed": 0, "failed": 0, "skipped": len(skipped), "stopped": False}

    chrome_options = Options()
    chrome_options.add_argument("--window-size=1600,950")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_argument("--disable-extensions")
    chrome_options.add_argument("--disable-infobars")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-logging", "enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)
    if HEADLESS:
        chrome_options.add_argument("--headless=new")
        _emit("[INFO] Starting Chrome in HEADLESS mode.")
    else:
        _emit("[INFO] Starting Chrome in normal (visible) mode.")

    global CURRENT_DRIVER
    failed_rows = 0
    processed = 0
    stopped = False
    try:

        for row in todo:
            if STOP_FLAG is not None and STOP_FLAG():
                stopped = True
                _emit("[STOP] Stop requested. Halting before next row.")
                break
            _emit(f"\n===== Processing Row {row['row']} | supplier={row['supplier']} | "
                  f"vendor={row['vendor_name']} | company={cfg['company_code']} "
                  f"| bank={cfg['bank_code']} =====")
            log_step("Start row", "OK",
                     f"row {row['row']} supplier={row['supplier']} "
                     f"company={cfg['company_code']} bank={cfg['bank_code']}")
            driver = None
            try:
                if STOP_FLAG is not None and STOP_FLAG():
                    stopped = True
                    break
                driver = webdriver.Chrome(options=chrome_options)
                CURRENT_DRIVER = driver
                wait = WebDriverWait(driver, 15)
                login_to_sage(driver, wait, cfg)
                if STOP_FLAG is not None and STOP_FLAG():
                    stopped = True
                    break
                step("Open Payment screen")
                open_payment_screen(driver, wait)
                if STOP_FLAG is not None and STOP_FLAG():
                    stopped = True
                    break
                process_row(driver, wait, row, cfg, file_path)
                processed += 1
            except Exception as e:
                if STOP_FLAG is not None and STOP_FLAG():
                    stopped = True
                    break
                failed_rows += 1
                _emit(f"[FAILED] Row {row['row']} not processed: {type(e).__name__}: {e}")
                log_step("Row failed", "FAIL",
                         f"row {row['row']} supplier={row['supplier']} -> {e}")
            finally:
                CURRENT_DRIVER = None
                if driver is not None:
                    try:
                        driver.quit()
                    except Exception:
                        pass
                time.sleep(0.2)

        _emit(f"\nProcess finished. Rows processed: {processed}/{len(todo)}, "
              f"failed: {failed_rows}, skipped: {len(skipped)}"
              + (" (stopped early)." if stopped else "."))
        log_step("Run", "OK" if failed_rows == 0 and not stopped else "PARTIAL",
                 f"processed={processed} failed={failed_rows} skipped={len(skipped)} stopped={stopped}")

    finally:
        CURRENT_DRIVER = None

    return {"processed": processed, "failed": failed_rows,
            "skipped": len(skipped), "stopped": stopped}


if __name__ == "__main__":
    try:
        run()
        print(f"\nStep log: {STEPS_LOG}")
    except Exception as e:
        print(f"\n[FAILED] Step: {CURRENT_STEP}")
        print(f"{type(e).__name__}: {e}")
        print(f"Step log: {STEPS_LOG}")
        sys.exit(1)
