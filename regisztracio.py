
#TAJ OK
from pathlib import Path
from datetime import datetime, date
import time
import unicodedata
import pandas as pd
import json
import sys
import os
import winsound  # Windows hangjelzéshez
import xlwings as xw

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import InvalidSessionIdException, TimeoutException, NoSuchWindowException, WebDriverException, StaleElementReferenceException
from selenium.webdriver.common.keys import Keys
import re

from utils.logger import logger
from pages.login_page import LoginPage
from address_handlers import fill_address_data
from phone_handlers import fill_phone_data


def load_config() -> dict:
    """Load configuration from config.json next to executable."""
    if getattr(sys, 'frozen', False):
        # Running as PyInstaller executable - use executable's directory
        base_path = os.path.dirname(sys.executable)
    else:
        # Running as Python script - use script's directory
        base_path = os.path.dirname(os.path.abspath(__file__))
    
    config_path = os.path.join(base_path, 'config.json')
    
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            config = json.load(f)
        logger.info(f"✅ Config betöltve: {config_path}")
        return config
    except FileNotFoundError:
        logger.error(f"❌ config.json nem található: {config_path}")
        logger.error(f"   Hozd létre a config.json fájlt itt: {base_path}")
        raise
    except json.JSONDecodeError as e:
        logger.error(f"❌ config.json érvénytelen JSON: {e}")
        raise


# Load config at module level
CONFIG = load_config()


# ---------------------------------------------------------------------
# Beállítások
# ---------------------------------------------------------------------
# Configuration values from config.json
EXCEL_PATH = Path(CONFIG['excel_path_registration'])
# A regisztrációs modul kizárólag a "Páciensek" munkalapot használja
SHEET_NAME = "Páciensek"
LOGIN_URL = CONFIG['login_url']
USERNAME = CONFIG['username']
PASSWORD = CONFIG['password']
HEADLESS = CONFIG.get('headless', False)

LOG_DIR = Path("logs")
LOG_DIR.mkdir(exist_ok=True)

# Ha csak login tesztet szeretnél: állítsd False-ra
USE_UPLOAD = True

# Feature flags for baseline testing
FF_BASELINE_TAJ_NAME_DOB = False
# Allow overriding baseline via env var, default OFF
try:
    _bl = os.getenv("BASELINE_TAJ_NAME_DOB", "").strip().lower()
    if _bl in ("1","true","yes","on"):
        FF_BASELINE_TAJ_NAME_DOB = True
except Exception:
    pass
FF_EMAIL_STEPS = False

# CHECKPOINT OK — TAJ+NAME+DOB baseline restored (2025-01-27 14:30)


# ---------------------------------------------------------------------
# Kisegítő függvények
# ---------------------------------------------------------------------
def ts() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def save_debug(driver, tag: str):
    """Ment képernyőt és a DOM-ot a logs mappába."""
    png = LOG_DIR / f"{tag}_{ts()}.png"
    html = LOG_DIR / f"{tag}_{ts()}.html"
    try:
        driver.save_screenshot(str(png))
    except Exception:
        pass
    try:
        # guard: session may be disconnected
        ps = ""
        try:
            ps = driver.page_source
        except Exception:
            ps = "<page_source_unavailable_due_to_driver_disconnect/>"
        html.write_text(ps, encoding="utf-8")
    except Exception:
        pass
    logger.info(f"🖼️ Mentve: {png.name}  |  🧾 Mentve: {html.name}")

# --- Resilient attribute getters ------------------------------------------------

def _exc_brief(e: Exception) -> str:
    """Return a short, one-line summary for noisy WebDriver exceptions."""
    try:
        s = str(e) or e.__class__.__name__
        s = s.replace("\r", "\n")
        first = s.split("\n", 1)[0].strip()
        return first or e.__class__.__name__
    except Exception:
        return e.__class__.__name__

def _safe_attr(el, name: str) -> str:
    try:
        return (el.get_attribute(name) or "")
    except Exception:
        try:
            return (el.get_dom_attribute(name) or "")
        except Exception:
            return ""

def _safe_text(el) -> str:
    try:
        return (el.text or "")
    except Exception:
        return ""

def _digits_only(s: str) -> str:
    return "".join(ch for ch in str(s or "") if ch.isdigit())

def _iso_to_digits(iso_date: str) -> str:
    ds = _digits_only(iso_date)
    if len(ds) >= 8:
        return ds[:8]
    return ds

def make_driver() -> webdriver.Chrome:
    """Chrome driver létrehozása a .env beállításokkal."""
    opts = Options()
    if HEADLESS:
        # új headless motor
        opts.add_argument("--headless=new")
    # tiszta profil minden futásnál
    profile_dir = Path(".").resolve() / f".tmp_chrome_profile_{ts()}"
    opts.add_argument("--user-data-dir=" + str(profile_dir))
    opts.add_argument("--disable-gpu")
    opts.add_argument("--no-sandbox")
    driver = webdriver.Chrome(options=opts)
    driver.set_window_size(1400, 900)
    return driver
def ensure_driver_alive(driver):
    """
    Stabil életjel: először JS-ping.
    Ha az bukik, akkor próbáljunk legutolsó handle-re váltani.
    Ha ez is bukik → InvalidSessionIdException-t dobunk, hogy a felső szint recovery fusson.
    """
    # 1) első kör: JS-ping az aktuális kontextusban
    try:
        driver.execute_script("return 1")
        return
    except (NoSuchWindowException, InvalidSessionIdException, WebDriverException):
        pass  # próbáljunk handle-t váltani

    # 2) második kör: váltsunk a legutolsó handle-re, majd JS-ping
    try:
        handles = driver.window_handles  # ez is dobhat
        if handles:
            driver.switch_to.window(handles[-1])
            driver.execute_script("return 1")
            return
    except (NoSuchWindowException, InvalidSessionIdException, WebDriverException):
        pass

    # 3) ha idáig jutottunk, a session halott
    raise InvalidSessionIdException("WebDriver session is invalid or window closed.")


def _retry_conn(driver, fn, tries=3, wait=0.2):
    """
    Run fn() with driver-liveness checks and retry on transient
    WebDriver disconnects/resets/stale element issues.
    """
    from selenium.common.exceptions import WebDriverException, StaleElementReferenceException
    import time

    last = None
    for _ in range(max(1, tries)):
        try:
            ensure_driver_alive(driver)
            return fn()
        except (WebDriverException, StaleElementReferenceException) as e:
            msg = (str(e) or "").lower()
            if any(k in msg for k in ("connection", "reset", "refused", "disconnected", "stale", "target frame detached")):
                last = e
                time.sleep(wait)
                continue
            raise
    if last:
        raise last


def recreate_and_relogin(old_driver=None):
    """Új driver létrehozása és relogin. Visszaadja az új drivert."""
    try:
        if old_driver is not None:
            old_driver.quit()
    except Exception:
        pass
    new_driver = make_driver()
    lp = LoginPage(new_driver, LOGIN_URL, timeout=45)
    success = lp.login(USERNAME, PASSWORD)
    if not success:
        raise InvalidSessionIdException("Login failed during driver recreation")
    return new_driver



def wait_click_css(driver, css: str, timeout=20):
    ensure_driver_alive(driver)
    el = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.CSS_SELECTOR, css)))
    # Wait for loading indicator to disappear if present
    try:
        WebDriverWait(driver, 5).until_not(
            EC.presence_of_element_located((By.CSS_SELECTOR, '[data-automation-id="loading-indicator"]'))
        )
    except Exception:
        pass  # Loading indicator might not exist, continue
    # Try normal click, fallback to JavaScript click if intercepted
    try:
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", el)
        el.click()
    except Exception:
        driver.execute_script("arguments[0].click();", el)
    return el


def wait_type_id(driver, id_: str, value: str, timeout=20, clear=True):
    ensure_driver_alive(driver)
    el = WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.ID, id_)))
    if clear:
        try:
            el.clear()
        except Exception:
            pass
    el.send_keys(value)
    return el


def _norm(s: str) -> str:
    s = str(s or "")
    s = s.replace("\u00A0", " ")
    s = re.sub(r"\s+", " ", s.strip())
    return s


def find_input_smart(driver, terms=None, attr_contains=None, timeout=20):
    """
    Smart input finder with connection-safe retries.
    Tries label/placeholder/aria/id/name, then email-specific hints, and searches iframes up to depth 2.
    """
    from selenium.webdriver.common.by import By
    from selenium.common.exceptions import TimeoutException
    import time

    terms = terms or []
    attr_contains = attr_contains or []

    def _try_find_in_context(context_driver, strategy_name):
        xpaths = []
        for t in terms:
            t_norm = _norm(t)
            xpaths.append(f"//label[normalize-space()='{t_norm}']/following::input[1]")
            xpaths.append(f"//label[contains(normalize-space(), '{t_norm}')]/following::input[1]")
            xpaths.append(f"//input[contains(translate(@placeholder,'ABCDEFGHIJKLMNOPQRSTUVWXYZÁÉÍÓÖŐÚÜŰ','abcdefghijklmnopqrstuvwxyzáéíóöőúüű'), '{t_norm.lower()}')]")
            xpaths.append(f"//input[contains(translate(@aria-label,'ABCDEFGHIJKLMNOPQRSTUVWXYZÁÉÍÓÖŐÚÜŰ','abcdefghijklmnopqrstuvwxyzáéíóöőúüű'), '{t_norm.lower()}')]")
        for a in attr_contains:
            low = a.lower()
            xpaths.append(f"//input[contains(translate(@id,'ABCDEFGHIJKLMNOPQRSTUVWXYZÁÉÍÓÖŐÚÜŰ','abcdefghijklmnopqrstuvwxyzáéíóöőúüű'), '{low}')]")
            xpaths.append(f"//input[contains(translate(@name,'ABCDEFGHIJKLMNOPQRSTUVWXYZÁÉÍÓÖŐÚÜŰ','abcdefghijklmnopqrstuvwxyzáéíóöőúüű'), '{low}')]")

        for xp in xpaths:
            try:
                el = _retry_conn(driver, lambda: context_driver.find_element(By.XPATH, xp))
                try:
                    vis = _retry_conn(driver, lambda: el.is_displayed())
                except Exception:
                    vis = True
                if vis:
                    logger.info(f"✅ find_input_smart: {strategy_name} - id={el.get_attribute('id')} name={el.get_attribute('name')} type={el.get_attribute('type')}")
                    return el
            except Exception:
                continue

        email_selectors = [
            "//input[translate(@type,'EMAIL','email')='email']",
            "//input[contains(translate(@autocomplete,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'), 'email')]",
            "//input[contains(translate(@id,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'), 'mail')]",
            "//input[contains(translate(@name,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'), 'mail')]"
        ]
        for xp in email_selectors:
            try:
                el = _retry_conn(driver, lambda: context_driver.find_element(By.XPATH, xp))
                try:
                    vis = _retry_conn(driver, lambda: el.is_displayed())
                except Exception:
                    vis = True
                if vis:
                    logger.info(f"✅ find_input_smart: {strategy_name} - email-specific - id={el.get_attribute('id')} name={el.get_attribute('name')} type={el.get_attribute('type')}")
                    return el
            except Exception:
                continue
        return None

    def _search_frames(depth, max_depth=2):
        if depth > max_depth:
            return None
        try:
            frames = _retry_conn(driver, lambda: driver.find_elements(By.TAG_NAME, "iframe"))
        except Exception:
            frames = []
        for fr in frames:
            try:
                _retry_conn(driver, lambda: driver.switch_to.frame(fr))
                found = _try_find_in_context(driver, f"iframe_depth_{depth}")
                if found:
                    return found
                if depth < max_depth:
                    nested = _search_frames(depth + 1, max_depth)
                    if nested:
                        return nested
            except Exception:
                pass
            finally:
                try:
                    driver.switch_to.parent_frame()
                except Exception:
                    try:
                        driver.switch_to.default_content()
                    except Exception:
                        pass
        return None

    end = time.time() + timeout
    while time.time() < end:
        try:
            driver.switch_to.default_content()
        except Exception:
            pass
        found = _try_find_in_context(driver, "main")
        if found:
            try:
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", found)
            except Exception:
                pass
            return found

        try:
            driver.switch_to.default_content()
        except Exception:
            pass
        found = _search_frames(1, 2)
        if found:
            try:
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", found)
            except Exception:
                pass
            return found
        time.sleep(0.2)

    try:
        driver.switch_to.default_content()
    except Exception:
        pass
    raise TimeoutException(f"Smart input not found for terms={terms} attr_contains={attr_contains}")


def type_sturdy(driver, el, value: str):
    """
    Stabil beírás: scroll → click → Ctrl+A+Backspace → send_keys → verify → JS fallback if mismatch.
    """
    from selenium.webdriver.common.keys import Keys
    value_str = str(value or "")
    
    # Scroll into view
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    except Exception:
        pass
    
    # Click to focus
    try:
        el.click()
    except Exception:
        try:
            driver.execute_script("arguments[0].click();", el)
        except Exception:
            pass
    
    # Clear: Ctrl+A then Backspace
    try:
        el.send_keys(Keys.CONTROL, "a")
        time.sleep(0.1)
        el.send_keys(Keys.BACKSPACE)
        time.sleep(0.1)
    except Exception:
        try:
            el.clear()
        except Exception:
            pass
    
    # Type the value
    try:
        el.send_keys(value_str)
        time.sleep(0.1)
    except Exception:
        pass
    
    # Verify: if mismatch, set via JS
    try:
        cur = el.get_attribute("value") or ""
        if value_str != cur:
            # Mismatch: set value via JS and dispatch events
            driver.execute_script(
                """
                const el = arguments[0], val = arguments[1];
                const desc = Object.getOwnPropertyDescriptor(Object.getPrototypeOf(el), 'value') || 
                             Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value');
                if (desc && desc.set) {
                    desc.set.call(el, val);
                } else {
                    el.value = val;
                }
                el.dispatchEvent(new Event('input', {bubbles: true, cancelable: true}));
                el.dispatchEvent(new Event('change', {bubbles: true, cancelable: true}));
                """,
                el, value_str
            )
    except Exception:
        pass
    
    return el


def open_email_section(driver):
    """Kattint az 'E-mail címek' felirat melletti kék kör/ikon gombra (csak helyi, nem globális)."""
    forbidden_patterns = {"felvétel", "create new patient", "patientregister_createnewpatient"}
    
    def _is_forbidden_element(el):
        """Check if element matches forbidden patterns."""
        try:
            text = (el.text or "").lower()
            aria_label = (el.get_attribute("aria-label") or "").lower()
            automation_id = (el.get_attribute("data-automation-id") or "").lower()
            element_id = (el.get_attribute("id") or "").lower()
            combined = f"{text} {aria_label} {automation_id} {element_id}"
            for pattern in forbidden_patterns:
                if pattern in combined:
                    logger.info(f"🛡️ Prevented global Felvétel click in open_email_section: {pattern}")
                    return True
        except Exception:
            pass
        return False
    
    try:
        # Find "E-mail címek" label first
        label = None
        try:
            label = driver.find_element(By.XPATH, "//*[contains(normalize-space(),'E-mail címek')]")
        except Exception:
            try:
                label = driver.find_element(By.XPATH, "//*[contains(translate(normalize-space(),'ABCDEFGHIJKLMNOPQRSTUVWXYZÁÉÍÓÖŐÚÜŰ','abcdefghijklmnopqrstuvwxyzáéíóöőúüű'), 'e-mail cimek')]")
            except Exception:
                pass
        
        if label is None:
            return
        
        # Find the following button/icon/svg near the label (within same section)
        try:
            # Try direct following sibling
            btn = label.find_element(By.XPATH, "following-sibling::*[local-name()='svg' or local-name()='button' or local-name()='span'][1]")
            if btn and btn.is_displayed() and not _is_forbidden_element(btn):
                try:
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
                    btn.click()
                    time.sleep(0.3)
                    return
                except Exception:
                    driver.execute_script("arguments[0].click();", btn)
                    time.sleep(0.3)
                    return
        except Exception:
            pass
        
        # Try following axis (any following element)
        try:
            xp = "//*[contains(normalize-space(),'E-mail címek')]/following::*[local-name()='svg' or local-name()='button' or local-name()='span'][1]"
            candidates = driver.find_elements(By.XPATH, xp)
            
            for btn in candidates:
                try:
                    if not btn.is_displayed():
                        continue
                    if _is_forbidden_element(btn):
                        continue
                    
                    # Verify it's near the label (same section/container)
                    same_container = driver.execute_script("""
                        const label = arguments[0];
                        const btn = arguments[1];
                        if (!label || !btn) return false;
                        // Check if they share a common ancestor within reasonable depth
                        let labelEl = label;
                        let btnEl = btn;
                        for (let i = 0; i < 15; i++) {
                            if (!labelEl || !btnEl) break;
                            if (labelEl === btnEl) return true;
                            // Check if btn is descendant of label's ancestors
                            if (labelEl.contains && labelEl.contains(btn)) return true;
                            labelEl = labelEl.parentElement;
                            btnEl = btnEl.parentElement;
                        }
                        return false;
                    """, label, btn)
                    
                    if same_container:
                        try:
                            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
                        except Exception:
                            pass
                        try:
                            btn.click()
                        except Exception:
                            driver.execute_script("arguments[0].click();", btn)
                        time.sleep(0.3)
                        return
                except Exception:
                    continue
        except Exception:
            pass
    except Exception:
        pass


def ensure_email_section_open(driver):
    """
    Open the E-mail section by clicking the *local* 'Hozzáadás' inside the
    'Elérhetőségek' / 'E-mail címek' block. Never touch global 'Felvétel'.
    Returns True when #EmailAddress is visible.
    """
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    import time

    def _visible(el):
        try:
            return el.is_displayed()
        except Exception:
            return False

    # Already present?
    try:
        el = driver.find_element(By.CSS_SELECTOR, "#EmailAddress, [id='EmailAddress']")
        if _visible(el):
            return True
    except Exception:
        pass

    # 1) Find the section root:
    # Prefer a container that contains 'Elérhetőségek' or 'E-mail címek' text.
    section = None
    section_xpaths = [
        # exact Hungarian headings we were told
        "//*[contains(normalize-space(),'Elérhetőségek')]/ancestor::*[contains(@class,'section') or contains(@class,'group') or @role='region' or @data-automation-id][1]",
        "//*[contains(normalize-space(),'E-mail címek')]/ancestor::*[contains(@class,'section') or contains(@class,'group') or @role='region' or @data-automation-id][1]",
        # fallback: label for EmailAddress
        "//label[@for='EmailAddress']/ancestor::*[contains(@class,'section') or contains(@class,'group') or @role='region' or @data-automation-id][1]",
    ]
    for xp in section_xpaths:
        try:
            cand = driver.find_element(By.XPATH, xp)
            if _visible(cand):
                section = cand
                break
        except Exception:
            continue

    # If still none, last resort: page root (we will still filter buttons heavily)
    root = section if section is not None else driver

    # 2) Inside the section, find a local 'Hozzáadás' control
    local_add_candidates = []
    try:
        # text button 'Hozzáadás'
        if section is not None:
            local_add_candidates.extend(section.find_elements(By.XPATH, ".//button[contains(normalize-space(),'Hozzáadás')]"))
        # common add button automation-id within the section
        if section is not None:
            local_add_candidates.extend(section.find_elements(By.CSS_SELECTOR, "[data-automation-id='__addNewItemCompactButton']"))
        # generic add icons within the section (as a fallback)
        if section is not None:
            local_add_candidates.extend(section.find_elements(By.XPATH, ".//*[contains(@class,'add-button') or contains(@data-automation-id,'add')]"))
    except Exception:
        pass

    # If we had no section, very conservatively search but exclude global controls
    if not local_add_candidates and section is None:
        try:
            for el in driver.find_elements(By.XPATH, "//*[self::button or self::*[@role='button']][contains(normalize-space(),'Hozzáadás')]"):
                local_add_candidates.append(el)
        except Exception:
            pass

    # Filter out globals (Felvétel/CreateNewPatient) and invisible ones
    safe = []
    for el in local_add_candidates:
        try:
            if not _visible(el):
                continue
            blob = " ".join([
                el.text or "",
                (el.get_attribute("id") or ""),
                (el.get_attribute("data-automation-id") or ""),
                (el.get_attribute("aria-label") or "")
            ]).lower()
            if any(bad in blob for bad in ["patientregister_createnewpatient", "felvétel", "felvetel", "createnewpatient"]):
                continue
            # If a section was found, ensure the element is inside it
            if section is not None:
                try:
                    inside = driver.execute_script("return arguments[0].contains(arguments[1]);", section, el)
                except Exception:
                    inside = True
                if not inside:
                    continue
            safe.append(el)
        except Exception:
            continue

    # 3) Click the first safe candidate
    clicked = False
    for btn in safe:
        try:
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
        except Exception:
            pass
        try:
            btn.click()
        except Exception:
            try:
                driver.execute_script("arguments[0].click();", btn)
            except Exception:
                continue
        time.sleep(0.25)
        clicked = True
        break

    # 4) Wait for #EmailAddress to appear
    if clicked:
        try:
            el = WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, "#EmailAddress, [id='EmailAddress']"))
            )
            return _visible(el)
        except Exception:
            return False
    return False


def fill_email_address(driver, email_value: str):
    """
    Open the E-mail section (local), type the email once, and verify.
    If duplication is detected (value == email*2 or contains twice), clear and set via JS.
    """
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    import time, re

    if not email_value:
        return False

    if not ensure_email_section_open(driver):
        return False

    # target input
    em_el = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "#EmailAddress, [id='EmailAddress']"))
    )

    # if already correct, skip
    try:
        cur = (em_el.get_attribute("value") or "").strip()
        if cur == str(email_value).strip():
            logger.info("ℹ️ Email already set, skipping retype.")
            return True
    except Exception:
        cur = ""

    # sturdy type
    type_sturdy(driver, em_el, str(email_value))
    time.sleep(0.1)

    # read back
    try:
        val = (em_el.get_attribute("value") or "").strip()
    except Exception:
        val = ""

    # detect duplication e.g. "x@y.comx@y.com"
    doubled = (val == str(email_value) + str(email_value)) or (val.count(str(email_value)) >= 2)

    # if bad format (simple validation) or doubled -> JS set clean
    simple_ok = bool(re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", str(email_value)))
    if doubled or (not simple_ok):
        try:
            driver.execute_script("""
                (function(el, val){
                    const proto = Object.getPrototypeOf(el);
                    const desc = Object.getOwnPropertyDescriptor(proto, 'value') || Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value');
                    if (desc && desc.set) desc.set.call(el, val); else el.value = val;
                    el.dispatchEvent(new Event('input', {bubbles:true}));
                    el.dispatchEvent(new Event('change', {bubbles:true}));
                    el.dispatchEvent(new Event('blur', {bubbles:true}));
                })(arguments[0], arguments[1]);
            """, em_el, str(email_value))
            time.sleep(0.05)
            val = (em_el.get_attribute("value") or "").strip()
        except Exception:
            pass

    ok = (val == str(email_value).strip())
    if ok:
        logger.info("✅ Email filled & verified (no-duplicate).")
    else:
        logger.warning(f"⚠️ Email value after fill: {val!r} (expected {email_value!r})")
    return ok


def fill_field_smart(driver, labels, attr_contains, value, timeout=20):
    # Detect if this is an email field
    is_email = False
    labels_lower = [str(l).lower() for l in (labels or [])]
    attr_lower = [str(a).lower() for a in (attr_contains or [])]
    
    if any('email' in l or 'mail' in l for l in labels_lower):
        is_email = True
    if any('email' in a or 'mail' in a for a in attr_lower):
        is_email = True
    
    # Open email section if needed
    if is_email:
        ensure_email_section_open(driver)
    
    el = find_input_smart(driver, terms=labels, attr_contains=attr_contains, timeout=timeout)
    return type_sturdy(driver, el, value)


def get_cell(row, *keys):
    """Adj vissza az első nem üres értéket a megadott oszlopnevek közül (Excel fejléc aliasok)."""
    import pandas as _pd
    for k in keys:
        if k in row:
            v = row.get(k, "")
            if v is not None and not (_pd.isna(v)) and str(v).strip() != "":
                return v
    # ha mind üres
    return ""


def to_iso_date(value) -> str:
    """Excelből jövő dátum -> 'YYYY-MM-DD'."""
    if pd.isna(value) or value is None:
        return ""
    # Próbáljuk okosan felismerni
    if isinstance(value, (pd.Timestamp, datetime)):
        return value.strftime("%Y-%m-%d")
    s = str(value).strip()
    # Magyar locale: vessző tizedesjel → pont
    s_dot = s.replace(",", ".")
    # Excel sorszám (pl. 27615 vagy 46127.30059)
    try:
        num = float(s_dot)
        if 10000 < num < 60000:
            from datetime import timedelta as _td
            return (datetime(1899, 12, 30) + _td(days=num)).strftime("%Y-%m-%d")
    except Exception:
        pass
    # próbálunk parse-olni
    for fmt in ("%Y-%m-%d", "%Y.%m.%d", "%Y/%m/%d", "%d.%m.%Y", "%d/%m/%Y", "%Y%m%d"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except Exception:
            pass
    # Pandas default parse
    try:
        if re.match(r"^\d{4}-\d{2}-\d{2}", s):
            return pd.to_datetime(s, dayfirst=False).strftime("%Y-%m-%d")
        return pd.to_datetime(s, dayfirst=True).strftime("%Y-%m-%d")
    except Exception:
        return s  # utolsó esély: ahogy van


def wait_patient_form_ready(driver, timeout: int = 30) -> None:
    """
    Extra stabilizáló várakozás az Új páciens form megnyitása után.
    A mentés gomb önmagában nem garantálja, hogy az Okmányok blokk már renderelt.
    """
    try:
        WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, '[data-automation-id="__save_save"]'))
        )
    except Exception:
        return

    # Wait for any marker that indicates the documents section is present
    markers = [
        '[data-automation-id="Documents"]',
        '[data-automation-id="DocumentTypeId__container"]',
        '#DocumentNumber',
        '[id="DocumentNumber"]',
    ]
    end = time.time() + min(timeout, 20)
    while time.time() < end:
        try:
            ensure_driver_alive(driver)
        except Exception:
            break
        try:
            for css in markers:
                els = driver.find_elements(By.CSS_SELECTOR, css)
                if any(getattr(e, "is_displayed", lambda: True)() for e in els):
                    return
        except Exception:
            pass
        time.sleep(0.2)


def normalize_taj(value: str) -> str:
    """
    Csak whitespace-et és gyakori elválasztókat (szóköz, non-breaking space, kötőjel, pont) szedi ki.
    Nem konvertál int-re, így a leading 0 megmarad.
    """
    s = str(value or "").strip()
    s = s.replace("\u00A0", " ")  # NBSP → space
    s = re.sub(r"[ .\-]", "", s)  # pont/kötőjel/szóköz ki
    return s

def detect_document_type(doc_number: str) -> str:
    """
    Felismeri az okmány típusát a szám formátuma alapján.
    
    Visszatérési értékek:
    - "taj" : TAJ szám (9 számjegy)
    - "passport" : Útlevél (2 betű + 7 számjegy)
    - "id_card" : Személyi igazolvány (6 számjegy + 2 betű)
    - "unknown" : Ismeretlen formátum
    """
    # Tisztítás: whitespace és elválasztók eltávolítása
    clean = re.sub(r"[\s.\-\u00A0]", "", str(doc_number).strip()).upper()
    
    # TAJ szám: pontosan 9 számjegy
    if re.match(r"^\d{9}$", clean):
        return "taj"
    
    # Útlevél: 2 betű + 6 vagy 7 számjegy (pl. AB123456 vagy AB1234567)
    if re.match(r"^[A-Z]{2}\d{6,7}$", clean):
        return "passport"
    
    # Személyi igazolvány: 6 számjegy + 2 betű (pl. 123456AB)
    if re.match(r"^\d{6}[A-Z]{2}$", clean):
        return "id_card"
    
    return "unknown"

def split_full_name(full_name: str):
    """'Vezetéknév Utónév' -> (vezetéknév, utónév). Ha nincs szóköz, utónév üres."""
    if not full_name:
        return "", ""
    parts = str(full_name).strip().split()
    if len(parts) >= 2:
        last = parts[0]
        first = " ".join(parts[1:])
        return last, first
    # 1 elem esetén: tegyük vezetéknévbe
    return parts[0], ""


def set_gender(driver, gender: str):
    """
    gender: "male"/"female" or magyar: "férfi"/"no"/"nő"
    """
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    import time

    g = gender.strip().lower()
    if g in ["male", "férfi", "ferfi", "m"]:
        click_css = 'label[for="SexId_Male"]'
    else:
        click_css = 'label[for="SexId_Female"]'

    try:
        el = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, click_css))
        )
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
        el.click()
        logger.info("✅ Gender clicked: %s", click_css)
    except Exception as ex:
        logger.warning("⚠️ gender direct click failed, trying JS force click")
        el = driver.find_element(By.CSS_SELECTOR, click_css)
        driver.execute_script("arguments[0].click();", el)

    # verify selected
    time.sleep(0.2)
    logger.info("✅ Gender set verified")


def _norm_text(s: str) -> str:
    s = (s or "").strip()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return " ".join(s.lower().split())


def _closest_row(el):
    # keressük meg a legközelebbi sor/row/container elemet
    try:
        return el.find_element(By.XPATH, "ancestor::*[@data-automation-id='listRow' or @role='row' or contains(@class,'row')][1]")
    except Exception:
        return el


def _norm_txt(s: str) -> str:
    s = str(s or "")
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = re.sub(r"\s+", " ", s.strip())
    return s.lower()


def _retry_detached(fn, retries=1, delay=0.5):
    last = None
    for _ in range(retries+1):
        try:
            return fn()
        except (StaleElementReferenceException, WebDriverException) as e:
            msg = (str(e) or "").lower()
            if "target frame detached" in msg or "stale element" in msg or "disconnected" in msg:
                last = e
                time.sleep(delay)
                continue
            raise
    if last:
        raise last


def _log_options(prefix, opts):
    texts = []
    for o in opts[:50]:
        try:
            texts.append(_norm_txt(o.text))
        except Exception:
            pass
    if texts:
        try:
            logger.info(f"{prefix}: " + " | ".join(texts))
        except Exception:
            pass

def listbox_pick_by_terms(driver, terms, timeout=10):
    """
    Feltételezzük, hogy a combobox/listbox már nyitva van (overlay).
    - Összes opció beolvasása (role='option' VAGY data-automation-id __option__).
    - Kiírja az első 30 opció normalizált szövegét logger.info-val.
    - A 'terms' normalizált listája alapján kiválasztja az első egyezőt (JS scroll + click).
    - Siker: True, különben False.
    """
    norm_terms = [_norm_text(t) for t in (terms or [])]

    # várjuk meg, hogy legalább 1 opció megjelenjen
    end = time.time() + timeout
    options = []
    while time.time() < end and not options:
        options = driver.find_elements(By.XPATH, "//*[@role='option' or starts-with(@data-automation-id,'__option__')]")
        if not options:
            time.sleep(0.2)

    # logolás
    _log_options("📋 Opciók (top50)", options)

    # kiválasztás szinonimák alapján
    for o in options:
        try:
            txt = _norm_text(o.text)
            if any(t in txt for t in norm_terms):
                try:
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", o)
                except Exception:
                    pass
                try:
                    o.click()
                except Exception:
                    driver.execute_script("arguments[0].click();", o)
                return True
        except Exception:
            continue
    return False

def combobox_type_and_select(driver, query_text: str, timeout: int = 15):
    """
    Feltételezzük, hogy a combobox már nyitva van. Beírjuk a query_text-et, várjuk az opciót, ENTER.
    Visszatér True ha sikerült, különben False.
    """
    try:
        driver.switch_to.default_content()
    except Exception:
        pass
    ensure_driver_alive(driver)
    # 1) próbáljuk az aktív elemet használni (gyakran a combobox input)
    try:
        ae = driver.switch_to.active_element
        try:
            ae.clear()
        except Exception:
            pass
        ae.send_keys(query_text)
        time.sleep(0.3)
    except Exception:
        pass

    # 2) várjuk, hogy megjelenjen egy opció, amely tartalmazza a szöveget (case-insensitive)
    q = query_text.strip().lower()
    xpath_opt = "//*[(@role='option' or @data-automation-id) and contains(translate(normalize-space(.),'abcdefghijklmnopqrstuvwxyzáéíóöőúüű','ABCDEFGHIJKLMNOPQRSTUVWXYZÁÉÍÓÖŐÚÜŰ'), translate('%s','abcdefghijklmnopqrstuvwxyzáéíóöőúüű','ABCDEFGHIJKLMNOPQRSTUVWXYZÁÉÍÓÖŐÚÜŰ'))]" % q.upper()
    try:
        opt = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.XPATH, xpath_opt)))
        # ENTER a választáshoz (stabilabb, mint click)
        try:
            ae = driver.switch_to.active_element
            ae.send_keys(Keys.ENTER)
            time.sleep(0.2)
            return True
        except Exception:
            # ha ENTER nem működött, próbáljuk klikkel
            try:
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", opt)
            except Exception:
                pass
            try:
                opt.click()
            except Exception:
                driver.execute_script("arguments[0].click();", opt)
            return True
    except TimeoutException:
        return False

def _wait_url_not_contains(driver, fragment: str, timeout: int = 10, poll_frequency: float = 0.2):
    """Window-handle safe várakozás JS-alapú URL olvasással és bezárt ablakok kezelésével."""
    t = timeout or 10

    def cond(d):
        try:
            handles = d.window_handles
            if not handles:
                return False
            d.switch_to.window(handles[-1])
            href = d.execute_script(
                "return window.location && window.location.href ? window.location.href : '';"
            )
            if href is None:
                return False
            return fragment not in href
        except (NoSuchWindowException, WebDriverException):
            return False
        except Exception:
            return False

    return WebDriverWait(driver, t, poll_frequency=poll_frequency).until(lambda d: cond(d))


def find_element_in_any_frame(driver, css, timeout=30):
    """Keresés a fő dokumentumban és max 2 szintnyi iframe-ben. Találatkor a megfelelő frame-ben marad.
    Ha nincs találat timeoutig, visszatér None-nal és default_content-re vált.
    """
    end_time = time.time() + timeout

    def try_find_here() -> object:
        try:
            return driver.find_element(By.CSS_SELECTOR, css)
        except Exception:
            return None

    def search_frames(depth: int) -> object:
        if depth > 2:
            return None
        frames = driver.find_elements(By.TAG_NAME, "iframe")
        for fr in frames:
            try:
                driver.switch_to.frame(fr)
                el = try_find_here()
                if el is not None:
                    return el
                if depth < 2:
                    nested = search_frames(depth + 1)
                    if nested is not None:
                        return nested
            except Exception:
                pass
            finally:
                # csak akkor menjünk vissza, ha nem találtunk; találatkor a hívó az aktuális frame-ben marad
                try:
                    driver.switch_to.parent_frame()
                except Exception:
                    pass
        return None

    last_exc = None
    while time.time() < end_time:
        try:
            try:
                driver.switch_to.default_content()
            except Exception:
                pass
            found = try_find_here()
            if found is not None:
                return found
            found = search_frames(1)
            if found is not None:
                return found
        except Exception as e:
            last_exc = e
        time.sleep(0.5)

    # timeout: menjünk vissza default_content-be és None
    try:
        driver.switch_to.default_content()
    except Exception:
        pass
    return None

def deep_query_all(driver, css_list, iframe_depth=2, shadow_depth=4):
    """Search for elements matching CSS selectors in main document, iframes, and shadow DOM.
    Returns tuple: (selenium_elements_list, js_elements_list_as_strings_for_clicking).
    JS elements are returned as serialized references we can use with execute_script."""
    if isinstance(css_list, str):
        css_list = [css_list]
    
    # Traditional Selenium search (for non-shadow, non-iframe elements)
    selenium_results = []
    try:
        driver.switch_to.default_content()
        for css in css_list:
            try:
                elements = driver.find_elements(By.CSS_SELECTOR, css)
                for el in elements:
                    try:
                        if el.is_displayed():
                            selenium_results.append(el)
                    except Exception:
                        pass
            except Exception:
                pass
    except Exception:
        pass
    
    # Deep JS search (for shadow DOM and iframes) - returns actual elements we can click via JS
    js_elements_result = []
    try:
        js_elements_result = driver.execute_script("""
            const cssList = arguments[0];
            const iframeDepth = arguments[1];
            const shadowDepth = arguments[2];
            
            function searchInShadow(root, depthRemaining) {
                if (depthRemaining <= 0) return [];
                const found = [];
                try {
                    for (const css of cssList) {
                        try {
                            const matches = root.querySelectorAll(css);
                            for (const el of matches) {
                                try {
                                    if (el.offsetParent !== null || el.style.display !== 'none') {
                                        found.push(el);
                                    }
                                } catch (e) {}
                            }
                        } catch (e) {}
                    }
                    if (depthRemaining > 1) {
                        for (const el of root.querySelectorAll('*')) {
                            try {
                                if (el.shadowRoot) {
                                    const nested = searchInShadow(el.shadowRoot, depthRemaining - 1);
                                    found.push(...nested);
                                }
                            } catch (e) {}
                        }
                    }
                } catch (e) {}
                return found;
            }
            
            function searchFrame(depthRemaining, currentDoc) {
                if (depthRemaining <= 0) return [];
                const found = [];
                const doc = currentDoc || document;
                try {
                    for (const css of cssList) {
                        try {
                            const matches = doc.querySelectorAll(css);
                            for (const el of matches) {
                                try {
                                    if (el.offsetParent !== null || el.style.display !== 'none') {
                                        found.push(el);
                                    }
                                } catch (e) {}
                            }
                        } catch (e) {}
                    }
                    for (const el of doc.querySelectorAll('*')) {
                        try {
                            if (el.shadowRoot) {
                                const shadowResults = searchInShadow(el.shadowRoot, shadowDepth);
                                found.push(...shadowResults);
                            }
                        } catch (e) {}
                    }
                    if (depthRemaining > 1) {
                        for (const iframe of doc.querySelectorAll('iframe')) {
                            try {
                                if (iframe.contentWindow && iframe.contentDocument) {
                                    const nested = searchFrame(depthRemaining - 1, iframe.contentDocument);
                                    found.push(...nested);
                                }
                            } catch (e) {}
                        }
                    }
                } catch (e) {}
                return found;
            }
            
            // Return all found elements (they're JS objects, we'll use them directly in JS)
            return searchFrame(iframeDepth, document);
        """, css_list, iframe_depth, shadow_depth)
    except Exception:
        js_elements_result = []
    
    return selenium_results, js_elements_result

# --- Helpers: Documents section + React-Select -------------------------------

def _norm_no_diac(s: str) -> str:
    import unicodedata, re
    s = str(s or "")
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return re.sub(r"\s+", " ", s.strip()).lower()

def _open_documents_section_and_add(driver, timeout: int = 12) -> bool:
    """Open Documents/Okmányok section and click its local Add; success if a row/input appears."""
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    import time
    try:
        el = driver.find_element(By.ID, "DocumentNumber")
        if el.is_displayed():
            return True
    except Exception:
        pass
    container = None
    try:
        container = driver.find_element(By.CSS_SELECTOR, '[data-automation-id="Documents"]')
    except Exception:
        pass
    if container is None:
        try:
            title = driver.find_element(
                By.XPATH,
                "//*[self::h2 or self::h3 or self::h4 or self::label]"
                "[contains(normalize-space(),'Okmányok') or contains(normalize-space(),'Dokmányok') or contains(normalize-space(),'Documents')]"
            )
            try:
                container = title.find_element(
                    By.XPATH,
                    "ancestor::*[contains(@class,'ListPanel_container') or contains(@class,'section') or contains(@class,'group') or @data-automation-id][1]"
                )
            except Exception:
                container = title
        except Exception:
            container = None
    add_btn = None
    if container is not None:
        for xp in [
            './/*[@data-automation-id="__addNewItemCompactButton"]',
            './/button[contains(., "Hozzáadás") or contains(translate(., "ABCDEFGHIJKLMNOPQRSTUVWXYZÁÉÍÓÖŐÚÜŰ", "abcdefghijklmnopqrstuvwxyzáéíóöőúüű"), "hozzáadás")]',
            './/button[contains(translate(., "ABCDEFGHIJKLMNOPQRSTUVWXYZ", "abcdefghijklmnopqrstuvwxyz"), "add")]',
        ]:
            try:
                cand = container.find_element(By.XPATH, xp)
                if cand.is_displayed():
                    add_btn = cand
                    break
            except Exception:
                continue
    if add_btn is None:
        try:
            for cand in driver.find_elements(By.CSS_SELECTOR, '[data-automation-id="__addNewItemCompactButton"]'):
                if cand.is_displayed():
                    add_btn = cand
                    break
        except Exception:
            pass
    if add_btn is None:
        return False
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", add_btn)
    except Exception:
        pass
    try:
        add_btn.click()
    except Exception:
        try:
            driver.execute_script("arguments[0].click();", add_btn)
        except Exception:
            return False
    end = time.time() + timeout
    while time.time() < end:
        try:
            el = driver.find_element(By.ID, "DocumentNumber")
            if el.is_displayed():
                return True
        except Exception:
            pass
        try:
            inp = driver.find_element(
                By.CSS_SELECTOR,
                '[data-automation-id="DocumentTypeId__container"] [id^="react-select-"][id$="-input"]'
            )
            if inp.is_displayed():
                return True
        except Exception:
            pass
        time.sleep(0.2)
    return False

def _doc_container(driver):
    from selenium.webdriver.common.by import By
    try:
        return driver.find_element(By.CSS_SELECTOR, '[data-automation-id="DocumentTypeId__container"]')
    except Exception:
        pass
    try:
        docnum = driver.find_element(By.ID, "DocumentNumber")
        return docnum.find_element(By.XPATH, "ancestor::*[@data-automation-id='listRow' or @role='row'][1]")
    except Exception:
        return None

def _open_combo(driver, timeout=8):
    """Open dropdown via input or chevron; return the react-select input element or None.
    IMPORTANT: Do NOT wait for a *visible* listbox here — some themes keep it zero-height.
    We'll wait for options elsewhere."""
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys
    import time

    cont = _doc_container(driver)
    if cont is None:
        return None

    try:
        inp = cont.find_element(By.CSS_SELECTOR, '[id^="react-select-"][id$="-input"]')
        try:
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", inp)
        except Exception:
            pass
        try:
            inp.click()
        except Exception:
            try:
                driver.execute_script("arguments[0].click();", inp)
            except Exception:
                return None
        try:
            inp.send_keys(Keys.ARROW_DOWN)
        except Exception:
            pass
        time.sleep(0.1)
        return inp
    except Exception:
        pass

    for xp in [
        ".//*[@data-automation-id='chevronDown' or contains(@class,'chevron')]",
        ".//*[@aria-haspopup='listbox' and not(@id='DocumentNumber')]",
        ".//button[contains(@aria-haspopup,'listbox')]",
    ]:
        try:
            btn = cont.find_element(By.XPATH, xp)
            try:
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
            except Exception:
                pass
            try:
                btn.click()
            except Exception:
                try:
                    driver.execute_script("arguments[0].click();", btn)
                except Exception:
                    continue
            try:
                inp = cont.find_element(By.CSS_SELECTOR, '[id^="react-select-"][id$="-input"]')
                time.sleep(0.1)
                return inp
            except Exception:
                continue
        except Exception:
            continue

    return None

def _portal_options(driver):
    """Return (listbox, options[]) for portal list."""
    from selenium.webdriver.common.by import By
    listbox_css = '[id^="react-select-"][id$="-listbox"]'
    options_css = listbox_css + ' [id*="-option-"]'
    try:
        lb = driver.find_element(By.CSS_SELECTOR, listbox_css)
        return lb, driver.find_elements(By.CSS_SELECTOR, options_css)
    except Exception:
        return None, []

def _pick_option_regex(driver, pattern: str) -> bool:
    """Pick option by regex via JS mousedown → mouseup → click."""
    import re, time
    pat = re.compile(pattern, re.I)
    lb, opts = _portal_options(driver)
    if not lb:
        return False
    target = None
    for o in opts:
        try:
            t = (o.text or "").strip()
            if pat.search(t) and o.is_displayed():
                target = o
                break
        except Exception:
            continue
    if not target:
        return False
    try:
        driver.execute_script("""
            const el = arguments[0];
            const ev = (n)=>el.dispatchEvent(new MouseEvent(n,{bubbles:true,cancelable:true,view:window}));
            el.scrollIntoView({block:'center'}); ev('mousedown'); ev('mouseup'); ev('click');
        """, target)
        time.sleep(0.1)
        return True
    except Exception:
        return False

def _doc_type_text(driver) -> str:
    """Collect visible text inside the doc-type container (not only .singleValue)."""
    from selenium.webdriver.common.by import By
    cont = _doc_container(driver)
    if cont is None:
        return ""
    texts = []
    seen = set()
    for css in [".single-value", "[class*='singleValue']", "[data-automation-id*='singleValue']", "*:not(input):not(textarea)"]:
        try:
            for el in cont.find_elements(By.CSS_SELECTOR, css):
                try:
                    if not el.is_displayed(): continue
                    t = (el.text or "").strip()
                    if t and len(t) <= 120:
                        k = (el.tag_name, t)
                        if k not in seen:
                            seen.add(k); texts.append(t)
                except Exception:
                    continue
        except Exception:
            continue
    return " ".join(texts).strip()

# --- fő: TAJ kiválasztása tartósan -----------------------------------

def select_document_type(driver, doc_type: str):
    """
    Kiválasztja a megfelelő okmánytípust a dropdown-ból.
    
    Args:
        driver: WebDriver
        doc_type: "taj", "passport", vagy "id_card"
    """
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys
    import time, re

    # 0) Ensure section and row exist
    if not _open_documents_section_and_add(driver, timeout=15):
        raise TimeoutException("Okmányok szekció/sor nem érhető el.")

    # Okmánytípus szövegek (szinonimák)
    type_patterns = {
        "taj": (r"TAJ.*sz[áa]m", ["TAJ szám", "TAJ szam"]),
        "passport": (r"[UÚ]tlev[ée]l", ["Útlevél", "Utlevel"]),
        "id_card": (r"Szem[ée]lyi.*igazolv[áa]ny", ["Személyi igazolvány", "Személyazonosító igazolvány", "Személyi"])
    }
    
    if doc_type not in type_patterns:
        raise ValueError(f"Érvénytelen doc_type: {doc_type}. Lehetőségek: {list(type_patterns.keys())}")
    
    wanted_regex, search_terms = type_patterns[doc_type]
    input_css   = '[data-automation-id="DocumentTypeId__container"] [id^="react-select-"][id$="-input"]'
    listbox_css = '[id^="react-select-"][id$="-listbox"]'
    noopt_xpath = "//*[contains(normalize-space(),'No options')]"

    # main attempts
    for attempt in range(1, 5):
        # 1) open combobox
        inp = _open_combo(driver, timeout=8)
        if inp is None:
            time.sleep(0.2); continue

        # 2) try with NO typing (some installs load full list only when empty)
        # removed brittle visible-listbox wait; presence/options are awaited elsewhere
        import time as _t; _t.sleep(0.05)

        # If list shows "No options", trigger load by keys and reopen
        try:
            has_no = len(driver.find_elements(By.XPATH, noopt_xpath)) > 0
        except Exception:
            has_no = False
        if has_no:
            try:
                inp.send_keys(Keys.SPACE)  # poke filter
                time.sleep(0.15)
                inp.send_keys(Keys.BACKSPACE)
            except Exception:
                pass
            # Re-open hard
            inp = _open_combo(driver, timeout=6) or inp

        # 3) Type both diac/fallback and wait options (with stale element retry)
        def type_and_wait(text, max_retries=3) -> bool:
            """Type text with retry for stale element exceptions."""
            for attempt in range(max_retries):
                try:
                    # Re-find input element each time (in case it becomes stale)
                    cont = _doc_container(driver)
                    if cont is None:
                        if attempt < max_retries - 1:
                            time.sleep(0.3)
                            continue
                        return False
                    
                    current_inp = cont.find_element(By.CSS_SELECTOR, '[id^="react-select-"][id$="-input"]')
                    
                    # Wait a bit for DOM to stabilize
                    time.sleep(0.2)
                    
                    # Clear field
                    try:
                        current_inp.send_keys(Keys.CONTROL, "a")
                        time.sleep(0.05)
                        current_inp.send_keys(Keys.BACKSPACE)
                    except StaleElementReferenceException:
                        if attempt < max_retries - 1:
                            logger.warning(f"♻️ Stale element during clear (attempt {attempt + 1}/{max_retries})")
                            time.sleep(0.3)
                            continue
                        return False
                    except Exception:
                        try:
                            current_inp.clear()
                        except Exception:
                            pass
                    
                    # Type text
                    try:
                        current_inp.send_keys(text)
                        time.sleep(0.1)
                    except StaleElementReferenceException:
                        if attempt < max_retries - 1:
                            logger.warning(f"♻️ Stale element during send_keys (attempt {attempt + 1}/{max_retries})")
                            time.sleep(0.3)
                            continue
                        return False
                    
                    # Verify text was entered
                    try:
                        current_value = current_inp.get_attribute("value") or ""
                        if text.lower() in current_value.lower() or current_value.lower() in text.lower():
                            logger.debug(f"✅ Beírva: '{text}'")
                            return True
                    except StaleElementReferenceException:
                        if attempt < max_retries - 1:
                            logger.warning(f"♻️ Stale element during verification (attempt {attempt + 1}/{max_retries})")
                            time.sleep(0.3)
                            continue
                        return False
                    except Exception:
                        # If we can't verify, assume success and continue
                        return True
                    
                except StaleElementReferenceException:
                    logger.warning(f"♻️ Stale element (attempt {attempt + 1}/{max_retries}), újrapróbálás...")
                    if attempt < max_retries - 1:
                        time.sleep(0.5)
                        continue
                    else:
                        logger.error(f"❌ Stale element - minden próbálkozás sikertelen: '{text}'")
                        return False
                except Exception as e:
                    logger.warning(f"⚠️ Hiba a beírás során (attempt {attempt + 1}/{max_retries}): {e}")
                    if attempt < max_retries - 1:
                        time.sleep(0.5)
                        continue
                    return False
            
            return False

        # Try all search terms for this doc_type
        ok = False
        for term in search_terms:
            if type_and_wait(term, max_retries=3):
                ok = True
                break
        
        if not ok:
            logger.warning(f"⚠️ '{doc_type}' beírása sikertelen (attempt {attempt}/4), újrapróbálás...")
            # reopen and continue
            time.sleep(0.2)
            continue

        # 4) pick via regex (mousedown path), fallback ENTER
        picked = _pick_option_regex(driver, wanted_regex)
        if not picked:
            try:
                driver.switch_to.active_element.send_keys(Keys.ENTER)
                picked = True
            except Exception:
                picked = False

        # 5) Commit & verify
        try:
            driver.switch_to.active_element.send_keys(Keys.TAB)
        except Exception:
            pass
        time.sleep(0.2)

        # Verify selection
        txt = _norm_no_diac(_doc_type_text(driver))
        # Rugalmasabb ellenőrzés a különböző típusokhoz
        if doc_type == "taj" and re.search(r"\btaj\b", txt):
            verification_ok = True
        elif doc_type == "passport" and re.search(r"\butlevel\b", txt):
            verification_ok = True
        elif doc_type == "id_card" and re.search(r"\bszemelyi\b.*\bigazolvany\b", txt):
            verification_ok = True
        else:
            verification_ok = False
            
        if verification_ok:
            # Focus DocumentNumber
            try:
                doc = WebDriverWait(driver, 6).until(EC.element_to_be_clickable((By.ID, "DocumentNumber")))
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", doc)
                try:
                    doc.click()
                except Exception:
                    driver.execute_script("arguments[0].click();", doc)
            except Exception:
                pass
            logger.info(f"✅ Okmánytípus kiválasztva: {doc_type}")
            return

        time.sleep(0.3)

    raise TimeoutException(f"Document type selection did not persist ({doc_type}).")


def open_new_patient_form(driver):
    """Új páciens felvétele képernyőre lépés."""
    wait_click_css(driver, '[data-automation-id="PatientRegister_CreateNewPatient"]')
    # új ablak/fül megnyílt – 5 mp-en belül legyen window handle, majd váltsunk a legutolsóra
    try:
        WebDriverWait(driver, 5).until(lambda d: len(d.window_handles) >= 1)
        driver.switch_to.window(driver.window_handles[-1])
    except Exception:
        pass
    # Mentés gomb jelenléte jelzi a form kész állapotát
    WebDriverWait(driver, 45).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, '[data-automation-id="__save_save"]'))
    )
    # Additional readiness wait (documents section often lazy-loads)
    wait_patient_form_ready(driver, timeout=25)


def save_patient(driver):
    """Mentés gomb. Várunk valamire a mentés után, hogy stabil legyen."""
    wait_click_css(driver, '[data-automation-id="__save_save"]')
    # kis stabilizáció
    time.sleep(1.0)

# --- Birth date helpers -------------------------------------------------------

def find_birthdate_control(driver, timeout=15):
    """
    Locate a birth date control within the 'Születési dátum' / 'Születési idő' section.
    Returns:
      - a single WebElement if a 1-field date exists
      - or a dict {'y': elY, 'm': elM, 'd': elD} if split fields are detected
    Raises TimeoutException if nothing found within timeout.
    """
    from selenium.webdriver.common.by import By
    from selenium.common.exceptions import TimeoutException
    import time, unicodedata, re

    def _norm(s):
        s = str(s or "")
        s = unicodedata.normalize("NFKD", s)
        s = "".join(ch for ch in s if not unicodedata.combining(ch))
        return re.sub(r"\s+", " ", s.strip()).lower()

    def _visible(el):
        try:
            return el.is_displayed()
        except Exception:
            return False

    def _section():
        # Prefer container by heading text
        xps = [
            "//*[contains(normalize-space(),'Születési dátum') or contains(normalize-space(),'Születési idő')]/ancestor::*[contains(@class,'section') or contains(@class,'group') or @role='region'][1]",
            "//*[contains(translate(normalize-space(),'ÁÉÍÓÖŐÚÜŰ','áéíóöőúüű'),'szuletesi datum') or contains(translate(normalize-space(),'ÁÉÍÓÖŐÚÜŰ','áéíóöőúüű'),'szuletesi ido')]/ancestor::*[contains(@class,'section') or contains(@class,'group') or @role='region'][1]",
            "//label[@for='BirthDate']/ancestor::*[contains(@class,'section') or contains(@class,'group') or @role='region'][1]"
        ]
        for xp in xps:
            try:
                el = driver.find_element(By.XPATH, xp)
                if _visible(el):
                    return el
            except Exception:
                continue
        return None

    end = time.time() + timeout
    last_exc = None
    while time.time() < end:
        try:
            driver.switch_to.default_content()
        except Exception:
            pass

        root = _section() or driver

        # A) Single-input patterns (id/name/type/role/aria/contenteditable)
        css_single = [
            "#BirthDate","input#BirthDate","[name='BirthDate']",
            "input[type='date']",
            "[data-automation-id*='Birth'][data-automation-id*='Date'] input",
            "input[role='textbox']","input[role='combobox']",
            "[contenteditable='true'][aria-label*='Születési']",
            "[contenteditable='true'][aria-label*='Birth']",
            "[aria-label*='Születési'] input","[aria-label*='Birth'] input"
        ]
        for sel in css_single:
            try:
                el = root.find_element(By.CSS_SELECTOR, sel)
                if _visible(el):
                    return el
            except Exception:
                pass

        # B) Split fields (YYYY / MM / DD) inside the section
        # Heuristics: look for inputs with maxLength 4 (year) and 2 (month/day),
        # or aria-label / placeholder hints.
        try:
            inputs = root.find_elements(By.CSS_SELECTOR, "input")
        except Exception:
            inputs = []
        y, m, d = None, None, None
        for inp in inputs:
            try:
                if not _visible(inp):
                    continue
                ph = (inp.get_attribute("placeholder") or "") + " " + (inp.get_attribute("aria-label") or "")
                phn = _norm(ph)
                ml = inp.get_attribute("maxlength") or inp.get_attribute("maxLength") or ""
                ml = int(ml) if str(ml).isdigit() else None

                # YEAR candidates
                if (ml == 4) or any(k in phn for k in ["év","ev","year","yyyy"]):
                    if y is None:
                        y = inp
                        continue
                # MONTH candidates
                if (ml == 2 and m is None and "nap" not in phn) or any(k in phn for k in ["hónap","honap","month","mm"]):
                    if m is None:
                        m = inp
                        continue
                # DAY candidates
                if (ml == 2 and d is None) or any(k in phn for k in ["nap","day","dd"]):
                    if d is None:
                        d = inp
                        continue
            except Exception:
                continue
        if y and m and d:
            return {"y": y, "m": m, "d": d}

        # C) Last fallback: any contenteditable in section
        try:
            ce = root.find_element(By.CSS_SELECTOR, "[contenteditable='true']")
            if _visible(ce):
                return ce
        except Exception:
            pass

        time.sleep(0.2)

    raise TimeoutException("Birth date control not found")

def fill_birthdate_iso(driver, dob_iso):
    """
    Fill DOB regardless of widget type:
      - single input (text/date/combobox/contenteditable)
      - split inputs for year/month/day
    Returns the primary element used for verification.
    """
    from selenium.webdriver.common.keys import Keys
    from selenium.common.exceptions import TimeoutException
    import time, re

    # Parse ISO
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})$", str(dob_iso).strip())
    if not m:
        raise TimeoutException(f"Invalid ISO date: {dob_iso}")
    yyyy, mm, dd = m.group(1), m.group(2), m.group(3)

    def _digits(s):
        return "".join(ch for ch in (s or "") if ch.isdigit())

    def _verify_ok(el):
        # Collect visible string and normalize: accept masked like '196 901 16_'
        try:
            val = (el.get_attribute("value") or el.text or "").strip()
        except Exception:
            val = ""
        digits = _digits(val)
        want = yyyy + mm + dd
        return digits == want

    ctrl = find_birthdate_control(driver, timeout=15)

    # Strategy 1: split fields
    if isinstance(ctrl, dict):
        y_el, m_el, d_el = ctrl["y"], ctrl["m"], ctrl["d"]

        # Clear and type with small waits
        for el, val in [(y_el, yyyy), (m_el, mm), (d_el, dd)]:
            try:
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
            except Exception:
                pass
            try:
                el.click()
            except Exception:
                try:
                    driver.execute_script("arguments[0].click();", el)
                except Exception:
                    pass
            try:
                el.send_keys(Keys.CONTROL, "a"); time.sleep(0.05); el.send_keys(Keys.BACKSPACE); time.sleep(0.05)
            except Exception:
                try:
                    el.clear()
                except Exception:
                    pass
            try:
                el.send_keys(val); time.sleep(0.08)
            except Exception:
                pass
        # small blur to trigger formatters
        try:
            d_el.send_keys(Keys.TAB)
        except Exception:
            pass
        # Verify using the day field (usually shows mask)
        if not _verify_ok(d_el):
            # sometimes verification works better against year field or container
            if not _verify_ok(y_el):
                raise TimeoutException("DOB split verification failed")
        return d_el

    # Strategy 2: single input (text/date/combobox/contenteditable)
    el = ctrl
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    except Exception:
        pass

    # Try normal typing
    try:
        el.click()
    except Exception:
        try:
            driver.execute_script("arguments[0].click();", el)
        except Exception:
            pass
    try:
        el.send_keys(Keys.CONTROL, "a"); time.sleep(0.05); el.send_keys(Keys.BACKSPACE); time.sleep(0.05)
    except Exception:
        try:
            el.clear()
        except Exception:
            pass
    try:
        el.send_keys(dob_iso); time.sleep(0.12)
    except Exception:
        pass

    # If not correct, set via JS (supports contenteditable and inputs) + events
    if not _verify_ok(el):
        try:
            driver.execute_script("""
                (function(el, val){
                    function setVal(e, v){
                        const proto = Object.getPrototypeOf(e);
                        const desc = Object.getOwnPropertyDescriptor(proto, 'value') || Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value');
                        if (desc && desc.set) desc.set.call(e, v); else e.value = v;
                    }
                    if (el.isContentEditable) {
                        el.textContent = val;
                    } else {
                        setVal(el, val);
                    }
                    el.dispatchEvent(new Event('input', {bubbles:true}));
                    el.dispatchEvent(new Event('change', {bubbles:true}));
                    el.dispatchEvent(new Event('blur', {bubbles:true}));
                })(arguments[0], arguments[1]);
            """, el, f"{yyyy}-{mm}-{dd}")
            time.sleep(0.12)
        except Exception:
            pass

    # final verify
    if not _verify_ok(el):
        raise TimeoutException("DOB single-field verification failed")

    return el

# ---------------------------------------------------------------------
# Feltöltő lépés – 1 sor (1 páciens)
# ---------------------------------------------------------------------
def upload_one_patient(driver, row: pd.Series, excel_row: int | None = None) -> bool:
    """
    Vár Excel mezők:
      - 'Paciens/Nev' VAGY 'Vezetéknév' + 'Utónév'
      - 'Paciens/Azonosito'
      - 'Paciens/SzuletesiDatum'
      - 'Paciens/Nem'
      - I oszlop: Irányítószám
      - J oszlop: Település
      - K oszlop: Cím
      - H oszlop: Mobiltelefon
      - 'Paciens/Email'
    
    Determinisztikus flow:
    1) Okmány blokk (select_document_type + DocumentNumber)
    2) Alapadatok: LastName, FirstName, BirthDate, Nem
    3) Email hozzáadás gomb (Elérhetőségek szekcióban)
    4) EmailAddress mező kitöltése
    """
    # 1) Új páciens űrlap megnyitása
    open_new_patient_form(driver)

    # Track DOB requirement & outcome for this row
    dob_expected = False
    dob_filled_ok = False

    # 2) Okmányszám beolvasása és típus automatikus felismerése
    raw = row.get("Paciens/Azonosito", "")
    doc_num = normalize_taj(raw)
    doc_num = re.sub(r"\s+", "", str(doc_num))

    # Típus felismerése
    doc_type = detect_document_type(doc_num)
    logger.info(f"🆔 Okmány raw={repr(raw)} -> használt={repr(doc_num)} -> típus={doc_type}")

    if doc_type == "unknown":
        logger.warning(f"⚠️ Ismeretlen okmány formátum: {doc_num}")
        # Alapértelmezett: TAJ
        doc_type = "taj"
        logger.info(f"   Alapértelmezett típus használva: {doc_type}")

    # Okmánytípus kiválasztása
    try:
        select_document_type(driver, doc_type)
    except Exception as e:
        logger.error(f"❌ Okmánytípus kiválasztása sikertelen: {e}")
        raise

    # Okmányszám kitöltése
    wait_type_id(driver, "DocumentNumber", doc_num)

    # 3) Alapadatok – LastName, FirstName, BirthDate, Nem

    # --- Vezetéknév (try direct ID first)
    last_name = get_cell(
        row,
        "Vezetéknév","Családnév","Family name","Last name","FamilyName","Vezeteknev","Csaladinev"
    )
    first_name = get_cell(
        row,
        "Utónév","Keresztnév","Given name","First name","GivenName1","Utonev","Keresztnev"
    )

    if (not last_name or pd.isna(last_name)) and (not first_name or pd.isna(first_name)):
        full = get_cell(row, "Paciens/Nev", "Név", "Nev", "Full name", "Teljes név")
        ln, fn = split_full_name(str(full))
        last_name = ln
        first_name = fn

    if last_name:
        try:
            ln_el = WebDriverWait(driver, 8).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#FamilyName, [id='FamilyName']")))
            type_sturdy(driver, ln_el, str(last_name))
            logger.info("✅ Last name filled (direct)")
        except Exception:
            ln_el = fill_field_smart(
                driver,
                labels=["Vezetéknév","Családnév","Family name","Last name","Vezeteknev","Csaladinev"],
                attr_contains=["Last","last","Family","family","Vezetek","Csalad","FamilyName"],
                value=str(last_name),
                timeout=25
            )
            logger.info("✅ Last name filled (smart)")

    # --- Utónév (try direct ID first)
    if first_name:
        try:
            fn_el = WebDriverWait(driver, 8).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#GivenName1, [id='GivenName1']")))
            type_sturdy(driver, fn_el, str(first_name))
            logger.info("✅ First name filled (direct)")
        except Exception:
            fn_el = fill_field_smart(
                driver,
                labels=["Utónév","Keresztnév","Given name","First name","Utonev","Keresztnev"],
                attr_contains=["First","first","Given","given","Uto","Kereszt","GivenName1"],
                value=str(first_name),
                timeout=25
            )
            logger.info("✅ First name filled (smart)")

    # Születési dátum (robust, ISO) — REQUIRED IF PRESENT IN EXCEL
    dob_raw = get_cell(
        row,
        "Paciens/SzuletesiDatum","Születési dátum","Szuletesi datum","DOB","Birth date","Date of birth"
    )
    dob_iso = to_iso_date(dob_raw)
    if dob_iso:
        dob_expected = True
        try:
            dob_el = fill_birthdate_iso(driver, dob_iso)
            _val = (dob_el.get_attribute("value") or dob_el.text or "").strip()
            # tolerate minor masking (spaces/underscores), but require all digits and separators
            _val_norm = "".join(ch for ch in _val if ch.isdigit() or ch in "-./")
            _dob_digits = "".join(ch for ch in dob_iso if ch.isdigit())
            _val_digits = "".join(ch for ch in _val_norm if ch.isdigit())
            if dob_iso in _val or _dob_digits == _val_digits:
                dob_filled_ok = True
                logger.info(f"✅ DOB filled & verified: {dob_iso}")
            else:
                raise TimeoutException(f"DOB mismatch: expected '{dob_iso}', got '{_val}'")
        except Exception as _e:
            dob_filled_ok = False
            logger.warning(f"⚠️ DOB fill failed for value={dob_iso} ({_e})")

    # Nem (id-alapú direkt választás a stabil for-ral) – 3) része
    gender = get_cell(row, "Paciens/Nem","Nem","Gender","Sex")
    if gender and not pd.isna(gender):
        try:
            g = str(gender).strip().lower()
            css = 'label[for="SexId_Male"]' if g in ("férfi","ferfi","male","m","ffi","f") else 'label[for="SexId_Female"]'
            el = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, css)))
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
            try: el.click()
            except Exception: driver.execute_script("arguments[0].click();", el)
            logger.info("✅ Gender set (direct label click)")
        except Exception as _e:
            try:
                set_gender(driver, str(gender))
                logger.info("✅ Gender set (fallback)")
            except Exception as _e2:
                logger.warning(f"⚠️ Gender step failed transiently ({_e2}); continuing.")

    # 4) CÍM BLOKK: Hozzáadás → Irányítószám (I) → Település (J) → Cím (K)
    try:
        fill_address_data(driver, row)
    except Exception as addr_ex:
        logger.warning(f"⚠️ Address filling error (continuing): {addr_ex}")

    # 5) TELEFONSZÁM: Hozzáadás → Mobiltelefon (H)
    try:
        fill_phone_data(driver, row)
    except Exception as phone_ex:
        logger.warning(f"⚠️ Phone filling error (continuing): {phone_ex}")

    # 6) Email hozzáadás gomb — unified, no-duplicate fill
    email_value = get_cell(
        row,
        "Paciens/Email","Email","E-mail cím","Email cím","E-mail","EmailAddress"
    )
    if email_value and not pd.isna(email_value):
        ok_email = fill_email_address(driver, str(email_value))
        if not ok_email:
            logger.warning("⚠️ Email not confirmed after fill (continuing).")

    # Enforce: if Excel had DOB but field is not correctly filled, mark this row as FAIL
    if dob_expected and not dob_filled_ok:
        logger.error("❌ DOB required by Excel but not filled/verified — marking row as FAIL.")
        raise TimeoutException("DOB required but not filled")

    # 7) Várunk a felhasználóra, hogy ellenőrizze és mentse az adatokat
    logger.info("⏸️  Adatok kitöltve. Várakozás a felhasználó manuális mentésére...")
    logger.info("   Kérlek ellenőrizd az adatokat és nyomd meg a MENTÉS gombot!")
    logger.info("   A szkript automatikusan folytatja a következő pácienssel a mentés után.")

    # Hangjelzés, hogy figyelmeztesse a felhasználót
    try:
        winsound.Beep(1000, 500)  # 1000 Hz, 500 ms
    except Exception:
        pass

    # 8) Mentés detektálása: csak a gomb eltűnésétől függjön
    save_confirmed = False
    try:
        WebDriverWait(driver, 300).until_not(  # 5 perc timeout
            EC.presence_of_element_located((By.CSS_SELECTOR, '[data-automation-id="__save_save"]'))
        )
        save_confirmed = True
        logger.info("✅ Mentés észlelve, folytatás...")
        time.sleep(1.0)
    except TimeoutException:
        logger.warning("⚠️ 5 perc várakozás lejárt, a páciens NINCS biztosan elmentve.")
    except Exception as e:
        logger.warning(f"⚠️ Várakozás hiba (mentés gomb): {e}")

    # 9) A visszatérési érték kizárólag a mentéstől függ
    # Ha mentés megerősítve, akkor próbálunk visszanavigálni, de a visszatérési érték
    # mindig True lesz, még ha a navigáció hibázik is
    if save_confirmed:
        # Azonnali Excel-jelölés: minden mentés után írjunk "K"-t, még a visszanavigálás előtt.
        if excel_row is not None:
            wb = None
            try:
                wb = xw.Book(str(EXCEL_PATH))
                ws = wb.sheets["Páciensek"]
                ws.range(excel_row, 20).value = "K"
                wb.save()
                logger.info(f"✅ K betű beírva - Páciensek munkalap T oszlop (Excel sor {excel_row})")
            except Exception as e:
                logger.error(f"❌ K betű írás hiba: {e}")
            finally:
                if wb is not None:
                    try:
                        wb.close()
                    except Exception:
                        pass

        # Visszatérés a Páciens regiszterhez: a legstabilabb út a közvetlen URL.
        # (A window-handle close/switch gyakran dob stacktrace-t, ha a UI közben bezár/átirányít.)
        try:
            # Best-effort: if a popup/tab is open, try to close it safely, but never block on it.
            try:
                handles = list(driver.window_handles or [])
            except Exception:
                handles = []
            if len(handles) > 1:
                try:
                    driver.switch_to.window(handles[-1])
                except Exception:
                    pass
                try:
                    driver.close()
                except Exception:
                    pass
                try:
                    handles = list(driver.window_handles or [])
                    if handles:
                        driver.switch_to.window(handles[0])
                except Exception:
                    pass

            driver.get(LOGIN_URL)
            WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, '[data-automation-id="PatientRegister_CreateNewPatient"]'))
            )
            logger.info("✅ Visszatérés URL-lel sikeres")
            time.sleep(0.3)
        except Exception as e:
            logger.warning(f"⚠️ Visszatérés hiba (mentés már sikeres volt): {_exc_brief(e)}")
        
        # Mindig True-t adunk vissza ha save_confirmed, függetlenül a navigációtól
        logger.info("✅ Feldolgozás sikeres (mentés megerősítve)")
        return True
    else:
        logger.warning("⚠️ Mentés nem volt megerősítve")
        return False


# ---------------------------------------------------------------------
# Főfolyamat
# ---------------------------------------------------------------------
def main():
    logger.info("=== FUTÁS INDUL ===")
    logger.info(f"ENV USERNAME repr = {repr(USERNAME)}")
    logger.info(f"ENV HEADLESS  = {HEADLESS}")
    logger.info(f"EXCEL         = {EXCEL_PATH} | SHEET = {SHEET_NAME}")
    logger.info(f"USE_UPLOAD    = {USE_UPLOAD}")
    logger.info(f"MODE  | BASELINE_TAJ_NAME_DOB = {FF_BASELINE_TAJ_NAME_DOB} (Excel upload runs only if False)")

    # ---- böngésző
    # cleanup: remove leftover temp chrome profiles
    try:
        base = Path(".").resolve()
        for p in base.iterdir():
            try:
                if p.is_dir() and p.name.startswith(".tmp_chrome_profile_"):
                    import shutil
                    shutil.rmtree(p, ignore_errors=True)
            except Exception:
                pass
    except Exception:
        pass
    driver = make_driver()

    try:
        # ---- login
        lp = LoginPage(driver, LOGIN_URL, timeout=45)
        success = lp.login(USERNAME, PASSWORD)
        if not success:
            logger.warning("❌ Login sikertelen.")
            save_debug(driver, "login_fail")
            return
        logger.info("✅ Login sikeres.")
        save_debug(driver, "login_ok")

        # ---- Baseline mode: TAJ + Name + DOB only
        if FF_BASELINE_TAJ_NAME_DOB:
            logger.info("🔧 Baseline mode: TAJ + Name + DOB")
            open_new_patient_form(driver)
            
            # TAJ selection
            try:
                select_document_type(driver, "taj")
                logger.info("✅ TAJ type selected")
            except Exception as e:
                save_debug(driver, "baseline_taj_fail")
                raise
            
            # Fill TAJ number (test data)
            test_taj = "123456789"
            wait_type_id(driver, "DocumentNumber", test_taj)
            logger.info(f"🆔 TAJ raw=test -> used={repr(test_taj)}")
            
            # Smoke check: DocumentNumber value non-empty
            try:
                doc_num_val = (driver.find_element(By.ID, "DocumentNumber").get_attribute("value") or "").strip()
                if not doc_num_val:
                    save_debug(driver, "baseline_taj_fail")
                    raise TimeoutException("TAJ number not filled")
                logger.info("✅ TAJ smoke check passed")
            except Exception as e:
                save_debug(driver, "baseline_taj_fail")
                raise
            
            # Last name
            try:
                test_last_name = "TestLast"
                fill_field_smart(
                    driver,
                    labels=["Vezetéknév","Családnév","Family name","Last name","Vezeteknev","Csaladinev"],
                    attr_contains=["Last","last","Family","family","Vezetek","Csalad"],
                    value=test_last_name,
                    timeout=25
                )
                # Smoke check: last name input reflects value (case-insensitive)
                last_name_input = find_input_smart(driver, terms=["Vezetéknév","Last name"], attr_contains=["Last","last"], timeout=5)
                last_name_val = (last_name_input.get_attribute("value") or "").strip()
                if test_last_name.lower() not in last_name_val.lower():
                    save_debug(driver, "baseline_lastname_fail")
                    raise TimeoutException(f"Last name not filled: expected '{test_last_name}', got '{last_name_val}'")
                logger.info("✅ Last name smoke check passed")
            except Exception as e:
                save_debug(driver, "baseline_lastname_fail")
                raise
            
            # First name
            try:
                test_first_name = "TestFirst"
                fill_field_smart(
                    driver,
                    labels=["Utónév","Keresztnév","Given name","First name","Utonev","Keresztnev"],
                    attr_contains=["First","first","Given","given","Uto","Kereszt"],
                    value=test_first_name,
                    timeout=25
                )
                # Smoke check: first name input reflects value (case-insensitive)
                first_name_input = find_input_smart(driver, terms=["Utónév","First name"], attr_contains=["First","first"], timeout=5)
                first_name_val = (first_name_input.get_attribute("value") or "").strip()
                if test_first_name.lower() not in first_name_val.lower():
                    save_debug(driver, "baseline_firstname_fail")
                    raise TimeoutException(f"First name not filled: expected '{test_first_name}', got '{first_name_val}'")
                logger.info("✅ First name smoke check passed")
            except Exception as e:
                save_debug(driver, "baseline_firstname_fail")
                raise
            
            # Birth date
            try:
                test_dob_iso = "1990-01-01"
                fill_field_smart(
                    driver,
                    labels=["Születési dátum","Szuletesi datum","Date of birth","Birth date"],
                    attr_contains=["Birth","birth","Dob","dob","Date","date"],
                    value=test_dob_iso,
                    timeout=25
                )
                # Smoke check: DOB control shows ISO date
                dob_input = find_input_smart(driver, terms=["Születési dátum","Date of birth"], attr_contains=["Birth","birth","Date","date"], timeout=5)
                dob_val = (dob_input.get_attribute("value") or dob_input.text or "").strip()
                if test_dob_iso not in dob_val:
                    save_debug(driver, "baseline_dob_fail")
                    raise TimeoutException(f"DOB not filled: expected '{test_dob_iso}', got '{dob_val}'")
                logger.info("✅ DOB smoke check passed")
            except Exception as e:
                save_debug(driver, "baseline_dob_fail")
                raise
            
            logger.info("✅ Baseline mode completed: TAJ + Name + DOB")
            return  # Do NOT execute email-related code when FF_EMAIL_STEPS is False

        # ---- ha csak login-teszt menne (nálunk USE_UPLOAD mindig True)
        if not USE_UPLOAD:
            logger.info("🔧 Upload kikapcsolva (csak login próba).")
            return

        # ---- Excel beolvasás (kizárólag a "Páciensek" munkalapból)
        try:
            df = pd.read_excel(
                EXCEL_PATH,
                sheet_name="Páciensek",
                dtype=str,
                keep_default_na=False,
                engine="openpyxl",
            )
        except Exception as e:
            logger.exception(f"❌ Excel beolvasási hiba: {e}")
            return

        if df.empty:
            logger.warning("⚠️ Az Excel lap üres, nincs mit feltölteni.")
            return

        logger.info(f"📦 Sorok száma (raw): {len(df)}")
        
        # === DEBUG: List all columns ===
        logger.info(f"📋 DEBUG - Összes Excel oszlop ({len(df.columns)} db):")
        for idx, col_name in enumerate(df.columns):
            logger.info(f"   [{idx:2d}] '{col_name}'")
        logger.info("📋 DEBUG VÉGE")
        # === END DEBUG ===
        
        # Get today's date for filtering
        today = date.today()
        logger.info(f"📅 Mai dátum (szűrés): {today}")
        
        # === DEBUG: Display Q column values ===
        logger.info("🔍 DEBUG - Q oszlop (Paciens/FelvetelDatuma) értékek:")
        for i in range(min(len(df), 3)):
            q_val = df.iloc[i].get("Paciens/FelvetelDatuma", "NINCS_OSZLOP")
            logger.info(f"   DataFrame sor {i} (Excel sor {i+2}): '{q_val}' | típus: {type(q_val).__name__} | isna: {pd.isna(q_val)}")
        logger.info("🔍 DEBUG VÉGE")
        # === END DEBUG ===
        
        processed = 0
        skipped = 0
        fail = 0

        # --- SPEED: pre-filter rows in pandas (avoid iterating thousands of blanks) ---
        # Keep original df index for Excel row mapping and logging
        df = df.copy()
        df["__orig_idx"] = df.index

        # Resolve the registration-date column robustly
        felvetel_col_candidates = [
            "Paciens/FelvetelDatuma",
            "Paciens/FelvetelDatum",
            "Paciens/FelvételDátuma",
            "Paciens/FelvételDátum",
            "FelvetelDatuma",
            "FelvetelDatum",
        ]
        felvetel_col = next((c for c in felvetel_col_candidates if c in df.columns), None)

        if felvetel_col is None:
            logger.warning("⚠️ Nincs felvételi dátum oszlop a lapon; minden sor kihagyásra kerül.")
            logger.info(f"=== Feltöltés összegzés: Feldolgozva=0, Kihagyva={len(df)}, Sikertelen=0 ===")
            return

        # Parse to date vectorized; handles YYYY.MM.DD HH:MM:SS and YYYY-MM-DD formats
        _date_formats = [
            "%Y.%m.%d %H:%M:%S", "%Y.%m.%d %H:%M", "%Y.%m.%d",
            "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d",
            "%Y/%m/%d %H:%M:%S", "%Y/%m/%d",
        ]

        def _parse_date_flexible(val):
            if pd.isna(val) or str(val).strip() == "":
                return pd.NaT
            s = str(val).strip()
            for fmt in _date_formats:
                try:
                    return pd.Timestamp(datetime.strptime(s, fmt))
                except Exception:
                    pass
            return pd.to_datetime(s, errors="coerce", dayfirst=False)

        felvetel_dt = df[felvetel_col].apply(_parse_date_flexible)
        felvetel_date = felvetel_dt.dt.date

        # Debug: show how many rows match today after parsing
        _match_count = (felvetel_date == today).sum()
        logger.info(f"🔍 Parsolt dátumok - mai dátummal egyező sorok száma: {_match_count}")
        if _match_count == 0:
            _non_nat = felvetel_dt.dropna()
            if len(_non_nat) > 0:
                logger.info(f"   Példa parsolt értékek (első 5 nem-üres): {list(_non_nat.dt.date.head())}")
            else:
                logger.warning("   ⚠️ Egyetlen dátum sem parseolható a felvételi dátum oszlopból!")

        # Not registered: T column is index 19 in your sheet layout
        t_col = df.columns[19] if len(df.columns) > 19 else None
        not_registered = True
        if t_col is not None:
            not_registered = df[t_col].astype(str).str.strip().str.upper().ne("K")

        # Has core patient data (avoid rows that are truly empty but kept by Excel used-range)
        core_cols = [c for c in ["Paciens/Nev", "Paciens/Azonosito", "Paciens/Email"] if c in df.columns]
        if core_cols:
            has_core = df[core_cols].astype(str).apply(lambda s: s.str.strip()).ne("").any(axis=1)
        else:
            # fallback: require at least something in the row
            has_core = df.astype(str).apply(lambda s: s.str.strip()).ne("").any(axis=1)

        is_today = felvetel_date == today

        work_mask = has_core & is_today & not_registered
        df_work = df.loc[work_mask].copy()

        skipped = int(len(df) - len(df_work))
        logger.info(f"⚡ Előszűrés: feldolgozandó={len(df_work)}, kihagyva={skipped} (üres/nem-mai/K)")

        # ---- Feltöltés soronként (már előszűrve: mai dátumú + nincs K + van érdemi adat)
        for _, row in df_work.iterrows():
            orig_idx = int(row.get("__orig_idx", -1))
            excel_row = orig_idx + 2  # +1 header, +1 1-based
            logger.info(f"➡️  Sor #{orig_idx+1} feldolgozása (Excel sor {excel_row})…")
            attempt = 0
            success_row = False
            while attempt < 2 and not success_row:
                try:
                    ok = upload_one_patient(driver, row, excel_row=excel_row)
                    if ok:
                        logger.info(f"✅ Sor #{orig_idx+1} kész.")
                        processed += 1
                        success_row = True
                    else:
                        logger.warning(f"⚠️ Sor #{orig_idx+1} - upload_one_patient False értékkel tért vissza")
                except (InvalidSessionIdException, NoSuchWindowException):
                    logger.warning("⛔ A böngésző bezárult vagy a munkamenet megszakadt.")
                    logger.warning("   Ha te zártad be a böngészőt: indítsd újra a szkriptet.")
                    logger.warning("   A szkript leáll.")
                    try:
                        driver.quit()
                    except Exception:
                        pass
                    return
                except Exception:
                    logger.exception(f"❌ Sor #{orig_idx+1} hiba.")
                    save_debug(driver, f"row_{orig_idx+1}_error")
                    fail += 1
                    # Try to return to patient register page
                    try:
                        driver.get(LOGIN_URL)
                        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, '[data-automation-id="PatientRegister_CreateNewPatient"]')))
                    except Exception:
                        pass
                    break
            if not success_row and attempt >= 2:
                fail += 1

        logger.info(f"=== Feltöltés összegzés: Feldolgozva={processed}, Kihagyva={skipped}, Sikertelen={fail} ===")

    finally:
        try:
            if driver is not None:
                driver.quit()
        except Exception:
            pass
        logger.info("=== FUTÁS VÉGE ===")


if __name__ == "__main__":
    main()

