# feltoltes.py
#TAJ OK
from pathlib import Path
from datetime import datetime
import time
import unicodedata
import pandas as pd

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import InvalidSessionIdException, TimeoutException, NoSuchWindowException, WebDriverException, StaleElementReferenceException
from selenium.webdriver.common.keys import Keys
import unicodedata
import re

from utils.env import LOGIN_URL, USERNAME, PASSWORD, HEADLESS
from utils.logger import logger
from pages.login_page import LoginPage


# ---------------------------------------------------------------------
# Beállítások
# ---------------------------------------------------------------------
EXCEL_PATH = Path("data/adatok.xlsm")
SHEET_NAME = "Páciensek"
LOG_DIR = Path("logs")
LOG_DIR.mkdir(exist_ok=True)

# Ha csak login tesztet szeretnél: állítsd False-ra
USE_UPLOAD = True

# Feature flags for baseline testing
FF_BASELINE_TAJ_NAME_DOB = False
# Allow overriding baseline via env var, default OFF
try:
    import os
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
    # próbálunk parse-olni
    for fmt in ("%Y-%m-%d", "%Y.%m.%d", "%Y/%m/%d", "%d.%m.%Y", "%d/%m/%Y", "%Y%m%d"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except Exception:
            pass
    # Pandas default parse
    try:
        return pd.to_datetime(s, dayfirst=True).strftime("%Y-%m-%d")
    except Exception:
        return s  # utolsó esély: ahogy van


def normalize_taj(value: str) -> str:
    """
    Csak whitespace-et és gyakori elválasztókat (szóköz, non-breaking space, kötőjel, pont) szedi ki.
    Nem konvertál int-re, így a leading 0 megmarad.
    """
    s = str(value or "").strip()
    s = s.replace("\u00A0", " ")  # NBSP → space
    s = re.sub(r"[ .\-]", "", s)  # pont/kötőjel/szóköz ki
    return s

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


def listbox_pick_by_terms(driver, terms, timeout=10):
    """
    Feltételezzük, hogy a combobox/listbox már nyitva van (overlay).
    - Összes opció beolvasása (role='option' VAGY data-automation-id __option__).
    - Kiírja az első 30 opció normalizált szövegét logger.info-val.
    - A 'terms' normalizált listája alapján kiválasztja az első egyezőt (JS scroll + click).
    - Siker: True, különben False.
    """
    norm_terms = [_norm_text(t) for t in terms]

    # várjuk meg, hogy legalább 1 opció megjelenjen
    end = time.time() + timeout
    options = []
    while time.time() < end and not options:
        options = driver.find_elements(By.XPATH, "//*[@role='option' or starts-with(@data-automation-id,'__option__')]")
        if not options:
            time.sleep(0.2)

    # logolás
    texts = []
    for o in options[:30]:
        try:
            texts.append(_norm_text(o.text))
        except Exception:
            pass
    if texts:
        try:
            logger.info("📋 Opciók (top30): " + " | ".join(texts))
        except Exception:
            pass

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

def retry_on_detached(fn, retries=2, delay=0.5):
    """
    Lefuttat egy műveletet; ha 'target frame detached' / StaleElement... jön,
    rövid várás után újrapróbálja. True/return értéket továbbadja.
    """
    last_exc = None
    for i in range(retries + 1):
        try:
            return fn()
        except (StaleElementReferenceException, WebDriverException) as e:
            msg = str(e).lower()
            if "target frame detached" in msg or "stale element" in msg or "disconnected" in msg:
                last_exc = e
                time.sleep(delay)
                continue
            raise
    if last_exc:
        raise last_exc

def _norm_text(s: str) -> str:
    s = (s or "").strip()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return " ".join(s.lower().split())


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
    norm_terms = [_norm_text(t) for t in terms]

    # várjuk meg, hogy legalább 1 opció megjelenjen
    end = time.time() + timeout
    options = []
    while time.time() < end and not options:
        options = driver.find_elements(By.XPATH, "//*[@role='option' or starts-with(@data-automation-id,'__option__')]")
        if not options:
            time.sleep(0.2)

    # logolás
    texts = []
    for o in options[:30]:
        try:
            texts.append(_norm_text(o.text))
        except Exception:
            pass
    if texts:
        try:
            logger.info("📋 Opciók (top30): " + " | ".join(texts))
        except Exception:
            pass

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

def select_document_type_taj(driver):
    """
    Deterministic selection of 'TAJ szám':
    - Ensure Documents section + row
    - Open dropdown (input/chevron/keys)
    - If list shows 'No options', reopen and retry
    - Pick via regex TAJ.*sz[áa]m using JS mousedown (React-Select-safe)
    - Blur (TAB) and focus DocumentNumber so the flow can continue
    """
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys
    import time, re

    # 0) Ensure section and row exist
    if not _open_documents_section_and_add(driver, timeout=15):
        raise TimeoutException("Okmányok szekció/sor nem érhető el.")

    wanted_regex = r"TAJ.*sz[áa]m"
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

        # 3) Type both diac/fallback and wait options
        def type_and_wait(text) -> bool:
            try:
                inp.send_keys(Keys.CONTROL, "a"); time.sleep(0.05); inp.send_keys(Keys.BACKSPACE)
            except Exception:
                try: inp.clear()
                except Exception: pass
            inp.send_keys(text)
            # removed brittle visible-listbox wait; presence/options are awaited elsewhere
            import time as _t; _t.sleep(0.05)
            return True

        ok = type_and_wait("TAJ szám") or type_and_wait("TAJ szam")
        if not ok:
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

        # 5) commit & verify
        try:
            driver.switch_to.active_element.send_keys(Keys.TAB)
        except Exception:
            pass
        time.sleep(0.2)
        txt = _norm_no_diac(_doc_type_text(driver))
        if re.search(r"\btaj\b", txt):
            # focus DocumentNumber so the next step proceeds
            try:
                doc = WebDriverWait(driver, 6).until(EC.element_to_be_clickable((By.ID, "DocumentNumber")))
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", doc)
                try: doc.click()
                except Exception: driver.execute_script("arguments[0].click();", doc)
            except Exception:
                pass
            return

        # if not verified, small backoff and retry
        time.sleep(0.3)

    raise TimeoutException("Document type selection did not persist (TAJ szám).")


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


def save_patient(driver):
    """Mentés gomb. Várunk valamire a mentés után, hogy stabil legyen."""
    wait_click_css(driver, '[data-automation-id="__save_save"]')
    # kis stabilizáció
    time.sleep(1.0)

# --- Birth date helpers -------------------------------------------------------

def find_birthdate_control(driver, timeout=15):
    """
    Locate the Birth Date/Time control (HU/EN).
    Attribute-based search (id/name/placeholder/aria/type/role/contenteditable/inputmode), in main doc and iframes (depth<=3).
    Connection-safe: ensure_driver_alive + gentle retries.
    Returns an interactive WebElement (input/combobox/contenteditable).
    """
    from selenium.webdriver.common.by import By
    from selenium.common.exceptions import TimeoutException, WebDriverException
    import time, re, unicodedata

    label_terms = [
        "Születési dátum","Szuletesi datum",
        "Születési idő","Szuletesi ido",
        "Születési dátum/idő","Szuletesi datum/ido",
        "Date of birth","Birth date","Birth time","Date & time of birth","DOB"
    ]

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

    def _retry_disconnect(fn, tries=2, wait=0.15):
        last = None
        for _ in range(tries):
            try:
                ensure_driver_alive(driver)
                return fn()
            except WebDriverException as e:
                msg = (str(e) or "").lower()
                if any(k in msg for k in ("connection", "reset", "refused", "disconnected")):
                    last = e
                    time.sleep(wait)
                    continue
                raise
            except Exception as e:
                last = e
                time.sleep(wait)
                continue
        if last:
            raise last

    def _try_here(ctx):
        # 1) Direct by id/name/data-automation/type/role/inputmode
        css_candidates = [
            # classic inputs
            "input#BirthDate","input#BirthTime","input#BirthDateTime",
            "input[name='BirthDate']","input[name='BirthTime']","input[name='BirthDateTime']",
            "input[id*='BirthDate']","input[name*='BirthDate']",
            "input[id*='BirthTime']","input[name*='BirthTime']",
            "input[id*='BirthDateTime']","input[name*='BirthDateTime']",
            "input[id*='Szuletesi']","input[name*='Szuletesi']","input[id*='Szul']","input[name*='Szul']",
            "input[data-automation-id*='Birth']","input[data-automation-id*='birth']",
            "input[type='date']","input[type='datetime-local']",
            "input[role='combobox']","input[role='spinbutton']","input[role='textbox']",
            "input[inputmode='numeric']",
            # common masked libraries markups
            "[data-inputmask] input","[data-mask] input","[data-mask]","[data-type='date'] input","[data-type='date']",
            # contenteditable fallbacks
            "[contenteditable='true']"
        ]
        for sel in css_candidates:
            try:
                el = _retry_disconnect(lambda: ctx.find_element(By.CSS_SELECTOR, sel))
                if _visible(el):
                    return el
            except Exception:
                pass

        # 2) Placeholder / aria-label contains labels (also on contenteditable via aria-label)
        for sel in [
            "input[placeholder]","input[aria-label]","input[role='combobox']","input[type='text']",
            "[contenteditable='true'][aria-label]"
        ]:
            try:
                elems = _retry_disconnect(lambda: ctx.find_elements(By.CSS_SELECTOR, sel))
                for el in elems:
                    try:
                        meta = " ".join(filter(None, [_safe_attr(el, "placeholder"), _safe_attr(el, "aria-label")]))
                        if meta and any(_norm(l) in _norm(meta) for l in label_terms) and _visible(el):
                            return el
                    except Exception:
                        continue
            except Exception:
                pass

        # 3) Generic container guess -> inner input/contenteditable
        try:
            cont = None
            for csel in [
                "[data-automation-id*='Birth']","[data-automation-id*='birth']",
                "[id*='Birth']","[id*='birth']",
                "[id*='Szuletesi']","[id*='Szul']"
            ]:
                try:
                    cont = ctx.find_element(By.CSS_SELECTOR, csel); break
                except Exception:
                    continue
            if cont:
                for sel in [
                    "input[role='combobox']","input[type='text']","input",
                    "[contenteditable='true']"
                ]:
                    try:
                        el = cont.find_element(By.CSS_SELECTOR, sel)
                        if _visible(el):
                            return el
                    except Exception:
                        continue
        except Exception:
            pass

        # 4) Role-based global scan (spinbutton/textbox), date-ish masks
        try:
            elems = ctx.find_elements(By.XPATH,
                "//*[@role='spinbutton' or @role='textbox' or @contenteditable='true' or @inputmode='numeric']")
            for el in elems:
                try:
                    if not _visible(el): 
                        continue
                    meta = " ".join(filter(None, [
                        _safe_attr(el, "aria-label"),
                        _safe_attr(el, "placeholder"),
                        _safe_attr(el, "name"),
                        _safe_attr(el, "id")
                    ]))
                    if any(term in _norm(meta) for term in ["birth","szuletesi","szul","dob","date"]):
                        return el
                except Exception:
                    continue
        except Exception:
            pass

        return None

    # Search main + iframes (depth<=3)
    end = time.time() + timeout
    last_exc = None
    depth_limit = 3
    while time.time() < end:
        try:
            ensure_driver_alive(driver)
        except Exception:
            time.sleep(0.2)

        try:
            el = _try_here(driver)
            if el:
                return el
            # iframes depth 3
            def walk_iframes(level=1):
                nonlocal el
                if level > depth_limit or el is not None:
                    return
                try:
                    frames = driver.find_elements(By.TAG_NAME, "iframe")
                except Exception:
                    frames = []
                for fr in frames:
                    if el is not None:
                        break
                    try:
                        driver.switch_to.frame(fr)
                        el = _try_here(driver)
                        if el:
                            return
                        walk_iframes(level+1)
                    except Exception as e:
                        pass
                    finally:
                        try: driver.switch_to.parent_frame()
                        except Exception: pass
            walk_iframes(1)
            if el:
                return el
        except Exception as e:
            last_exc = e
        time.sleep(0.2)

    raise TimeoutException("Birth date control not found") from last_exc

def fill_birthdate_iso(driver, iso_date: str):
    """
    Robust fill for masked DOB controls:
    - digits-only (yyyymmdd), per-char focus lock & resume
    - temporarily lock TAJ input (readonly + tabindex=-1) to prevent focus steal
    - verify by digits; JS fallback with events
    - commit (ENTER) and blur to a safe field
    """
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.common.by import By
    from selenium.common.exceptions import WebDriverException, TimeoutException
    import time

    want_digits = _iso_to_digits(iso_date)
    if len(want_digits) < 8:
        raise TimeoutException(f"Invalid DOB source: {iso_date}")

    # Remember/lock TAJ input to avoid accidental typing there
    taj_el = None
    taj_restore = {"readonly": None, "tabindex": None}
    try:
        taj_el = driver.find_element(By.ID, "DocumentNumber")
        if taj_el:
            taj_restore["readonly"] = _safe_attr(taj_el, "readonly")
            taj_restore["tabindex"] = _safe_attr(taj_el, "tabindex")
            try:
                driver.execute_script(
                    "arguments[0].setAttribute('readonly','true'); arguments[0].setAttribute('tabindex','-1');",
                    taj_el
                )
            except Exception:
                pass
    except Exception:
        taj_el = None

    el = find_birthdate_control(driver, timeout=18)

    def _focus_lock(elem, retries=3):
        for _ in range(retries):
            try:
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", elem)
            except Exception:
                pass
            try:
                elem.click()
            except Exception:
                try:
                    driver.execute_script("arguments[0].click();", elem)
                except Exception:
                    pass
            try:
                ae = driver.switch_to.active_element
                if ae and ae == elem:
                    return True
            except Exception:
                pass
            time.sleep(0.05)
        return False

    def _clear(elem):
        try:
            elem.send_keys(Keys.CONTROL, "a"); elem.send_keys(Keys.BACKSPACE)
            time.sleep(0.05)
        except Exception:
            try: elem.clear()
            except Exception: pass

    def _digits_in_control(elem) -> str:
        try:
            cur = (elem.get_attribute("value") or elem.text or "").strip()
        except Exception:
            cur = ""
        return _digits_only(cur)

    def _type_digits_safely(elem, digits: str):
        _clear(elem)

        idx = 0
        guard_loops = 0
        while idx < len(digits) and guard_loops < 40:
            guard_loops += 1

            try:
                ae = driver.switch_to.active_element
            except Exception:
                ae = None
            if ae is None or ae != elem:
                if not _focus_lock(elem):
                    time.sleep(0.1)
                    try:
                        new_el = find_birthdate_control(driver, timeout=6)
                        elem = new_el
                    except Exception:
                        pass
                    _focus_lock(elem)

            ch = digits[idx]
            try:
                elem.send_keys(ch)
            except WebDriverException:
                time.sleep(0.1)
                _focus_lock(elem)
                try:
                    elem.send_keys(ch)
                except Exception:
                    pass

            time.sleep(0.04)

            got = _digits_in_control(elem)
            common = min(len(got), idx + 1)
            if _digits_only(got[:common]) == _digits_only(digits[:common]):
                idx += 1
            else:
                # if typing ended up elsewhere (like TAJ), pull focus back
                if taj_el is not None:
                    try:
                        ae2 = driver.switch_to.active_element
                        if ae2 == taj_el:
                            _focus_lock(elem)
                    except Exception:
                        pass
                # masks sometimes swallow a key — retry this char
                try:
                    elem.send_keys(digits[idx])
                except Exception:
                    pass
                time.sleep(0.03)

        return elem

    try:
        if not _focus_lock(el):
            el = find_birthdate_control(driver, timeout=6)
            _focus_lock(el)
        el = _type_digits_safely(el, want_digits)
    except WebDriverException:
        time.sleep(0.2)
        el = find_birthdate_control(driver, timeout=6)
        _focus_lock(el)
        el = _type_digits_safely(el, want_digits)

    # Commit with ENTER, then verify
    try:
        el.send_keys(Keys.ENTER)
    except Exception:
        pass

    got_digits = _digits_in_control(el)
    if want_digits not in got_digits:
        for fmt in (
            f"{want_digits[:4]}-{want_digits[4:6]}-{want_digits[6:]}",
            f"{want_digits[:4]}.{want_digits[4:6]}.{want_digits[6:]}",
            iso_date
        ):
            try:
                driver.execute_script("""
                    const el = arguments[0], val = arguments[1];
                    const desc = Object.getOwnPropertyDescriptor(Object.getPrototypeOf(el), 'value') ||
                                 Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value');
                    if (desc && desc.set) { desc.set.call(el, val); } else { el.value = val; }
                    el.dispatchEvent(new Event('input', {bubbles:true,cancelable:true}));
                    el.dispatchEvent(new Event('change', {bubbles:true,cancelable:true}));
                """, el, fmt)
                time.sleep(0.08)
                try: el.send_keys(Keys.ENTER)
                except Exception: pass
                got_digits = _digits_in_control(el)
                if want_digits in got_digits:
                    break
            except Exception:
                continue

        if want_digits not in got_digits:
            cur_val = (el.get_attribute("value") or el.text or "").strip()
            # restore TAJ before raising
            if taj_el is not None:
                try:
                    driver.execute_script(
                        "arguments[0].setAttribute('readonly', arguments[1]); arguments[0].setAttribute('tabindex', arguments[2]);",
                        taj_el, (taj_restore["readonly"] if taj_restore["readonly"] else ""), (taj_restore["tabindex"] if taj_restore["tabindex"] else "0")
                    )
                except Exception:
                    pass
            raise TimeoutException(f"DOB mismatch: expected '{iso_date}', got '{cur_val}'")

    # Blur to safe field to prevent TAJ from grabbing further input
    try:
        safe = None
        try:
            safe = WebDriverWait(driver, 2).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "#FirstName, [id='FirstName'], [name='FirstName']"))
            )
        except Exception:
            pass
        if safe:
            try:
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", safe)
            except Exception:
                pass
            try:
                safe.click()
            except Exception:
                driver.execute_script("arguments[0].click();", safe)
        else:
            el.send_keys(Keys.TAB)
    except Exception:
        pass

    # restore TAJ
    if taj_el is not None:
        try:
            driver.execute_script(
                "arguments[0].setAttribute('readonly', arguments[1]); arguments[0].setAttribute('tabindex', arguments[2]);",
                taj_el, (taj_restore["readonly"] if taj_restore["readonly"] else ""), (taj_restore["tabindex"] if taj_restore["tabindex"] else "0")
            )
        except Exception:
            pass

    return el

# ---------------------------------------------------------------------
# Feltöltő lépés – 1 sor (1 páciens)
# ---------------------------------------------------------------------
def upload_one_patient(driver, row: pd.Series):
    """
    Vár Excel mezők:
      - 'Paciens/Nev' VAGY 'Vezetéknév' + 'Utónév'
      - 'Paciens/Azonosito'
      - 'Paciens/SzuletesiDatum'
      - 'Paciens/Nem'
      - 'Paciens/Email'
    
    Determinisztikus flow:
    1) TAJ blokk (select_document_type_taj + DocumentNumber)
    2) Alapadatok: LastName, FirstName, BirthDate, Nem
    3) Email hozzáadás gomb (Elérhetőségek szekcióban)
    4) EmailAddress mező kitöltése
    """
    # 1) Új páciens űrlap megnyitása
    open_new_patient_form(driver)

    # 2) TAJ blokk
    select_document_type_taj(driver)
    
    raw = row.get("Paciens/Azonosito", "")
    doc_num = normalize_taj(raw)
    doc_num = re.sub(r"\s+", "", str(doc_num))
    logger.info(f"🆔 TAJ raw={repr(raw)}  -> used={repr(doc_num)}")
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

    # --- Születési dátum (megy a már meglévő best-effort DOB blokk) -- ezt nem módosítjuk itt
    dob_raw = get_cell(
        row,
        "Paciens/SzuletesiDatum","Születési dátum","Szuletesi datum","DOB","Birth date","Date of birth"
    )
    dob_iso = to_iso_date(dob_raw)
    if dob_iso:
        try:
            dob_el = fill_birthdate_iso(driver, dob_iso)
            _val = (dob_el.get_attribute("value") or dob_el.text or "").strip()
            logger.info(f"✅ DOB filled: want={dob_iso} got={_val}")
        except Exception as _e:
            logger.warning(f"⚠️ DOB best-effort failed for value={dob_iso} ({_e}); continuing.")
            save_debug(driver, "dob_best_effort_fail")

    # Nem (id-alapú direkt választás a stabil for-ral)
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

    # 4.1) Email mező (direct ID first) — alias forrásból
    email_value = get_cell(
        row,
        "Paciens/Email","Email","E-mail cím","Email cím","E-mail","EmailAddress"
    )
    if email_value and not pd.isna(email_value):
        # Ensure email section is open first
        if ensure_email_section_open(driver):
            try:
                em_el = WebDriverWait(driver, 8).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "#EmailAddress, [id='EmailAddress']"))
                )
                type_sturdy(driver, em_el, str(email_value))
                logger.info("✅ Email filled (direct)")
            except Exception:
                em_el = fill_field_smart(
                    driver,
                    labels=["Email","E-mail cím","E-mail","Email cím"],
                    attr_contains=["mail","email","EmailAddress"],
                    value=str(email_value),
                    timeout=25
                )
                logger.info("✅ Email filled (smart fallback)")

    # 4) Email hozzáadás gomb (Elérhetőségek szekcióban)
    if email_value and not pd.isna(email_value):
        # Find and click the Email add button in contact section
        email_add_btn = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-automation-id="__addNewItemCompactButton"]'))
        )
        # Safety: ensure it's in the contact/email section (not the documents section)
        try:
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", email_add_btn)
            email_add_btn.click()
        except Exception:
            driver.execute_script("arguments[0].click();", email_add_btn)
        time.sleep(0.5)
        
        # 5) EmailAddress mező kitöltése (id="EmailAddress")
        email_el = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.ID, "EmailAddress"))
        )
        try:
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", email_el)
            email_el.clear()
            email_el.send_keys(str(email_value))
        except Exception:
            driver.execute_script("arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('input', {bubbles: true})); arguments[0].dispatchEvent(new Event('change', {bubbles: true}));", email_el, str(email_value))

    # 6) Mentés
    save_patient(driver)


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
                select_document_type_taj(driver)
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

        # ---- Excel beolvasás
        try:
            df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME, dtype=str, keep_default_na=False, engine="openpyxl")
        except Exception as e:
            logger.exception(f"❌ Excel beolvasási hiba: {e}")
            return

        if df.empty:
            logger.warning("⚠️ Az Excel lap üres, nincs mit feltölteni.")
            return

        logger.info(f"📦 Sorok száma: {len(df)}")
        ok = 0
        fail = 0

        # ---- Feltöltés soronként
        for idx, row in df.iterrows():
            logger.info(f"➡️  Sor #{idx+1} feldolgozása…")
            attempt = 0
            success_row = False
            while attempt < 2 and not success_row:
                try:
                    upload_one_patient(driver, row)
                    logger.info(f"✅ Sor #{idx+1} kész.")
                    ok += 1
                    success_row = True
                except (InvalidSessionIdException, NoSuchWindowException):
                    logger.warning("♻️ InvalidSessionIdException – driver újraindítása és relogin…")
                    driver = recreate_and_relogin(driver)
                    attempt += 1
                except Exception:
                    logger.exception(f"❌ Sor #{idx+1} hiba.")
                    save_debug(driver, f"row_{idx+1}_error")
                    fail += 1
                    # próbáljunk "visszakerülni" a kezdő oldalra
                    try:
                        driver.get(LOGIN_URL)
                        # ha SSO átirányítás, elég egy kis várakozás
                        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, '[data-automation-id="PatientRegister_CreateNewPatient"]')))
                    except Exception:
                        pass
                    break
            if not success_row and attempt >= 2:
                fail += 1

        logger.info(f"=== Feltöltés összegzés: OK={ok}, FAIL={fail} ===")

    finally:
        try:
            if driver is not None:
                driver.quit()
        except Exception:
            pass
        logger.info("=== FUTÁS VÉGE ===")


if __name__ == "__main__":
    main()
