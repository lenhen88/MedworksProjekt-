# -*- coding: utf-8 -*-
"""
TÉTELEK FELVITELE MODUL

- A tételek bepipálása után MEGÁLL; a felhasználónak kell megnyomnia a mentés gombot
- Mentés után a keresés ikonra kattint (következő páciens); K betű az Excelbe win32com-mal
- `python tetelek.py` (standalone) a végén bezárja a böngészőt; más belépési pont eltérhet
- XPath alapú checkbox keresés (MINDEN HTML elem típust támogat)
- D oszlop (KÓD) alapú keresés a szűrőmezőben
"""

from pathlib import Path
import time
import re
import os
import sys
import json
import unicodedata
from typing import List, Dict, Tuple, Optional, Callable, Any
from datetime import datetime, timedelta, date

import pandas as pd
import winsound
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchWindowException,
    WebDriverException,
    StaleElementReferenceException,
    InvalidSessionIdException,
)

from utils.logger import logger
from pages.login_page import LoginPage


def _exc_brief(e: Exception) -> str:
    """Short one-line summary for noisy Selenium exceptions."""
    try:
        s = str(e) or e.__class__.__name__
        first = s.replace("\r", "\n").split("\n", 1)[0].strip()
        return first or e.__class__.__name__
    except Exception:
        return e.__class__.__name__


def load_config() -> dict:
    """Load configuration from config.json next to executable."""
    if getattr(sys, 'frozen', False):
        base_path = os.path.dirname(sys.executable)
    else:
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


CONFIG = load_config()

LOGIN_URL = CONFIG['login_url']
USERNAME = CONFIG['username']
PASSWORD = CONFIG['password']
HEADLESS = CONFIG.get('headless', False)

EXCEL_PATH = Path(CONFIG['excel_path'])
SHEET_NAME = CONFIG['sheet_name']
DEBUG_VERBOSE = os.getenv("TETELEK_DEBUG", "").strip().lower() in ("1", "true", "yes", "on")

# Dátumformátumok (egységesen használva az összes parse helyen)
_DATE_FORMATS = [
    "%Y-%m-%d",
    "%Y.%m.%d",
    "%Y/%m/%d",
    "%d.%m.%Y",
    "%d-%m-%Y",
]


# ========== SEGÉDFÜGGVÉNYEK ==========

def _norm_txt(s: str) -> str:
    """Normalizált szöveg: kisbetűs, diakritika nélkül."""
    s = str(s or "")
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return re.sub(r"\s+", " ", s.strip().lower())


def _visible(el) -> bool:
    """Biztonságos láthatóság ellenőrzés."""
    try:
        return el and el.is_displayed()
    except (StaleElementReferenceException, Exception):
        return False


def _parse_date(raw_date) -> Optional[date]:
    """
    Egységes dátum-parse helper.
    Kezeli: pd.Timestamp, datetime, string (több formátum), Excel numerikus dátum.
    Visszaad: date objektumot, vagy None ha nem értelmezhető.
    """
    if raw_date is None or str(raw_date).strip() == "":
        return None

    if isinstance(raw_date, pd.Timestamp):
        return raw_date.date()

    if isinstance(raw_date, datetime):
        return raw_date.date()

    date_str = str(raw_date).strip()
    date_part = date_str.split()[0] if ' ' in date_str else date_str

    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(date_part, fmt).date()
        except ValueError:
            continue

    # Excel numerikus dátum fallback
    try:
        excel_num = float(date_str.replace(",", "."))
        if excel_num > 40000:
            return (datetime(1899, 12, 30) + timedelta(days=excel_num)).date()
    except Exception:
        pass

    return None


def expand_all_categories(driver):
    """Összes kategória megnyitása hogy minden tétel látható legyen."""
    logger.info("🔓 ÖSSZES kategória megnyitása...")
    try:
        expand_buttons = driver.find_elements(
            By.XPATH,
            "//button[contains(@class, 'expand') or .//svg[contains(@class, 'chevron')] or contains(@aria-label, 'Expand')]"
        )

        opened = 0
        for btn in expand_buttons:
            try:
                if _visible(btn):
                    aria_expanded = btn.get_attribute("aria-expanded")
                    if aria_expanded == "false" or aria_expanded is None:
                        safe_click(driver, btn)
                        opened += 1
                        time.sleep(0.2)
            except Exception:
                continue

        logger.info(f"✅ {opened} kategória megnyitva")
        time.sleep(0.5)
        return True
    except Exception as e:
        logger.debug(f"Kategória nyitás hiba: {e}")
        return False


def safe_click(driver, element, retries: int = 3):
    """Biztonságos kattintás retry mechanizmussal."""
    for attempt in range(retries):
        try:
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", element)
            time.sleep(0.1)
            element.click()
            return True
        except StaleElementReferenceException:
            if attempt < retries - 1:
                logger.debug(f"♻️ Stale element, újrapróbálás... ({attempt + 1}/{retries})")
                time.sleep(0.3)
                continue
            raise
        except Exception:
            try:
                driver.execute_script("arguments[0].click();", element)
                return True
            except Exception as e:
                if attempt == retries - 1:
                    raise e
                time.sleep(0.2)
    return False


def make_driver() -> webdriver.Chrome:
    """Chrome driver létrehozása."""
    opts = Options()
    if HEADLESS:
        opts.add_argument("--headless=new")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--window-size=1400,900")
    opts.page_load_strategy = 'normal'
    return webdriver.Chrome(options=opts)


# ========== EXCEL PARSING ==========

# Visszatérési típus: (név, [(kód, megnevezés), ...], dátum, okmányszám, [excel_sorok])
PatientTuple = Tuple[str, List[Tuple[str, str]], Any, Optional[str], List[int]]


def _read_excel_df(path: Path, sheet: str, header) -> pd.DataFrame:
    """Egyetlen helyen: openpyxl + dtype, hogy ne legyen két eltérő read_excel hívás."""
    return pd.read_excel(
        path,
        sheet_name=sheet,
        dtype=str,
        keep_default_na=False,
        engine="openpyxl",
        header=header,
    )


def parse_items_excel(path: Path, sheet: str) -> List[PatientTuple]:
    """
    Excel beolvasása rugalmasan.
    Visszaad: [(név, [(kód, megnevezés), ...], dátum, okmányszám, [excel_sorok]), ...]
    """
    logger.info(f"📥 Excel betöltés: {path} | lap: {sheet}")
    if not path.exists():
        logger.warning(f"⚠️ Excel nem található: {path}")
        return []

    # ELSŐDLEGES: Fejléc nélküli formátum (A=Név, D=Kód, E=Megnevezés)
    try:
        df = _read_excel_df(path, sheet, header=None)
        if not df.empty and len(df.columns) >= 5:
            first_row_A = str(df.iloc[0, 0] if len(df) > 0 else "").strip()
            if first_row_A and _norm_txt(first_row_A) not in ("nev", "név", "paciens", "páciens"):
                result = _parse_without_header(df)
                if result:
                    logger.info("📋 Fejléc nélküli formátum használva")
                    return result
    except Exception as e:
        logger.debug(f"Fejléc nélküli olvasás sikertelen: {e}")

    # FALLBACK: Fejléces mód
    try:
        df_h = _read_excel_df(path, sheet, header=0)
        if not df_h.empty:
            result = _parse_with_header(df_h)
            if result:
                logger.info("📋 Fejléces formátum használva")
                return result
    except Exception as e:
        logger.debug(f"Fejléces olvasás sikertelen: {e}")

    logger.warning("⚠️ Egyik formátum sem működött")
    return []


def _parse_with_header(df: pd.DataFrame) -> List[PatientTuple]:
    """Fejléces Excel feldolgozása."""
    name_keys = ["Név", "Nev", "Paciens/Nev", "Páciens név", "Full name", "Patient"]
    name_col = None

    for k in name_keys:
        if k in df.columns:
            name_col = k
            break

    if not name_col:
        for c in df.columns:
            if _norm_txt(c) in ("nev", "név", "paciens"):
                name_col = c
                break

    if not name_col:
        logger.warning("⚠️ Név oszlop nem található")
        return []

    out = []
    truthy = {"1", "x", "X", "true", "TRUE", "igen", "Igen"}

    for idx, row in df.iterrows():
        name_val = str(row.get(name_col, "")).strip()
        if not name_val:
            continue

        items = []

        if "Tételek" in df.columns:
            raw = str(row.get("Tételek", "") or "").strip()
            if raw:
                parts = [p.strip() for p in raw.split(";") if p.strip()]
                items.extend([(p, p) for p in parts])

        for col in df.columns:
            if col == name_col or _norm_txt(col) in ("tételek",):
                continue
            val = str(row.get(col, "") or "").strip()
            if val in truthy:
                items.append((col, col))

        items = list(dict.fromkeys(items))
        # header=0 → Excel sor = idx + 2 (fejléc + 1-based)
        out.append((name_val, items, None, None, [idx + 2]))

    logger.info(f"📊 Betöltve: {len(out)} páciens (fejléces)")
    return out


def extract_code_from_parentheses(text: str) -> Optional[str]:
    """
    Kinyeri a zárójelben lévő kódot (max 10 karakter, nem ár).

    Példák:
    - "Vérkép (VK-22) - 400 Ft" -> "VK-22"
    - "Alkalikus foszfatáz (AP) - 400 Ft" -> "AP"
    """
    text = str(text or "").strip()
    matches = re.findall(r'\(([^)]+)\)', text)

    if matches:
        candidates = []
        for match in matches:
            clean = match.strip()
            is_price = bool(re.search(r'\d+\s*ft', clean.lower())) or clean.lower() == "ft"
            if len(clean) <= 10 and not is_price and "forint" not in clean.lower():
                candidates.append(clean)

        if candidates:
            return min(candidates, key=len)

    return None


def _parse_without_header(df: pd.DataFrame) -> List[PatientTuple]:
    """
    Fejléc nélküli Excel: A=Név, C=Okmányszám, D=Kód, E=Megnevezés, I=Dátum, K=Feldolgozva.

    Csak mai dátumú, K betű nélküli sorokat dolgoz fel.
    Visszaad: [(név, [(D_kód, E_megnevezés), ...], dátum, okmányszám, [excel_sorok]), ...]
    """
    if df.empty:
        return []

    today = datetime.now().date()
    logger.info(f"📅 Dátum szűrés aktív: MAI DÁTUM = {today.strftime('%Y.%m.%d')}")
    logger.info(
        "🔤 K oszlop szűrés aktív: a K oszlopban már «K» levő sorok kihagyva (feldolgozottnak jelölt)."
    )

    name_to_items: Dict[str, List[Tuple[str, str]]] = {}
    name_to_date: Dict[str, Any] = {}
    name_to_doc: Dict[str, str] = {}
    name_to_rows: Dict[str, List[int]] = {}
    skipped_count = 0
    skipped_k_count = 0

    for idx, row in df.iterrows():
        name_val = str(row.iloc[0] if len(row) > 0 else "").strip()
        if not name_val:
            continue

        # Okmányszám (C oszlop = index 2)
        doc_number = str(row.iloc[2] if len(row) > 2 else "").strip()
        if doc_number.endswith(".0") and doc_number[:-2].isdigit():
            doc_number = doc_number[:-2]
        if re.match(r"^\d{8}$", doc_number):
            doc_number = "0" + doc_number

        # Dátum ellenőrzés (I oszlop = index 8)
        raw_date = row.iloc[8] if len(row) > 8 else None
        date_only = _parse_date(raw_date)

        if date_only is None:
            logger.debug(f"⏭️ Kihagyva (nincs/érvénytelen dátum): {name_val}")
            skipped_count += 1
            continue

        if date_only != today:
            logger.debug(f"⏭️ Kihagyva (nem mai: {date_only}): {name_val}")
            skipped_count += 1
            continue

        # K betű ellenőrzés (K oszlop = index 10)
        k_oszlop = str(row.iloc[10] if len(row) > 10 else "").strip()

        if DEBUG_VERBOSE:
            logger.info(f"🔬 DEBUG K oszlop [{name_val}]: '{k_oszlop}' → upper='{k_oszlop.upper()}'")

        if k_oszlop.upper() == "K":
            logger.info(f"⏭️ KIHAGYVA (már feldolgozva, K betű): {name_val}")
            skipped_count += 1
            skipped_k_count += 1
            continue

        logger.debug(f"✅ Feldolgozásra kerül: {name_val}")

        # Tételek (D és E oszlop)
        kod = str(row.iloc[3] if len(row) > 3 else "").strip()
        megnevezes = str(row.iloc[4] if len(row) > 4 else "").strip()

        if kod or megnevezes:
            search_code = kod if kod else megnevezes
            item_tuple = (search_code, megnevezes if megnevezes else kod)

            if name_val not in name_to_items:
                name_to_items[name_val] = []
                name_to_date[name_val] = raw_date
                name_to_doc[name_val] = doc_number
                name_to_rows[name_val] = []

            if item_tuple not in name_to_items[name_val]:
                name_to_items[name_val].append(item_tuple)

            # header=None → Excel sor = idx + 1 (1-based)
            excel_row = idx + 1
            if excel_row not in name_to_rows[name_val]:
                name_to_rows[name_val].append(excel_row)

    out = [
        (name, items, name_to_date[name], name_to_doc.get(name, ""), name_to_rows.get(name, []))
        for name, items in name_to_items.items()
    ]
    logger.info(f"📊 Betöltve: {len(out)} páciens (MAI DÁTUM + nincs K a K oszlopban)")
    if skipped_count > 0:
        other = skipped_count - skipped_k_count
        logger.info(
            f"⏭️ Kihagyva: {skipped_count} sor — "
            f"K betű (már feldolgozott): {skipped_k_count}, "
            f"nem mai / dátum hiba: {other} (I oszlop)"
        )
    return out


# ========== TÉTELEK JELÖLÉSE ==========

def is_checkbox_checked(checkbox) -> bool:
    """Checkbox állapot ellenőrzése."""
    try:
        aria = (checkbox.get_attribute("aria-checked") or "").lower()
        if aria in ("true", "1", "mixed"):
            return True

        if checkbox.get_attribute("type") == "checkbox" and checkbox.is_selected():
            return True

        if checkbox.get_attribute("checked") is not None:
            return True

        cls = (checkbox.get_attribute("class") or "").lower()
        if any(word in cls for word in ["checked", "selected", "is-checked", "is-selected"]):
            return True

        data_checked = checkbox.get_attribute("data-checked")
        if data_checked and data_checked.lower() in ("true", "1", "yes"):
            return True

        try:
            parent = checkbox.find_element(By.XPATH, "..")
            parent_cls = (parent.get_attribute("class") or "").lower()
            if any(word in parent_cls for word in ["checked", "selected", "is-checked"]):
                return True
        except Exception:
            pass

        return False
    except Exception as e:
        logger.debug(f"is_checkbox_checked hiba: {e}")
        return False


def find_checkbox_by_label(driver, label: str, timeout: int = 10) -> Optional[any]:
    """Checkbox keresése felirat alapján – fallback módszer."""
    norm_label = _norm_txt(label)
    end = time.time() + timeout
    logger.debug(f"🔍 Fallback: label keresés '{label}'")

    while time.time() < end:
        try:
            labels = driver.find_elements(By.TAG_NAME, "label")
            labels.extend(driver.find_elements(By.XPATH, "//*[@role='checkbox']//.."))

            for lbl in labels:
                if not _visible(lbl):
                    continue

                text = _norm_txt(lbl.text)
                if norm_label not in text and text not in norm_label:
                    continue

                try:
                    for_id = lbl.get_attribute("for")
                    if for_id:
                        cb = driver.find_element(By.ID, for_id)
                        if _visible(cb):
                            return cb
                except Exception:
                    pass

                try:
                    container = lbl.find_element(By.XPATH, "ancestor::*[1]")
                    checkboxes = container.find_elements(
                        By.CSS_SELECTOR,
                        "input[type='checkbox'], [role='checkbox']"
                    )
                    for cb in checkboxes:
                        if _visible(cb):
                            return cb
                except Exception:
                    pass

        except StaleElementReferenceException:
            time.sleep(0.2)
            continue
        except Exception:
            pass

        time.sleep(0.2)

    return None


def get_fresh_service_filter(driver, timeout: int = 5):
    """
    SZIGORÚ serviceFilterTextBox keresés.
    Csak INPUT elemet fogad el, DIV-et NEM.
    """
    end = time.time() + timeout

    while time.time() < end:
        try:
            all_text_inputs = driver.find_elements(
                By.XPATH, "//input[@type='text' or @type='search' or not(@type)]"
            )

            for inp in all_text_inputs:
                try:
                    if not inp.is_displayed() or inp.tag_name.lower() != "input":
                        continue

                    inp_id = inp.get_attribute("id") or ""
                    inp_class = inp.get_attribute("class") or ""
                    inp_placeholder = inp.get_attribute("placeholder") or ""
                    inp_automation = inp.get_attribute("data-automation-id") or ""

                    if any(kw in inp_id.lower() for kw in ["diagnosis", "diagnóz"]):
                        continue
                    if any(kw in inp_class for kw in ["__single-value", "css-1uccc91"]):
                        continue
                    if "__option" in inp_automation:
                        continue

                    if "serviceFilterTextBox" in inp_id:
                        logger.debug(f"✅ Talált: ID={inp_id}")
                        return inp

                    if "serviceFilterTextBox" in inp_automation:
                        logger.debug(f"✅ Talált: automation-id={inp_automation}")
                        return inp

                    if any(kw in inp_placeholder.lower() for kw in ["szolgáltat", "service", "filter", "szűr"]):
                        if "diagnóz" not in inp_placeholder.lower():
                            logger.debug(f"✅ Talált: placeholder={inp_placeholder}")
                            return inp

                except Exception:
                    continue

        except Exception:
            pass

        time.sleep(0.2)

    logger.debug("serviceFilterTextBox nem található ezen az oldalon")
    return None


def _reset_service_filter(driver):
    """Szűrő törlése és alaphelyzetbe állítása."""
    try:
        fresh_filter = get_fresh_service_filter(driver, timeout=3)
        if fresh_filter:
            fresh_filter.clear()
            fresh_filter.send_keys(Keys.ENTER)
            time.sleep(1.5)
            logger.info("✅ Szűrő alaphelyzetbe")
        else:
            logger.warning("⚠️ Filter mező nem található a törléshez")
    except Exception as e:
        logger.debug(f"Szűrő törlés hiba: {e}")


def verify_service_filter_active(driver, expected_input) -> bool:
    """Ellenőrzi hogy a HELYES mező aktív-e."""
    try:
        active = driver.switch_to.active_element

        if active.tag_name.lower() != "input":
            logger.error(f"❌ AKTÍV ELEM NEM INPUT! Tag: {active.tag_name}")
            return False

        active_id = active.get_attribute("id") or ""
        expected_id = expected_input.get_attribute("id") or ""

        if active_id != expected_id:
            logger.error(f"❌ ROSSZ MEZŐ AKTÍV! Aktív: {active_id} | Kellene: {expected_id}")
            return False

        logger.info(f"✅ HELYES mező aktív: {active_id}")
        return True

    except Exception as e:
        logger.error(f"❌ Aktív elem ellenőrzés hiba: {e}")
        return False


def find_and_check_item_via_filter(driver, item_code: str, item_name: str, timeout: int = 12) -> bool:
    """
    XPath alapú keresés – csak zárójelben lévő kódra keres.

    1. Beírja a D oszlop kódját a szűrőmezőbe
    2. XPath-tal megkeresi az elemeket ahol ZÁRÓJELBEN szerepel a kód
    3. Checkbox keresés flexibilisen, akár 5 szint feljebb is
    """
    search_term = item_code.strip() if item_code else item_name.strip()
    logger.info(f"🔍 Tétel keresése (CSAK zárójel): '{search_term}'")

    # Fókusz tisztítás
    try:
        logger.info("🧹 Fókusz tisztítás...")
        try:
            active = driver.switch_to.active_element
            active_tag = active.tag_name.lower()
            active_class = active.get_attribute("class") or ""
            if active_tag == "div" or "__single-value" in active_class:
                logger.warning("⚠️ Dropdown aktív, blur!")
                driver.execute_script("arguments[0].blur();", active)
                time.sleep(0.2)
        except Exception:
            pass

        body = driver.find_element(By.TAG_NAME, "body")
        driver.execute_script("arguments[0].click();", body)
        time.sleep(0.5)
    except Exception as e:
        logger.debug(f"Fókusz tisztítás hiba: {e}")

    filter_input = get_fresh_service_filter(driver, timeout=5)
    if not filter_input:
        logger.error("❌ serviceFilterTextBox NEM található!")
        return False

    logger.info(f"✅ Filter mező megtalálva: ID={filter_input.get_attribute('id')}")

    # Fókusz beállítás (max 3 kísérlet)
    try:
        for attempt in range(3):
            driver.execute_script("arguments[0].focus();", filter_input)
            time.sleep(0.2)
            driver.execute_script("arguments[0].click();", filter_input)
            time.sleep(0.3)

            if verify_service_filter_active(driver, filter_input):
                break

            if attempt < 2:
                logger.warning(f"⚠️ Fókusz probléma, újra... ({attempt+1}/3)")
                try:
                    body = driver.find_element(By.TAG_NAME, "body")
                    body.click()
                    time.sleep(0.3)
                except Exception:
                    pass
        else:
            logger.error("❌ 3 próbálkozás után sem sikerült fókuszt állítani!")
            return False
    except Exception as e:
        logger.error(f"Fókusz hiba: {e}")
        return False

    # Szűrő beírása
    try:
        filter_input.clear()
        time.sleep(0.1)
        logger.info(f"⌨️ Szűrés: '{search_term}'")
        for ch in search_term:
            filter_input.send_keys(ch)
            time.sleep(0.02)
        filter_input.send_keys(Keys.ENTER)
        logger.info("   ⏎ ENTER + várakozás...")
        time.sleep(3.5)
    except Exception as e:
        logger.error(f"❌ Szűrés hiba: {e}")
        return False

    norm_code = _norm_txt(search_term)

    for search_attempt in range(6):
        try:
            logger.info(f"   🔍 #{search_attempt + 1}. próbálkozás...")

            try:
                service_title_elements = driver.find_elements(
                    By.CSS_SELECTOR,
                    "span[class*='RequestedServiceListPanel_row-header-title']"
                )
                if not service_title_elements:
                    logger.debug("      RequestedServiceListPanel nem talált, fallback XPath...")
                    service_title_elements = driver.find_elements(
                        By.XPATH,
                        "//*[contains(text(), '(') and contains(text(), ')')]"
                    )
            except Exception as e:
                logger.debug(f"      Keresési hiba, fallback XPath: {e}")
                service_title_elements = driver.find_elements(
                    By.XPATH,
                    "//*[contains(text(), '(') and contains(text(), ')')]"
                )

            elements = [elem for elem in service_title_elements if _visible(elem)]
            logger.info(f"   📦 {len(elements)} LÁTHATÓ szolgáltatás (összes: {len(service_title_elements)})")

            found_elem = None

            for elem in elements:
                try:
                    elem_text = elem.text.strip()
                    if len(elem_text) < 3:
                        continue

                    if "ft" not in elem_text.lower() and elem_text.count('(') < 2:
                        if len(elem_text) < 15:
                            continue

                    extracted_code = extract_code_from_parentheses(elem_text)

                    logger.info(f"      🔬 DEBUG: elem_text='{elem_text[:60]}' | extracted_code='{extracted_code}'")

                    if not extracted_code:
                        continue

                    norm_extracted = _norm_txt(extracted_code)
                    logger.info(f"      🔬 DEBUG: norm_extracted='{norm_extracted}' | norm_code='{norm_code}'")

                    if norm_extracted == norm_code:
                        logger.info(f"   ✅ TALÁLAT! '{elem_text[:80]}' | Zárójeles kód: '{extracted_code}'")
                        found_elem = elem
                        break
                    else:
                        logger.debug(f"      ❌ Nem egyezik: '{norm_extracted}' != '{norm_code}'")

                except StaleElementReferenceException:
                    continue
                except Exception as e:
                    logger.debug(f"   Elem ellenőrzés hiba: {e}")
                    continue

            if not found_elem:
                if search_attempt < 2:
                    logger.info("   🔓 Kategóriák megnyitása...")
                    try:
                        buttons = driver.find_elements(
                            By.XPATH, "//button[.//svg or contains(@class, 'expand')]"
                        )
                        for btn in buttons[:5]:
                            try:
                                if _visible(btn):
                                    safe_click(driver, btn)
                                    time.sleep(0.3)
                            except Exception:
                                continue
                    except Exception:
                        pass
                time.sleep(0.5)
                continue

            # Konténer keresés
            logger.info("   🔍 Konténer elem keresése...")
            row_elem = None

            try:
                row_elem = found_elem.find_element(By.XPATH, "ancestor::tr[1]")
                logger.info("   ✅ <tr> elem megtalálva!")
            except Exception:
                pass

            if not row_elem:
                try:
                    row_elem = found_elem.find_element(
                        By.XPATH,
                        "ancestor::*[.//input[@type='checkbox'] or .//*[@role='checkbox']][1]"
                    )
                    logger.info("   ✅ Div konténer megtalálva (checkbox-szal)!")
                except Exception:
                    pass

            if not row_elem:
                try:
                    row_elem = found_elem.find_element(By.XPATH, "ancestor::*[2]")
                    logger.info("   ⚠️ Fallback: 2. szülő elem használata")
                except Exception:
                    row_elem = found_elem
                    logger.warning("   ⚠️ Konténer NEM található – found_elem használata")

            # Már pipálva?
            try:
                fresh_cbs_check = row_elem.find_elements(By.CSS_SELECTOR, "input[type='checkbox']")
                if fresh_cbs_check and is_checkbox_checked(fresh_cbs_check[0]):
                    logger.info("   ✓ Már pipálva (friss ellenőrzés)")
                    _reset_service_filter(driver)
                    return True
            except Exception:
                pass

            click_success = False

            # Stratégia #1: Kattintás a szöveges elemre
            logger.info("   🎯 Stratégia #1: Kattintás a szöveges elemre (stabil)")
            try:
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", found_elem)
                time.sleep(0.5)
                try:
                    active = driver.switch_to.active_element
                    driver.execute_script("arguments[0].blur();", active)
                    time.sleep(0.2)
                except Exception:
                    pass

                driver.execute_script("arguments[0].click();", found_elem)
                logger.info("      ✓ Kattintva")
                time.sleep(1.0)

                try:
                    fresh_checkboxes = row_elem.find_elements(By.CSS_SELECTOR, "input[type='checkbox']")
                    if fresh_checkboxes and is_checkbox_checked(fresh_checkboxes[0]):
                        logger.info("   ✅✅✅ SIKERES PIPÁLÁS (szöveges elem)!")
                        click_success = True
                except Exception as e:
                    logger.debug(f"      Friss checkbox ellenőrzés hiba: {e}")
            except Exception as e:
                logger.debug(f"      ⚠️ Stratégia #1 hiba: {e}")

            # Stratégia #2: XPath alapú checkbox (relatív pozíció)
            if not click_success:
                logger.info("   🎯 Stratégia #2: XPath alapú checkbox (relatív)")
                try:
                    xpath_checkbox = (
                        ".//preceding::input[@type='checkbox'][1] | "
                        ".//preceding::*[@role='checkbox'][1] | "
                        ".//following::input[@type='checkbox'][1] | "
                        ".//following::*[@role='checkbox'][1] | "
                        ".//ancestor::*[.//input[@type='checkbox']]//input[@type='checkbox'][1] | "
                        ".//ancestor::*[.//*[@role='checkbox']].//*[@role='checkbox'][1]"
                    )
                    relative_cb = found_elem.find_element(By.XPATH, xpath_checkbox)

                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", relative_cb)
                    time.sleep(0.3)
                    try:
                        driver.execute_script("arguments[0].click();", relative_cb)
                    except Exception:
                        relative_cb.click()

                    logger.info("      ✓ XPath checkbox-ra kattintva")
                    time.sleep(1.0)

                    try:
                        verify_cbs = row_elem.find_elements(By.CSS_SELECTOR, "input[type='checkbox']")
                        if verify_cbs and is_checkbox_checked(verify_cbs[0]):
                            logger.info("   ✅✅✅ SIKERES PIPÁLÁS (XPath)!")
                            click_success = True
                    except Exception:
                        pass
                except Exception as e:
                    logger.debug(f"      ⚠️ Stratégia #2 hiba: {e}")

            # Stratégia #3: Kattintás a teljes TR elemre
            if not click_success and row_elem:
                logger.info("   🎯 Stratégia #3: Kattintás a teljes TR elemre")
                try:
                    safe_click(driver, row_elem)
                    logger.info("      ✓ TR elemre kattintva")
                    time.sleep(1.0)

                    fresh_cbs = row_elem.find_elements(By.CSS_SELECTOR, "input[type='checkbox']")
                    if fresh_cbs and is_checkbox_checked(fresh_cbs[0]):
                        logger.info("   ✅✅✅ SIKERES PIPÁLÁS (TR elem)!")
                        click_success = True
                except Exception as e:
                    logger.debug(f"      ⚠️ Stratégia #3 hiba: {e}")

            # Stratégia #4: Direkt checkbox + azonnali kattintás
            if not click_success:
                logger.info("   🎯 Stratégia #4: Direkt checkbox keresés + azonnal kattintás")
                for retry in range(3):
                    try:
                        fresh_cbs = row_elem.find_elements(
                            By.CSS_SELECTOR, "input[type='checkbox'], [role='checkbox']"
                        )
                        if not fresh_cbs:
                            time.sleep(0.3)
                            continue

                        cb = fresh_cbs[0]
                        logger.info(f"      Próba #{retry+1}: checkbox id='{cb.get_attribute('id')}'")

                        driver.execute_script("arguments[0].click();", cb)
                        time.sleep(0.8)

                        verify_cbs = row_elem.find_elements(By.CSS_SELECTOR, "input[type='checkbox']")
                        if verify_cbs and is_checkbox_checked(verify_cbs[0]):
                            logger.info(f"   ✅✅✅ SIKERES PIPÁLÁS (direkt, {retry+1}. próba)!")
                            click_success = True
                            break
                    except StaleElementReferenceException:
                        logger.debug(f"      Stale element #{retry+1} - retry...")
                        time.sleep(0.5)
                    except Exception as e:
                        logger.debug(f"      Próba #{retry+1} hiba: {e}")
                        time.sleep(0.3)

            if click_success:
                logger.info("   ✅✅✅ SIKERES PIPÁLÁS!")
                _reset_service_filter(driver)
                return True
            else:
                logger.error("   ❌❌❌ MINDEN STRATÉGIA SIKERTELEN!")

                if DEBUG_VERBOSE:
                    try:
                        all_cbs = row_elem.find_elements(By.CSS_SELECTOR, "input[type='checkbox']")
                        for i, cb in enumerate(all_cbs[:3]):
                            try:
                                logger.info(f"      CB #{i}: id='{cb.get_attribute('id')}' checked={is_checkbox_checked(cb)}")
                            except Exception:
                                logger.info(f"      CB #{i}: stale or error")
                    except Exception:
                        pass

                _reset_service_filter(driver)
                return False

        except Exception as e:
            logger.debug(f"Próbálkozás hiba: {e}")

        time.sleep(0.5)

    logger.warning(f"⚠️ Sikertelen: {search_term}")
    _reset_service_filter(driver)
    return False


def ensure_checkbox_checked(driver, item_code: str, item_name: str, timeout: int = 10) -> bool:
    """Checkbox bepipálása. Elsődleges: XPath szűrős módszer; fallback: label keresés."""
    logger.info(f"☐ Tétel jelölése: KÓD='{item_code}' | NÉV='{item_name}'")

    if find_and_check_item_via_filter(driver, item_code, item_name, timeout=timeout):
        return True

    logger.info("🔁 Fallback: közvetlen label keresés megnevezéssel...")
    checkbox = find_checkbox_by_label(driver, item_name, timeout=5)

    if not checkbox:
        logger.warning(f"⚠️ Checkbox nem található: {item_name}")
        return False

    if is_checkbox_checked(checkbox):
        logger.info(f"✓ Már pipálva: {item_name}")
        return True

    try:
        safe_click(driver, checkbox)
        time.sleep(0.2)

        if is_checkbox_checked(checkbox):
            logger.info(f"☑️ Bejelölve: {item_name}")
            return True
        else:
            logger.warning(f"⚠️ Jelölés sikertelen: {item_name}")
            return False
    except Exception as e:
        logger.warning(f"⚠️ Hiba a jelölés során ({item_name}): {e}")
        return False


def check_items_on_admit_page(driver, items: List[Tuple[str, str]]) -> List[str]:
    """Tételek bepipálása a felvételi oldalon. Visszaadja a sikertelen tételek nevét."""
    failed = []

    logger.info(f"📦 ÖSSZESEN {len(items)} tétel feldolgozása kezdődik:")
    for idx, (code, name) in enumerate(items, start=1):
        logger.info(f"   {idx}. {code} - {name}")
    logger.info("")

    for idx, (item_code, item_name) in enumerate(items, start=1):
        logger.info(f"🧪 [{idx}/{len(items)}] Tétel feldolgozása: KÓD='{item_code}' | NÉV='{item_name}'")

        ok = ensure_checkbox_checked(driver, item_code, item_name, timeout=8)

        if ok:
            logger.info(f"✅ [{idx}/{len(items)}] Sikeres: {item_name}")
        else:
            logger.warning(f"❌ [{idx}/{len(items)}] Sikertelen: {item_name}")
            failed.append(item_name)

    return failed


# ========== DÁTUMMEZŐ KEZELÉS ==========

def _write_datetime_input(driver, inp, value: str):
    """React datetime input mező beállítása JavaScript-tel."""
    driver.execute_script("""
        const input = arguments[0];
        const value = arguments[1];
        const nativeSetter = Object.getOwnPropertyDescriptor(
            HTMLInputElement.prototype, 'value'
        ).set;
        nativeSetter.call(input, value);
        input.dispatchEvent(new Event('input', {bubbles: true}));
        input.dispatchEvent(new Event('change', {bubbles: true}));
        input.dispatchEvent(new Event('blur', {bubbles: true}));
    """, inp, value)
    time.sleep(0.2)
    try:
        inp.send_keys(Keys.TAB)
    except Exception:
        pass


def set_specimen_date_to_today(driver, timeout: int = 15) -> bool:
    """SpecimenCollectedAt mező beállítása mai dátumra."""
    try:
        value = datetime.now().strftime("%Y. %m. %d. %H:%M")
        inp = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((
                By.CSS_SELECTOR, "input[data-automation-id='__SpecimenCollectedAt']"
            ))
        )
        _write_datetime_input(driver, inp, value)
        logger.info(f"📅 Mintavétel dátum beállítva: {value}")
        return True
    except Exception as e:
        logger.warning(f"⚠️ Mintavétel dátum beállítás sikertelen: {e}")
        return False


def set_specimen_date_to_excel_value(driver, excel_date_str: str, timeout: int = 15) -> bool:
    """SpecimenCollectedAt mező beállítása Excel I oszlop értékéből."""
    logger.info(f"📅 Mintavétel dátum beállítása Excel-ből: {excel_date_str}")

    datetime_formats = [
        "%Y-%m-%d %H:%M:%S.%f",
        "%Y.%m.%d %H:%M:%S.%f",
        "%Y.%m.%d %H:%M",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%Y.%m.%d %H:%M:%S",
    ]

    date_obj = None
    for fmt in datetime_formats:
        try:
            date_obj = datetime.strptime(excel_date_str.strip(), fmt)
            logger.debug(f"   ✓ Parse sikeres ({fmt})")
            break
        except ValueError:
            continue

    if not date_obj:
        logger.warning(f"⚠️ Nem sikerült parse-olni: '{excel_date_str}' – fallback mai időre")
        return set_specimen_date_to_today(driver, timeout)

    try:
        value = date_obj.strftime("%Y. %m. %d. %H:%M")
        inp = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((
                By.CSS_SELECTOR, "input[data-automation-id='__SpecimenCollectedAt']"
            ))
        )
        _write_datetime_input(driver, inp, value)
        logger.info(f"✅ Mintavétel dátum beállítva (Excel): {value}")
        return True
    except Exception as e:
        logger.warning(f"⚠️ Excel dátum beállítás sikertelen: {e}")
        return set_specimen_date_to_today(driver, timeout)


def set_expected_completion_time(driver, timeout: int = 15) -> bool:
    """Mintaszállítás (expectedCompletionTime) mező beállítása mai dátumra 15:00-ra."""
    try:
        value = datetime.now().strftime("%Y. %m. %d. 15:00")
        logger.info(f"🚚 Mintaszállítás dátum beállítása: {value}")

        inp = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((
                By.CSS_SELECTOR, "input[data-automation-id='__expectedCompletionTime']"
            ))
        )
        _write_datetime_input(driver, inp, value)
        logger.info(f"✅ Mintaszállítás dátum beállítva: {value}")
        return True
    except TimeoutException:
        logger.warning("⚠️ Mintaszállítás mező nem található (timeout)")
        return False
    except Exception as e:
        logger.warning(f"⚠️ Mintaszállítás dátum beállítás sikertelen: {e}")
        return False


# ========== PÁCIENS FELDOLGOZÁS ==========

def wait_for_admit_page(driver, timeout: int = 7200) -> bool:
    """
    Várakozás amíg a dolgozó a tételek felvételi oldalára navigál.
    Jelzi: serviceFilterTextBox megjelenik.
    """
    logger.info("⏳ Várakozás a felvételi oldalra...")
    end = time.time() + timeout
    last_diag = 0.0

    while time.time() < end:
        try:
            current_url = driver.current_url
        except Exception:
            logger.warning("⚠️ Böngésző elveszett")
            return False

        if get_fresh_service_filter(driver, timeout=1) is not None:
            logger.info("✅ Felvételi oldal érzékelve")
            return True

        now = time.time()
        if now - last_diag >= 5:
            last_diag = now
            try:
                all_inputs = driver.find_elements(By.TAG_NAME, "input")
                handles = driver.window_handles
                logger.info(
                    f"🔍 Diagnosztika: URL={current_url[:60]}... | "
                    f"ablakok={len(handles)} | input mezők={len(all_inputs)}"
                )
            except Exception as e:
                logger.info(f"🔍 Diagnosztika hiba: {e}")

        time.sleep(0.5)

    logger.warning("⚠️ Timeout: felvételi oldal nem jelent meg")
    return False


def process_patient_items(
    driver,
    name: str,
    items: List[Tuple[str, str]],
    specimen_date: str = None,
    excel_rows: Optional[List[int]] = None,
    mark_saved: Optional[Callable[[], None]] = None,
) -> bool:
    """
    Egy páciens teljes feldolgozása.
    A dolgozó kézzel keresi meg a pácienst és navigál a felvételi oldalra,
    a program érzékeli ezt és elvégzi a tételek bepipálását.
    """
    logger.info(f"📋 Páciens feldolgozása: {name}")

    try:
        try:
            winsound.Beep(800, 400)
        except Exception:
            pass

        if not wait_for_admit_page(driver, timeout=7200):
            logger.error(f"❌ Felvételi oldal nem jelent meg: {name}")
            return False

        time.sleep(0.5)
        expand_all_categories(driver)

        failed = check_items_on_admit_page(driver, items)
        if failed:
            logger.warning(f"⚠️ Nem jelölhető tételek: {', '.join(failed)}")

        if specimen_date:
            set_specimen_date_to_excel_value(driver, specimen_date, timeout=15)
        else:
            set_specimen_date_to_today(driver, timeout=15)

        set_expected_completion_time(driver, timeout=15)

        try:
            winsound.Beep(1000, 500)
        except Exception:
            pass

        logger.info("⏳ Várakozás hogy a 'Mentés és feladás' gombot megnyomd...")

        # seen_active: csak akkor tekintjük inaktívnak a gombot, ha előtte már
        # láttuk aktívan – megakadályozza a hamis korai K-írást, ha a gomb
        # még nem is jelent meg az oldalon (pl. betöltés közben).
        seen_active = False
        inactive_streak = 0
        while True:
            try:
                save_buttons = driver.find_elements(
                    By.XPATH,
                    "//button[.//span[contains(text(), 'Mentés és feladás')]]"
                )
                btn_active = any(
                    _visible(btn) and btn.is_enabled() for btn in save_buttons
                )

                if btn_active:
                    seen_active = True
                    inactive_streak = 0
                elif seen_active:
                    inactive_streak += 1
                    if inactive_streak >= 2:
                        logger.info("✅ Mentés gomb inaktív (2× megerősítve) – mentés sikeres!")
                        break
                # ha seen_active=False: gomb még nem jelent meg, tovább várunk
            except Exception as e:
                logger.debug(f"Detektálás hiba: {e}")

            time.sleep(0.5)

        try:
            if mark_saved is not None:
                logger.info(f"📝 K betű írása (sorok: {excel_rows})...")
                mark_saved()
                # win32com Excel-aktiválás után a fókusz gyakran az Excelnél marad (tálca villanás).
                try:
                    driver.execute_script("window.focus();")
                except Exception:
                    pass
            else:
                logger.warning("⚠️ mark_saved nincs beállítva – K betű nem írható")
        except Exception as e:
            logger.warning(f"⚠️ K betű írás hiba ({name}): {_exc_brief(e)}")

        time.sleep(2.0)

        search_box_found = False
        end_check = time.time() + 15
        while time.time() < end_check:
            try:
                boxes = driver.find_elements(By.ID, "searchPatientTextBox")
                if any(_visible(b) for b in boxes):
                    search_box_found = True
                    break
            except Exception:
                pass
            time.sleep(0.5)

        if not search_box_found:
            logger.warning("⚠️ searchPatientTextBox nem jelent meg – mentés bizonytalan, de K már beírva")
        else:
            logger.info("✅ Mentés megerősítve")

        try:
            icons = driver.find_elements(By.CSS_SELECTOR, "[data-automation-id='searchPatientIcon']")
            if icons:
                icons[0].click()
                logger.info("🔍 Keresés ikonra kattintva – várakozás a következő páciensre")
            else:
                logger.warning("⚠️ Keresés ikon nem található")
        except Exception as e:
            logger.warning(f"⚠️ Keresés ikon kattintás hiba: {_exc_brief(e)}")

        logger.info(f"✅ Feldolgozva: {name}")
        return True

    except Exception as e:
        logger.exception(f"❌ Feldolgozási hiba ({name}): {e}")
        return False


def _write_k_sync(
    pname: str,
    rows: List[int],
    excel_path: Path,
    sheet_name: str,
):
    """
    K betű írás win32com-mal – CSAK a főszálon (Selenium főszál; COM STA).

    win32com.GetActiveObject a ROT-ból veszi az Excel példányt.
    Ha win32com nem elérhető, csak figyelmeztet – soha nem ír fájlba.
    """
    if not rows:
        logger.warning(f"⚠️ K betű: nincs sor ({pname})")
        return

    valid_rows = sorted({int(r) for r in rows if int(r) > 0})
    if not valid_rows:
        return

    target_full = str(Path(excel_path).resolve()).lower().replace("/", "\\")
    target_name = Path(excel_path).name.lower()
    _excel_is_running = False

    # 1. Próba: win32com (főszálon, lásd fent)
    try:
        import win32com.client
        xl = win32com.client.GetActiveObject("Excel.Application")
        _excel_is_running = True

        found_wb = None
        for i in range(1, xl.Workbooks.Count + 1):
            try:
                wb = xl.Workbooks(i)
                wb_full = (wb.FullName or "").lower().replace("/", "\\")
                if wb_full == target_full or wb.Name.lower() == target_name:
                    found_wb = wb
                    break
            except Exception:
                continue

        if found_wb:
            # Teljes A1 cím az Application.Range-en: Application.Range("'Lap'!K123").
            # A mentések után gyakori 0x800A03EC (nem írható / Excel elfoglalt), ha a
            # munkafüzet nincs aktiválva, CutCopyMode aktív, vagy számítás fut – ezért
            # aktiválás, CutCopyMode törlés, rövid várakozás és cellánkénti újrapróbálás.
            #
            # FONTOS (UX): a found_wb.Activate() + ActiveWindow.Activate() + SendKeys(ESC)
            # folyamatosan elrabolja a fókuszt a böngészőtől / a felhasználótól, a tálcán
            # „figyelj rám” (sárga/narancs) villanást okoz, és úgy tűnhet, hogy az Excelben
            # nem látszanak az adatok kattintásra. Ezért ELŐSZÖR megpróbálunk csendes írást
            # (ScreenUpdating/EnableEvents kikapcsolva, Activate és ESC NÉLKÜL); csak a
            # sikertelen sorokra jön a korábbi, agresszívabb útvonal.
            #
            # Ha futás közben másolsz Excelből (pl. okmányszám a következő kereséshez),
            # a CutCopyMode (szaggatott keret) aktív maradhat – ez tipikusan blokkolja a
            # következő COM cellaírásokat. Ezért többször, 0-val is töröljük.
            def _addr_k(row: int) -> str:
                sn = str(sheet_name).replace("'", "''")
                return f"'{sn}'!K{row}"

            def _clear_cut_copy_mode(application) -> None:
                """Másolás/kivágás mód törlése (xlCopy=1, xlCut=2 → 0)."""
                for _ in range(3):
                    try:
                        application.CutCopyMode = 0
                    except Exception:
                        try:
                            application.CutCopyMode = False
                        except Exception:
                            pass
                    time.sleep(0.04)

            def _exit_cell_edit_if_any(application) -> None:
                """
                Dupla katt / F2 miatti cellaszerkesztés megszakítása.
                Ilyen módban a Range.Value gyakran 0x800A03EC-et dob – a CutCopyMode
                törlése erre nem elég. ESC csak az aktivált Excel ablaknak megy.
                """
                try:
                    aw = application.ActiveWindow
                    if aw is not None:
                        aw.Activate()
                except Exception:
                    pass
                time.sleep(0.06)
                for _ in range(2):
                    try:
                        application.SendKeys("{ESC}")
                    except Exception:
                        pass
                    time.sleep(0.06)

            def _wait_calc_idle(application, timeout_s: float = 2.5) -> None:
                deadline = time.monotonic() + timeout_s
                while time.monotonic() < deadline:
                    try:
                        # xlCalculationState.xlDone == 0
                        if int(application.CalculationState) == 0:
                            return
                    except Exception:
                        return
                    time.sleep(0.05)

            def _write_single_k(
                application,
                worksheet,
                row: int,
                addr: str,
                max_attempts: int = 8,
            ) -> bool:
                for attempt in range(max_attempts):
                    _clear_cut_copy_mode(application)
                    try:
                        application.Range(addr).Value = "K"
                        return True
                    except Exception:
                        pass
                    if worksheet is not None:
                        try:
                            worksheet.Range(f"K{row}").Value = "K"
                            return True
                        except Exception:
                            pass
                    time.sleep(0.2)
                return False

            app = found_wb.Application
            ws = None
            try:
                try:
                    ws = found_wb.Worksheets(str(sheet_name))
                except Exception:
                    try:
                        ws = found_wb.Sheets(str(sheet_name))
                    except Exception:
                        ws = None

                try:
                    if bool(getattr(found_wb, "ReadOnly", False)):
                        logger.warning(
                            f"⚠️ A munkafüzet csak olvasható (ReadOnly=True) – "
                            f"K írás sikertelen lehet ({pname})"
                        )
                except Exception:
                    pass

                try:
                    had_clip = int(app.CutCopyMode) != 0
                except Exception:
                    had_clip = False
                _clear_cut_copy_mode(app)
                if had_clip:
                    logger.info(
                        "📋 Excel másolási mód (CutCopyMode) törlve a K írás előtt – "
                        "a futás közbeni másolás (pl. okmányszám) okozhatta"
                    )
                _wait_calc_idle(app)

                def _save_excel_ui_state(application) -> dict:
                    """Képernyőfrissítés / események állapota – mindig vissza kell állítani."""
                    st: dict = {}
                    try:
                        st["ScreenUpdating"] = bool(application.ScreenUpdating)
                    except Exception:
                        st["ScreenUpdating"] = True
                    try:
                        st["EnableEvents"] = bool(application.EnableEvents)
                    except Exception:
                        st["EnableEvents"] = True
                    try:
                        st["DisplayAlerts"] = int(application.DisplayAlerts)
                    except Exception:
                        st["DisplayAlerts"] = None
                    return st

                def _restore_excel_ui_state(application, st: dict) -> None:
                    try:
                        application.ScreenUpdating = st.get("ScreenUpdating", True)
                    except Exception:
                        try:
                            application.ScreenUpdating = True
                        except Exception:
                            pass
                    try:
                        application.EnableEvents = st.get("EnableEvents", True)
                    except Exception:
                        pass
                    try:
                        if st.get("DisplayAlerts") is not None:
                            application.DisplayAlerts = st["DisplayAlerts"]
                    except Exception:
                        pass

                failed_rows: List[int] = []
                ui_st = _save_excel_ui_state(app)
                try:
                    try:
                        app.ScreenUpdating = False
                        app.EnableEvents = False
                        try:
                            app.DisplayAlerts = False
                        except Exception:
                            pass
                    except Exception:
                        pass
                    for r in valid_rows:
                        addr = _addr_k(r)
                        if not _write_single_k(app, ws, r, addr):
                            failed_rows.append(r)
                finally:
                    _restore_excel_ui_state(app, ui_st)

                if failed_rows:
                    logger.info(
                        f"🔧 K betű: csendes írás "
                        f"{len(valid_rows) - len(failed_rows)}/{len(valid_rows)} sor OK – "
                        f"fallback (ablak aktiválás + ESC) {len(failed_rows)} sorra: {pname}"
                    )
                    _clear_cut_copy_mode(app)
                    try:
                        found_wb.Activate()
                    except Exception:
                        pass
                    try:
                        if ws is not None:
                            ws.Activate()
                    except Exception:
                        pass
                    _exit_cell_edit_if_any(app)
                    _clear_cut_copy_mode(app)
                    _wait_calc_idle(app)

                    still_failed: List[int] = []
                    for r in failed_rows:
                        addr = _addr_k(r)
                        if not _write_single_k(app, ws, r, addr):
                            still_failed.append(r)
                    failed_rows = still_failed

                if not failed_rows:
                    logger.info(f"✅ K betű beírva win32com ({len(valid_rows)} cella) – {pname}")
                else:
                    logger.warning(
                        f"⚠️ K betű részben sikertelen ({pname}): "
                        f"{len(failed_rows)}/{len(valid_rows)} sor (pl. {failed_rows[:3]}…). "
                        f"Ellenőrizd: lapvédelem, egyesített K cellák, nincs-e nyitva "
                        f"cellaszerkesztés (F2) az Excelben. Első cím: '{_addr_k(failed_rows[0])}'."
                    )
            except Exception as write_e:
                logger.warning(
                    f"⚠️ Cella írás hiba ({pname}), első cím pl. '{_addr_k(valid_rows[0])}': "
                    f"{_exc_brief(write_e)}",
                    exc_info=True,
                )
            return
        else:
            logger.warning(f"⚠️ Munkafüzet nem található win32com-ban ({pname})")

    except Exception as e:
        logger.warning(f"⚠️ win32com hiba ({pname}): {_exc_brief(e)}")

    # openpyxl fallback szándékosan NINCS – nagy/komplex Excel fájloknál
    # az openpyxl az egész fájlt újraírja és elveszíti a nem támogatott
    # funkciókat (képletek, stílusok, feltételes formázás, named range-ek stb.),
    # ami látszólag "üres" fájlt eredményez.
    # Ha win32com nem érhető el, a felhasználó jelöli kézzel a K betűt.
    if _excel_is_running:
        logger.warning(
            f"⚠️ Excel fut, de munkafüzet nem érhető el.\n"
            f"   K betű NEM lett beírva ({pname}) – kézi ellenőrzés szükséges!\n"
            f"   Sorok: {valid_rows}"
        )
    else:
        logger.warning(
            f"⚠️ Excel nem fut – K betű NEM lett beírva ({pname}).\n"
            f"   Nyisd meg az Excel fájlt és jelöld kézzel: sorok {valid_rows}"
        )


def admit_and_check_items_from_excel(driver, excel_path: Path, sheet_name: str):
    """Teljes folyamat az Excelből."""
    pairs = parse_items_excel(excel_path, sheet_name)

    if not pairs:
        logger.warning("⚠️ Nincsenek feldolgozható tételek")
        return

    logger.info(f"📊 Összesen {len(pairs)} páciens feldolgozása kezdődik...")

    success_count = 0
    failed_count = 0

    for idx, (name, items, specimen_date, _doc_number, excel_rows) in enumerate(pairs, start=1):
        logger.info(f"\n{'='*60}")
        logger.info(f"➡️ #{idx}/{len(pairs)} – {name}")
        logger.info(f"{'='*60}")

        # Dátum konverzió stringre
        specimen_date_str = None
        if isinstance(specimen_date, date) and not isinstance(specimen_date, datetime):
            dt = datetime.combine(specimen_date, datetime.now().time())
            specimen_date_str = dt.strftime("%Y.%m.%d %H:%M")
        elif isinstance(specimen_date, pd.Timestamp):
            specimen_date_str = specimen_date.strftime("%Y.%m.%d %H:%M")
        elif isinstance(specimen_date, datetime):
            specimen_date_str = specimen_date.strftime("%Y.%m.%d %H:%M")
        elif specimen_date is not None:
            specimen_date_str = str(specimen_date)

        try:
            ok_proc = process_patient_items(
                driver,
                name,
                items,
                specimen_date=specimen_date_str,
                excel_rows=excel_rows,
                mark_saved=lambda rows=list(excel_rows), pname=name: _write_k_sync(
                    pname, rows, excel_path, sheet_name
                ),
            )
            if ok_proc:
                success_count += 1
            else:
                failed_count += 1

        except (InvalidSessionIdException, NoSuchWindowException, WebDriverException) as e:
            logger.exception(f"❌ Kritikus WebDriver hiba ({name}): {e}")
            raise
        except Exception as e:
            logger.exception(f"❌ Váratlan hiba ({name}): {e}")
            failed_count += 1
            continue

    logger.info(f"\n{'='*60}")
    logger.info(f"📊 ÖSSZESÍTÉS:")
    logger.info(f"   ✅ Sikeres: {success_count}")
    logger.info(f"   ❌ Sikertelen: {failed_count}")
    logger.info(f"   📋 Összesen: {len(pairs)}")
    logger.info(f"{'='*60}\n")


def run_items_flow_standalone():
    """Belépés + tételek feldolgozása (standalone)."""
    logger.info("🚀 Tétel-feltöltés indul (standalone)...")
    driver = make_driver()

    try:
        lp = LoginPage(driver, LOGIN_URL, timeout=45)
        if not lp.login(USERNAME, PASSWORD):
            logger.error("❌ Login sikertelen (tétel modul).")
            return
        logger.info("✅ Login sikeres (tétel modul).")

        admit_and_check_items_from_excel(driver, EXCEL_PATH, SHEET_NAME)
        logger.info("🏁 Minden páciens feldolgozva.")

    except KeyboardInterrupt:
        logger.info("⚠️ Megszakítva (Ctrl+C).")
    except Exception as e:
        logger.exception(f"❌ Kritikus hiba: {e}")
    finally:
        try:
            driver.quit()
            logger.info("🔚 Böngésző bezárva.")
        except Exception:
            pass


if __name__ == "__main__":
    run_items_flow_standalone()
