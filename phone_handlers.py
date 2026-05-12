"""
Phone number handling functions for patient upload automation.
Handles: Mobiltelefon (Mobile Phone)

HASZNÁLAT:
    from phone_handlers import fill_phone_data
    
    # A patient upload függvényben:
    try:
        fill_phone_data(driver, row)
    except Exception as e:
        logger.warning(f"Phone error: {e}")
"""

import time
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

try:
    from utils.logger import logger
except ImportError:
    # Fallback if logger not available
    import logging
    logger = logging.getLogger(__name__)
    logging.basicConfig(level=logging.INFO)


def open_phone_section(driver):
    """
    Clicks the phone section 'Hozzáadás' button (local only, not global Felvétel).
    Returns True if successful, False otherwise.
    """
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
                    logger.info(f"🛡️ Prevented global Felvétel click in open_phone_section: {pattern}")
                    return True
        except Exception:
            pass
        return False
    
    try:
        # Find phone section by heading or known context
        section = None
        try:
            # Look for "Mobiltelefon" or "Telefonszám" or "Elérhetőségek" heading
            for heading_text in ["Mobiltelefon", "Telefonszám", "Elérhetőségek", "Phone"]:
                try:
                    label = driver.find_element(By.XPATH, f"//*[contains(normalize-space(),'{heading_text}')]")
                    section = label.find_element(By.XPATH, "ancestor::*[contains(@class,'section') or contains(@class,'group') or @role='region'][1]")
                    break
                except Exception:
                    continue
        except Exception:
            pass
        
        root = section if section is not None else driver
        
        # Find the add button with data-automation-id="__addNewItemCompactButton"
        # Look for the SVG with circleAdd specifically for phone section
        try:
            # Find button by the circleAdd SVG icon
            buttons = root.find_elements(By.CSS_SELECTOR, '[data-automation-id="__addNewItemCompactButton"]')
            for btn in buttons:
                try:
                    if not btn.is_displayed():
                        continue
                    if _is_forbidden_element(btn):
                        continue
                    
                    # Check if this button is near phone-related elements
                    try:
                        # Check parent context for phone-related text
                        parent_text = btn.find_element(By.XPATH, "ancestor::*[1]").text.lower()
                        if "telefon" in parent_text or "phone" in parent_text or "mobiltelefon" in parent_text:
                            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
                            try:
                                btn.click()
                            except Exception:
                                driver.execute_script("arguments[0].click();", btn)
                            time.sleep(0.3)
                            logger.info("✅ Phone section opened via __addNewItemCompactButton")
                            return True
                    except Exception:
                        pass
                except Exception:
                    continue
        except Exception:
            pass
        
        # Fallback: find any "Hozzáadás" button near phone-related elements
        try:
            buttons = root.find_elements(By.XPATH, ".//button[contains(normalize-space(),'Hozzáadás')]")
            for btn in buttons:
                try:
                    if not btn.is_displayed():
                        continue
                    if _is_forbidden_element(btn):
                        continue
                    
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
                    try:
                        btn.click()
                    except Exception:
                        driver.execute_script("arguments[0].click();", btn)
                    time.sleep(0.3)
                    logger.info("✅ Phone section opened via Hozzáadás button")
                    return True
                except Exception:
                    continue
        except Exception:
            pass
        
        return False
    except Exception as e:
        logger.warning(f"⚠️ open_phone_section failed: {e}")
        return False


def fill_phone_number(driver, phone_number: str, timeout=15):
    """
    Fill the phone number input field.
    Returns True if successful, False otherwise.
    """
    if not phone_number or pd.isna(phone_number):
        return False
    
    phone_str = str(phone_number).strip()
    
    try:
        # Find the phone number input by ID and data-automation-id
        phone_input = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, '#PhoneNumber, [data-automation-id="PhoneNumber"]'))
        )
        
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", phone_input)
        
        # Click to focus
        try:
            phone_input.click()
        except Exception:
            driver.execute_script("arguments[0].click();", phone_input)
        
        # Clear and type
        try:
            phone_input.send_keys(Keys.CONTROL, "a")
            time.sleep(0.05)
            phone_input.send_keys(Keys.BACKSPACE)
            time.sleep(0.05)
        except Exception:
            try:
                phone_input.clear()
            except Exception:
                pass
        
        phone_input.send_keys(phone_str)
        time.sleep(0.1)
        
        # Verify
        val = (phone_input.get_attribute("value") or "").strip()
        if val == phone_str:
            logger.info(f"✅ Phone number filled: {phone_str}")
            return True
        else:
            # Try JS fallback
            try:
                driver.execute_script("""
                    const el = arguments[0], val = arguments[1];
                    const desc = Object.getOwnPropertyDescriptor(Object.getPrototypeOf(el), 'value') || 
                                 Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value');
                    if (desc && desc.set) {
                        desc.set.call(el, val);
                    } else {
                        el.value = val;
                    }
                    el.dispatchEvent(new Event('input', {bubbles: true}));
                    el.dispatchEvent(new Event('change', {bubbles: true}));
                """, phone_input, phone_str)
                time.sleep(0.1)
                val = (phone_input.get_attribute("value") or "").strip()
            except Exception:
                pass
            
            if val == phone_str:
                logger.info(f"✅ Phone number filled (JS fallback): {phone_str}")
                return True
            else:
                logger.warning(f"⚠️ Phone number mismatch: expected '{phone_str}', got '{val}'")
                return False
            
    except Exception as e:
        logger.warning(f"⚠️ fill_phone_number failed: {e}")
        return False


def fill_phone_data(driver, row):
    """
    Main function to fill phone number data from Excel row.
    Expects columns with various possible names:
        - Mobiltelefon: "Paciens/Mobiltelefon", "H", "Mobiltelefon", "Phone", "PhoneNumber"
    
    Args:
        driver: Selenium WebDriver instance
        row: pandas Series/dict with Excel row data
    
    Returns:
        bool: True if phone section opened and filled, False if no phone or failed
    """
    # Helper function to get value from multiple possible column names
    def get_value(*column_names):
        """Try multiple column names and return first non-empty value."""
        for col in column_names:
            try:
                val = row.get(col, "") if hasattr(row, 'get') else row[col]
                if val and not pd.isna(val) and str(val).strip():
                    return str(val).strip()
            except (KeyError, TypeError, AttributeError):
                continue
        return ""
    
    # Try multiple column name variations for phone
    phone_number = get_value(
        "Paciens/Mobiltelefon", 
        "H", 
        "Mobiltelefon", 
        "Telefonszám", 
        "Telefonszam",
        "Phone", 
        "PhoneNumber",
        "Mobile"
    )
    
    # + jel hozzáadása ha nincs még
    if phone_number and not phone_number.startswith("+"):
        # Ha 06-tal kezdődik, cseréljük +36-ra
        if phone_number.startswith("06"):
            phone_number = "+36" + phone_number[2:]
        # Ha 36-tal kezdődik (ország előhívó de + nélkül)
        elif phone_number.startswith("36"):
            phone_number = "+" + phone_number
        # Egyéb esetben csak + jelet teszünk eléje
        else:
            phone_number = "+" + phone_number
    
    logger.info(f"📞 Telefonszám normalizálva: {phone_number}")
    
    # Debug logging
    logger.info(f"🔍 Phone data from Excel: phone={phone_number}")
    
    if not phone_number:
        logger.info("ℹ️ No phone number provided, skipping phone section")
        return False
    
    # Open phone section
    if not open_phone_section(driver):
        logger.warning("⚠️ Phone section could not be opened")
        return False
    
    logger.info("✅ Phone section opened")
    
    # Fill phone number
    if not fill_phone_number(driver, phone_number):
        logger.warning(f"⚠️ Phone number fill failed: {phone_number}")
        return False
    
    logger.info(f"✅ Phone number filled: {phone_number}")
    
    return True