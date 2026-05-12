"""
Address handling functions for patient upload automation.
Handles: Irányítószám (Zip Code), Település (Settlement), Cím (Address Line)

HASZNÁLAT:
    from address_handlers import fill_address_data
    
    # A patient upload függvényben:
    try:
        fill_address_data(driver, row)
    except Exception as e:
        logger.warning(f"Address error: {e}")
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


def open_address_section(driver):
    """
    Clicks the address section 'Hozzáadás' button (local only, not global Felvétel).
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
                    logger.info(f"🛡️ Prevented global Felvétel click in open_address_section: {pattern}")
                    return True
        except Exception:
            pass
        return False
    
    try:
        # Find address section by heading or known context
        section = None
        try:
            # Look for "Címek" or "Lakcím" or "Elérhetőségek" heading
            for heading_text in ["Címek", "Lakcím", "Elérhetőségek", "Address"]:
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
        try:
            btn = root.find_element(By.CSS_SELECTOR, '[data-automation-id="__addNewItemCompactButton"]')
            if btn.is_displayed() and not _is_forbidden_element(btn):
                try:
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
                    btn.click()
                    time.sleep(0.3)
                    logger.info("✅ Address section opened via __addNewItemCompactButton")
                    return True
                except Exception:
                    driver.execute_script("arguments[0].click();", btn)
                    time.sleep(0.3)
                    logger.info("✅ Address section opened via JS click")
                    return True
        except Exception:
            pass
        
        # Fallback: find any "Hozzáadás" button near address-related elements
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
                    logger.info("✅ Address section opened via Hozzáadás button")
                    return True
                except Exception:
                    continue
        except Exception:
            pass
        
        return False
    except Exception as e:
        logger.warning(f"⚠️ open_address_section failed: {e}")
        return False


def fill_zip_code(driver, zip_code: str, timeout=15):
    """
    Fill the zip code React-Select input field.
    Returns True if successful, False otherwise.
    """
    if not zip_code or pd.isna(zip_code):
        return False
    
    zip_str = str(zip_code).strip()
    
    try:
        # Find the zip code input by data-automation-id
        zip_input = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, '[data-automation-id="____settlementAndZipCodeSelector.zipCode"]'))
        )
        
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", zip_input)
        
        # Click to focus
        try:
            zip_input.click()
        except Exception:
            driver.execute_script("arguments[0].click();", zip_input)
        
        # Clear and type
        try:
            zip_input.send_keys(Keys.CONTROL, "a")
            time.sleep(0.05)
            zip_input.send_keys(Keys.BACKSPACE)
            time.sleep(0.05)
        except Exception:
            pass
        
        zip_input.send_keys(zip_str)
        time.sleep(0.3)  # Wait for settlement dropdown to populate
        
        # Verify
        val = (zip_input.get_attribute("value") or "").strip()
        if val == zip_str:
            logger.info(f"✅ Zip code filled: {zip_str}")
            return True
        else:
            logger.warning(f"⚠️ Zip code mismatch: expected '{zip_str}', got '{val}'")
            return False
            
    except Exception as e:
        logger.warning(f"⚠️ fill_zip_code failed: {e}")
        return False


def select_settlement(driver, settlement_name: str, timeout=15):
    """
    Select settlement from React-Select dropdown after zip code is entered.
    Returns True if successful, False otherwise.
    """
    if not settlement_name or pd.isna(settlement_name):
        return False
    
    settlement_str = str(settlement_name).strip()
    
    try:
        # Find the settlement input
        settlement_input = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, '[data-automation-id="____settlementAndZipCodeSelector.settlement"]'))
        )
        
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", settlement_input)
        
        # Click to open dropdown
        try:
            settlement_input.click()
        except Exception:
            driver.execute_script("arguments[0].click();", settlement_input)
        
        time.sleep(0.2)
        
        # Type the settlement name (it should filter/autocomplete)
        try:
            settlement_input.send_keys(Keys.CONTROL, "a")
            time.sleep(0.05)
            settlement_input.send_keys(Keys.BACKSPACE)
            time.sleep(0.05)
        except Exception:
            pass
        
        settlement_input.send_keys(settlement_str)
        time.sleep(0.3)
        
        # Wait for options to appear and select the first match
        try:
            # Look for option containing the settlement name
            option_xpath = f"//*[@role='option' and contains(translate(normalize-space(),'ABCDEFGHIJKLMNOPQRSTUVWXYZÁÉÍÓÖŐÚÜŰ','abcdefghijklmnopqrstuvwxyzáéíóöőúüű'), '{settlement_str.lower()}')]"
            option = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, option_xpath))
            )
            
            try:
                option.click()
            except Exception:
                driver.execute_script("arguments[0].click();", option)
            
            time.sleep(0.2)
            logger.info(f"✅ Settlement selected: {settlement_str}")
            return True
        except TimeoutException:
            # If no dropdown appears, just press ENTER to confirm
            try:
                settlement_input.send_keys(Keys.ENTER)
                time.sleep(0.2)
                logger.info(f"✅ Settlement confirmed with ENTER: {settlement_str}")
                return True
            except Exception:
                pass
        
        return False
        
    except Exception as e:
        logger.warning(f"⚠️ select_settlement failed: {e}")
        return False


def fill_address_line(driver, address: str, timeout=15):
    """
    Fill the AddressLine input field.
    Returns True if successful, False otherwise.
    """
    if not address or pd.isna(address):
        return False
    
    address_str = str(address).strip()
    
    try:
        # Find AddressLine input by ID
        address_input = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.ID, "AddressLine"))
        )
        
        # Scroll and click
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", address_input)
        
        try:
            address_input.click()
        except Exception:
            driver.execute_script("arguments[0].click();", address_input)
        
        # Clear
        try:
            address_input.send_keys(Keys.CONTROL, "a")
            time.sleep(0.05)
            address_input.send_keys(Keys.BACKSPACE)
            time.sleep(0.05)
        except Exception:
            try:
                address_input.clear()
            except Exception:
                pass
        
        # Type
        address_input.send_keys(address_str)
        time.sleep(0.1)
        
        # Verify
        val = (address_input.get_attribute("value") or "").strip()
        if val == address_str:
            logger.info(f"✅ Address line filled: {address_str}")
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
                """, address_input, address_str)
                time.sleep(0.1)
                val = (address_input.get_attribute("value") or "").strip()
            except Exception:
                pass
            
            if val == address_str:
                logger.info(f"✅ Address line filled (JS fallback): {address_str}")
                return True
            else:
                logger.warning(f"⚠️ Address line mismatch: expected '{address_str}', got '{val}'")
                return False
            
    except Exception as e:
        logger.warning(f"⚠️ fill_address_line failed: {e}")
        return False


def fill_address_data(driver, row):
    """
    Main function to fill all address data from Excel row.
    Expects columns with various possible names:
        - Irányítószám: "Paciens/Iranyitoszam", "I", "Irányítószám", "Iranyitoszam", "Zip", "ZipCode"
        - Település: "Paciens/Telepules", "J", "Település", "Telepules", "Settlement", "City"
        - Cím: "Paciens/Cim", "K", "Cím", "Cim", "Address", "AddressLine"
    Returns True if at least the address section was opened successfully.
    """

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

    # Try multiple column name variations for each field
    zip_code = get_value("Paciens/Iranyitoszam", "I", "Irányítószám", "Iranyitoszam", "Zip", "ZipCode")
    settlement = get_value("Paciens/Telepules", "J", "Település", "Telepules", "Settlement", "City")
    address_line = get_value("Paciens/Cim", "K", "Cím", "Cim", "Address", "AddressLine")

    logger.info(f"🔍 Address data from Excel: zip={zip_code}, settlement={settlement}, address={address_line}")

    if not zip_code:
        logger.info("ℹ️ No zip code provided, skipping address section")
        return False
    
    # Open address section
    if not open_address_section(driver):
        logger.warning("⚠️ Address section could not be opened")
        return False
    
    logger.info("✅ Address section opened")
    
    # Fill zip code
    if not fill_zip_code(driver, str(zip_code)):
        logger.warning(f"⚠️ Zip code fill failed: {zip_code}")
        return False
    
    logger.info(f"✅ Zip code filled: {zip_code}")
    
    # Select settlement (should auto-populate after zip)
    if settlement and not pd.isna(settlement):
        select_settlement(driver, str(settlement))
        logger.info(f"✅ Settlement handled: {settlement}")
    
    # Fill address line
    if address_line and not pd.isna(address_line):
        if fill_address_line(driver, str(address_line)):
            logger.info(f"✅ Address line filled: {address_line}")
        else:
            logger.warning(f"⚠️ Address line fill failed: {address_line}")
    
    return True