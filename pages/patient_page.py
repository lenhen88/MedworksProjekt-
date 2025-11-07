# pages/patient_page.py
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

from selenium.webdriver.support import expected_conditions as EC

class PatientPage:
    def __init__(self, driver, timeout: int = 20):
        self.driver = driver
        self.timeout = timeout

    # Lokátorok a küldött TXT alapján
    NEW_PATIENT_BTN = (By.CSS_SELECTOR, '[data-automation-id="PatientRegister_CreateNewPatient"]')

    ADD_DOCUMENT_BTN = (By.CSS_SELECTOR, '[data-automation-id="__addNewItemCompactButton"]')
    DOC_TYPE_DROPDOWN = (By.CSS_SELECTOR, '[data-automation-id="chevronDown"]')
    DOC_TAJ_OPTION = (By.CSS_SELECTOR, '[data-automation-id="__option__11"]')  # TAJ szám
    DOC_NUMBER_INPUT = (By.ID, "DocumentNumber")

    LASTNAME_INPUT = (By.ID, "LastName")
    FIRSTNAME_INPUT = (By.ID, "FirstName")
    BIRTHDATE_INPUT = (By.ID, "BirthDate")

    GENDER_MALE = (By.ID, "SexId_Male")
    GENDER_FEMALE = (By.ID, "SexId_Female")

    EMAIL_INPUT = (By.ID, "EmailAddress")

    SAVE_BTN = (By.CSS_SELECTOR, '[data-automation-id="__save_save"]')

    def wait_ready(self) -> bool:
        try:
            WebDriverWait(self.driver, self.timeout).until(
                EC.presence_of_element_located(self.NEW_PATIENT_BTN)
            )
            return True
        except TimeoutException:
            return False

    def click_new_patient(self):
        WebDriverWait(self.driver, self.timeout).until(
            EC.element_to_be_clickable(self.NEW_PATIENT_BTN)
        ).click()

    def add_document(self):
        WebDriverWait(self.driver, self.timeout).until(
            EC.element_to_be_clickable(self.ADD_DOCUMENT_BTN)
        ).click()

    def open_doc_type_dropdown(self):
        WebDriverWait(self.driver, self.timeout).until(
            EC.element_to_be_clickable(self.DOC_TYPE_DROPDOWN)
        ).click()

    def choose_taj_document_type(self):
        WebDriverWait(self.driver, self.timeout).until(
            EC.element_to_be_clickable(self.DOC_TAJ_OPTION)
        ).click()

    def fill_document_number(self, value: str):
        el = WebDriverWait(self.driver, self.timeout).until(
            EC.visibility_of_element_located(self.DOC_NUMBER_INPUT)
        )
        el.clear()
        el.send_keys(value or "")

    def fill_last_name(self, value: str):
        el = WebDriverWait(self.driver, self.timeout).until(
            EC.visibility_of_element_located(self.LASTNAME_INPUT)
        )
        el.clear()
        el.send_keys(value or "")

    def fill_first_name(self, value: str):
        el = WebDriverWait(self.driver, self.timeout).until(
            EC.visibility_of_element_located(self.FIRSTNAME_INPUT)
        )
        el.clear()
        el.send_keys(value or "")

    def fill_birthdate(self, value: str):
        el = WebDriverWait(self.driver, self.timeout).until(
            EC.visibility_of_element_located(self.BIRTHDATE_INPUT)
        )
        el.clear()
        el.send_keys(value or "")

    def select_gender(self, value: str):
        """
        value: "Férfi" / "Nő" / "F" / "M" stb. — megpróbálunk okosan dönteni
        """
        v = (value or "").strip().lower()
        if v in ("ferfi", "f", "male", "m"):
            target = self.GENDER_MALE
        elif v in ("no", "nő", "female"):
            target = self.GENDER_FEMALE
        else:
            return  # nem jelölünk semmit, ha üres vagy ismeretlen
        WebDriverWait(self.driver, self.timeout).until(
            EC.element_to_be_clickable(target)
        ).click()

    def fill_email(self, value: str):
        el = WebDriverWait(self.driver, self.timeout).until(
            EC.visibility_of_element_located(self.EMAIL_INPUT)
        )
        el.clear()
        el.send_keys(value or "")

    def click_save(self):
        WebDriverWait(self.driver, self.timeout).until(
            EC.element_to_be_clickable(self.SAVE_BTN)
        ).click()

def open_new_patient_form(driver):
    """Új páciens űrlap megnyitása és a betöltődés megvárása."""

    try:
        # Kattintás az "Új páciens" gombra
        el = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-automation-id="PatientRegister_CreateNewPatient"]'))
        )
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", el)
        el.click()
        
        # Várjuk, amíg a form ténylegesen betöltődik
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.ID, "LastName"))
        )
        print("🩺 Új páciens űrlap betöltve!")

    except Exception as e:
        print(f"⚠️ Hiba történt az űrlap megnyitásakor: {e}")
