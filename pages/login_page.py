# pages/login_page.py
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

class LoginPage:
    def __init__(self, driver, login_url: str, timeout: int = 30):
        self.driver = driver
        self.login_url = login_url
        self.timeout = timeout

    # --------- segédek ---------
    def _ready(self):
        """Megvárja, hogy a DOM teljesen betöltődjön."""
        WebDriverWait(self.driver, self.timeout).until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )

    def _wait_url_not_contains(self, fragment: str, t: int = None):
        t = t or self.timeout
        WebDriverWait(self.driver, t).until(
            lambda d: fragment not in d.current_url
        )

    def _wait_any_present(self, locators, t: int = None):
        """Bármelyik (By, selector) megjelenése elegendő."""
        t = t or self.timeout
        def any_present(d):
            for by, sel in locators:
                try:
                    if d.find_elements(by, sel):
                        return True
                except Exception:
                    pass
            return False
        WebDriverWait(self.driver, t).until(any_present)

    # --------- lépések ---------
    def open(self):
        self.driver.get(self.login_url)
        # login oldal első field-je (username) megjelenik
        WebDriverWait(self.driver, self.timeout).until(
            EC.presence_of_element_located((By.ID, "username"))
        )

    def fill_username(self, username: str):
        el = WebDriverWait(self.driver, self.timeout).until(
            EC.presence_of_element_located((By.ID, "username"))
        )
        el.clear()
        el.send_keys(username)

    def fill_password(self, password: str):
        el = WebDriverWait(self.driver, self.timeout).until(
            EC.presence_of_element_located((By.ID, "password"))
        )
        el.clear()
        el.send_keys(password)

    def submit(self):
        WebDriverWait(self.driver, self.timeout).until(
            EC.element_to_be_clickable((By.ID, "kc-login"))
        ).click()

    def login(self, username: str, password: str) -> bool:
        self.open()
        self.fill_username(username)
        self.fill_password(password)
        self.submit()

        try:
            # 1) elhagyjuk az auth URL-t
            self._wait_url_not_contains("protocol/openid-connect/auth", t=max(20, self.timeout))
        except TimeoutException:
            # nem mentünk tovább a login oldalról
            return False

        # 2) megvárjuk, míg a céloldal betölt
        try:
            self._ready()
        except TimeoutException:
            pass  # nem gond, megyünk tovább a vizuális elemekre

        # 3) bármelyik post-login elem megjelenése sikernek számít
        post_login_targets = [
            (By.CSS_SELECTOR, '[data-automation-id="master_PageBoxHeaderTitle"]'),
            (By.CSS_SELECTOR, '[data-automation-id="PatientRegister_CreateNewPatient"]'),
        ]
        try:
            self._wait_any_present(post_login_targets, t=max(25, self.timeout))
            return True
        except TimeoutException:
            return False
