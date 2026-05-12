from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException
from utils.logger import logger


class LoginPage:
    def __init__(self, driver, login_url: str, timeout: int = 30):
        self.driver = driver
        self.login_url = login_url
        self.timeout = timeout

    def login(self, username: str, password: str) -> bool:
        """
        Robusztus login:
        - mezők láthatóságára vár
        - kattintható login gombra vár
        - URL-váltást figyel
        - majd a dashboard markerre vár: div[title='Ellátási idők']
        """

        def attempt() -> bool:
            try:  # ← EZ AZ ELSŐ (ÉS EGYETLEN) TRY
                logger.info(f"Navigálás a login oldalra: {self.login_url}")
                self.driver.get(self.login_url)

                # mezők
                user_field = WebDriverWait(self.driver, self.timeout).until(
                    EC.visibility_of_element_located(
                        (By.CSS_SELECTOR, "input[name='username'], #username")
                    )
                )
                pass_field = WebDriverWait(self.driver, self.timeout).until(
                    EC.visibility_of_element_located(
                        (By.CSS_SELECTOR, "input[type='password'], #password")
                    )
                )

                # kitöltés
                user_field.clear()
                user_field.send_keys(username)
                pass_field.clear()
                pass_field.send_keys(password)

                # login gomb
                login_button = WebDriverWait(self.driver, self.timeout).until(
                    EC.element_to_be_clickable(
                        (By.CSS_SELECTOR, "#kc-login, button[type='submit']")
                    )
                )

                logger.info("Kattintás a login gombra...")
                login_button.click()

                # URL már ne tartalmazza a login-t
                WebDriverWait(self.driver, self.timeout).until(
                    lambda d: "login" not in d.current_url.lower()
                )

                # Dashboard marker → Ellátási idők
                WebDriverWait(self.driver, self.timeout).until(
                    EC.visibility_of_element_located(
                        (By.CSS_SELECTOR, 'div[title="Ellátási idők"]')
                    )
                )

                logger.info("Sikeres login – dashboard betöltött!")
                return True

            # ← ITT VANNAK A HOZZÁ TARTOZÓ EXCEPTEK
            except (TimeoutException, StaleElementReferenceException) as e:
                logger.warning(f"Login kísérlet sikertelen (timeout/stale): {e}")
                return False
            except Exception as e:
                logger.error(f"Váratlan hiba login közben: {e}")
                return False

        # első próbálkozás
        if attempt():
            return True

        # retry
        import time
        logger.info("Login újrapróbálása 2 másodperc múlva...")
        time.sleep(2)
        return attempt()
