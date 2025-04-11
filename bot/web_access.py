import os
import time
import logging
import pandas as pd
from pathlib import Path
from datetime import date
from logging import Logger
from selenium import webdriver
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.common.exceptions import ElementClickInterceptedException
from selenium.webdriver.common.keys import Keys


class WebAccess:
    def __init__(
        self,
        username: str,
        password: str,
        timeout: int = 10,
        headless: bool = False,
        download_directory: str = os.path.dirname(os.path.abspath(__file__)),
        logger: Logger = None,
        logger_name: str = __name__,
    ):
        os.makedirs(download_directory, exist_ok=True)
        options = webdriver.ChromeOptions()
        options.add_argument("--disablenotifications")  # Tắt thông báo
        options.add_experimental_option(
            "prefs",
            {
                "download.default_directory": download_directory,
                "download.prompt_for_download": False,
                "safebrowsing.enabled": True,
            },
        )

        if headless:
            options.add_argument("--headless=new")
        # Disable log
        options.add_argument("--disable-logging")
        options.add_argument("--log-level=3")  #
        options.add_argument("--silent")
        options.add_experimental_option("excludeSwitches", ["enable-logging"])
        # Attribute
        self.download_directory = download_directory
        if logger:
            self.logger = logger
        else:
            self.logger = logging.getLogger(logger_name)
        self.browser = webdriver.Chrome(options=options)
        self.browser.maximize_window()
        self.timeout = timeout
        self.wait = WebDriverWait(self.browser, timeout)
        self.username = username
        self.password = password
        # Trạng thái đăng nhập
        self.authenticated = self.__authentication(username, password)

    def __del__(self):
        if hasattr(self, "browser") and isinstance(self.browser, WebDriver):
            self.browser.quit()

    def __authentication(self, username: str, password: str) -> bool:
        try:
            self.browser.get("https://webaccess.nsk-cad.com")
            self.wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='text']"))
            ).send_keys(username)
            self.wait.until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, "input[type='password']")
                )
            ).send_keys(password)
            self.wait.until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, "button[class='btn login']")
                )
            ).click()
            self.logger.info("Xác thực thành công!")
            return True
        except Exception as e:
            self.logger.error(f"Xác thực thất bại! {e}.")
            return False

    def __switch_tab(self, tab: str) -> bool:
        try:
            a = self.wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, f"a[title='{tab}']"))
            )
            href = a.get_attribute("href")
            self.browser.get(href)
            return True
        except Exception as e:
            self.logger.error(e)
            return False

    def __get_newest_csv(self) -> Path | None:
        download_dir = Path(self.download_directory)
        csv_files = list(download_dir.glob("*.csv"))
        if csv_files:
            latest_file = max(csv_files, key=lambda f: f.stat().st_birthtime)
            return latest_file
        else:
            return None

    def get_information(
        self,
        ビルダー名: str = None,
        図面: list[str] = None,
        確定納品日: list[str] = [date.today().strftime("%Y/%m/%d"), ""],
        リセット: bool = True,
        FIELDS: list[str] = None,
        output_file: str = None,
    ) -> pd.DataFrame | None:
        try:
            self.__switch_tab("受注一覧")
            # Clear
            if リセット:
                self.wait.until(
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR, "button[type='reset']")
                    )
                )
                self.wait.until(
                    EC.element_to_be_clickable(
                        (By.CSS_SELECTOR, "button[type='reset']")
                    )
                ).click()
            # Filter
            if ビルダー名:  # Build name
                self.wait.until(
                    EC.presence_of_element_located(
                        (By.ID, "select2-search_builder_cd-container")
                    )
                )
                self.wait.until(
                    EC.element_to_be_clickable(
                        (By.ID, "select2-search_builder_cd-container")
                    )
                ).click()
                time.sleep(1)
                self.wait.until(
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR, 'input[class="select2-search__field"]')
                    )
                ).send_keys(ビルダー名)
                self.wait.until(
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR, 'input[class="select2-search__field"]')
                    )
                ).send_keys(Keys.ENTER)
                time.sleep(1)
            if 確定納品日:  # Delivery date
                self.wait.until(
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR, 'input[name="search_fix_deliver_date_from"]')
                    )
                ).clear()
                self.wait.until(
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR, 'input[name="search_fix_deliver_date_from"]')
                    )
                ).send_keys(確定納品日[0])
                self.wait.until(
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR, 'input[name="search_fix_deliver_date_to"]')
                    )
                ).clear()
                self.wait.until(
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR, 'input[name="search_fix_deliver_date_to"]')
                    )
                ).send_keys(確定納品日[1])
                self.wait.until(
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR, 'input[name="search_fix_deliver_date_to"]')
                    )
                ).send_keys(Keys.ESCAPE)
            if 図面:  # Drawing
                self.wait.until(
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR, "button[id='search_drawing_type_ms']")
                    )
                )
                self.wait.until(
                    EC.element_to_be_clickable(
                        (By.CSS_SELECTOR, "button[id='search_drawing_type_ms']")
                    )
                ).click()
                for e in 図面:
                    xpath = f"//span[text()='{e}']"
                    self.wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
                    self.wait.until(
                        EC.element_to_be_clickable((By.XPATH, xpath))
                    ).click()
                self.wait.until(
                    EC.element_to_be_clickable(
                        (By.CSS_SELECTOR, "button[id='search_drawing_type_ms']")
                    )
                ).send_keys(Keys.ESCAPE)
            # Search
            time.sleep(1)
            self.wait.until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, "button[type='submit']")
                )
            )
            self.wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "button[type='submit']"))
            ).click()
            time.sleep(2)
            # Download File
            self.wait.until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, "a[class='button fa fa-download']")
                )
            )
            self.wait.until(
                EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, "a[class='button fa fa-download']")
                )
            ).click()
            time.sleep(5)
            filePath = self.__get_newest_csv()
            df = pd.read_csv(filePath, encoding="CP932")
            df = df[FIELDS] if FIELDS else df
            filePath.unlink()
            df.to_csv(output_file, index=False, encoding="CP932")
            self.logger.info(f"Đã lấy thông tin {ビルダー名} thành công!")
            if output_file:
                self.logger.info(f"Đã lưu file {output_file}")
            return df
        except ElementClickInterceptedException as e:
            return self.get_information(
                ビルダー名=ビルダー名,
                図面=図面,
                確定納品日=確定納品日,
                リセット=リセット,
                FIELDS=FIELDS,
                output_file=output_file,
            )
        except Exception as e:
            self.logger.error(e)
            return None


__all__ = [WebAccess]
