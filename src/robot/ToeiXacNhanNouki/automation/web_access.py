import logging
import os
import time

import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import ElementClickInterceptedException, TimeoutException
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait


class WebAccess:
    def __init__(
        self,
        username: str,
        password: str,
        timeout: int = 10,
        headless: bool = False,
        logger: logging.Logger = logging.getLogger("WebAccess"),
    ):
        options = webdriver.ChromeOptions()
        options.add_argument("--disable-notifications")
        if headless:
            options.add_argument("--headless=new")
        # Disable log
        options.add_argument("--disable-logging")
        options.add_argument("--log-level=3")  #
        options.add_argument("--silent")
        options.add_experimental_option("excludeSwitches", ["enable-logging"])
        # Attribute
        self.logger = logger
        self.browser = webdriver.Chrome(options=options)
        self.browser.maximize_window()
        self.wait = WebDriverWait(self.browser, timeout)
        self.username = username
        self.password = password
        # Trạng thái đăng nhập
        self.root_window = self.browser.window_handles[0]
        self.authenticated = self.__authentication(username, password)

    def __del__(self):
        if hasattr(self, "browser") and isinstance(self.browser, WebDriver):
            self.browser.quit()

    def __authentication(self, username: str, password: str) -> bool:
        self.browser.get("https://webaccess.nsk-cad.com")
        self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='text']"))).send_keys(username)
        self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='password']"))).send_keys(password)
        self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "button[class='btn login']"))).click()
        try:
            error_box = self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[id='f-error-box']")))
            data = error_box.find_element(By.CSS_SELECTOR, "div[class='data']")
            self.logger.info(f"❌ Xác thực thất bại!: {data.text}")
            return False
        except TimeoutException:
            self.logger.info("✅ Xác thực thành công!")
            return True

    def __switch_tab(self, tab: str) -> bool:
        try:
            a = self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, f"a[title='{tab}']")))
            href = a.get_attribute("href")
            self.browser.get(href)
            return True
        except ElementClickInterceptedException:
            return self.__switch_tab(tab=tab)
        except Exception as e:
            self.logger.error(e)
            return False

    def open_new_tab(self) -> str:
        before_windows = self.browser.window_handles
        self.browser.execute_script("window.open('');")
        after_windows = self.browser.window_handles
        return list(set(after_windows) - set(before_windows))[0]

    def navigate(self, url, wait_for_complete: bool = True):
        time.sleep(1)
        self.browser.execute_script("window.stop()")
        time.sleep(1)
        self.browser.get(url)
        time.sleep(1)
        if wait_for_complete:
            while self.browser.execute_script("return document.readyState") != "complete":
                time.sleep(1)
        time.sleep(1)

    def wait_for_download_to_finish(self) -> tuple[str, str]:
        window_id = self.open_new_tab()
        self.browser.switch_to.window(window_id)
        self.navigate("chrome://downloads")
        download_items = self.browser.execute_script(
            """
            return document.
                querySelector("downloads-manager").shadowRoot
                .querySelector("#mainContainer #downloadsList #list")
                .querySelectorAll("downloads-item")
        """
        )
        item_id = download_items[0].get_attribute("id")
        while self.browser.execute_script(
            f"""
            return document
                .querySelector("downloads-manager").shadowRoot
                .querySelector("#downloadsList #list")
                .querySelector("#{item_id}").shadowRoot
                .querySelector("#content #details #progress")
            """
        ):  # Progess
            time.sleep(self.retry_interval)
        name = self.browser.execute_script(
            f"""
            return document
                .querySelector("downloads-manager").shadowRoot
                .querySelector("#downloadsList")
                .querySelector("#list")
                .querySelector("#{item_id}").shadowRoot
                .querySelector("#content")
                .querySelector("#details")
                .querySelector("#title-area")
                .querySelector("#name")
                .getAttribute("title")
            """
        )
        tag = self.browser.execute_script(
            f"""
            return document
                .querySelector("downloads-manager").shadowRoot
                .querySelector("#downloadsList")
                .querySelector("#list")
                .querySelector("#{item_id}").shadowRoot
                .querySelector("#content")
                .querySelector("#details")
                .querySelector("#title-area")
                .querySelector("#tag")
                .textContent.trim();
            """
        )
        self.browser.close()
        self.browser.switch_to.window(self.root_window)
        return name, tag

    def get_information(self, construction_id: str, fields: list[str] = None) -> pd.DataFrame:
        try:
            self.__switch_tab("受注一覧")
            self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "button[type='reset']"))).click()
            date_picker = self.wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "input[name='search_fix_deliver_date_from']"))
            )
            date_picker.clear()
            date_picker.send_keys(Keys.ESCAPE)
            self.wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "input[name='search_construction_no']"))
            ).send_keys(construction_id)

            self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "button[type='submit']"))).click()
            try:
                time.sleep(5)
                self.wait.until(EC.presence_of_element_located((By.XPATH, "//td[text()='検索結果はありません']")))
                self.logger.warning(f"❌ Construction:{construction_id} không có dữ liệu")
                return pd.DataFrame(columns=fields)
            except TimeoutException:
                time.sleep(1)
                self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[id='checkAll']"))).click()

                if not fields:
                    self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[id='checkAll']"))).click()
                else:
                    for field in fields:
                        xpath = f"//label[text()='{field}']//input[@type='checkbox']"
                        self.wait.until(EC.presence_of_element_located((By.XPATH, xpath))).click()

                data_tables = self.wait.until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "div[class='dataTables_scroll']"))
                )
                # Columns
                dataTables_scrollHead = data_tables.find_element(By.CSS_SELECTOR, "div[class='dataTables_scrollHead']")
                spans = dataTables_scrollHead.find_elements(By.TAG_NAME, "span")
                columns = [span.text for span in spans]

                df = pd.DataFrame(columns=columns)
                # Row
                dataTables_scrollBody = data_tables.find_element(By.CSS_SELECTOR, "div[class='dataTables_scrollBody']")
                dataTables_scrollBody_tbody = dataTables_scrollBody.find_element(By.TAG_NAME, "tbody")
                dataTables_scrollBody_tbody_trs = dataTables_scrollBody_tbody.find_elements(By.TAG_NAME, "tr")
                for tr in dataTables_scrollBody_tbody_trs:
                    tds = tr.find_elements(By.TAG_NAME, "td")
                    row = [td.text for td in tds][1:]
                    df.loc[len(df)] = row
                self.logger.info(f"✅ Lấy dữ liệu construction:{construction_id} thành công")
                return df
        except ElementClickInterceptedException:
            return self.get_information(
                construction_id=construction_id,
                fields=fields,
            )
        except TimeoutException:
            self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a[class='button fa fa-download']"))).click()
            time.sleep(5)
            name, _ = self.wait_for_download_to_finish()
            name = os.path.join(os.path.join(os.path.expanduser("~"), "downloads"), name)
            df = pd.read_csv(name, encoding="cp932")
            os.remove(name)
            if fields:
                df = df[fields]
            return df
        except Exception as e:
            self.logger.error(e)
            return pd.DataFrame(columns=fields)


__all__ = [WebAccess]
