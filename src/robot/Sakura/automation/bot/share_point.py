import os
import re
import time
from typing import List, Tuple

from selenium.common.exceptions import (
    ElementClickInterceptedException,
    StaleElementReferenceException,
    TimeoutException,
)
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC

from src.robot.Sakura.automation.bot.base import IBot
from src.robot.Sakura.automation.common import retry_if_exception


class SharePoint(IBot):
    def __init__(self, url: str, username: str, password: str, **kwargs):
        super().__init__(**kwargs)
        self.url = url
        self.authenticated = self._authentication(username, password)
        if not self.authenticated:
            raise ConnectionRefusedError("The username or password is incorrect.")

    def navigate(self, url, wait_for_complete=True):
        super().navigate(url, wait_for_complete)
        try:
            ms_error_header = self.wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "div[id='ms-error-header']"))
            )
            self.logger.error(ms_error_header.text)
            self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "div[id='ms-error'] a"))).click()
            time.sleep(self.retry_interval)
            while self.browser.execute_script("return document.readyState") != "complete":
                continue
            time.sleep(self.retry_interval)
            return self.navigate(url, wait_for_complete)
        except Exception:
            pass

    @retry_if_exception(
        exceptions=(
            StaleElementReferenceException,
            ElementClickInterceptedException,
            TimeoutException,
            Exception,
        ),
        failure_return=False,
    )
    def _authentication(self, username: str, password: str) -> bool:
        self.navigate("https://login.microsoftonline.com/")
        time.sleep(self.retry_interval)
        if self.browser.current_url.startswith("https://m365.cloud.microsoft/?auth="):
            self.logger.info("Authenticated")
            return True
        # -- Username
        time.sleep(self.retry_interval)
        try:
            self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'input[type="email"]'))).send_keys(
                username
            )
            # -- Next
            btn = self.wait.until(EC.presence_of_element_located((By.ID, "idSIButton9")))
            self.wait.until(EC.element_to_be_clickable(btn)).click()
            # -- Password
            self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'input[type="password"]'))).send_keys(password)
            # -- Sign in
            self.wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()
        except TimeoutException:
            alert = self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[role='alert']")))
            self.logger.error(alert.text)
            return False

        try:  # Password Error
            alert = self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[role='alert']")))
            self.logger.error(alert.text)
            return False
        except TimeoutException:
            pass
        # -- Stay Signed In
        self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[id='idSIButton9']"))).click()
        time.sleep(1)
        self.navigate(self.url)
        if self.browser.current_url.find(".sharepoint.com") == -1:
            self.logger.info(" Xác thực thất bại!")
            return False
        self.logger.info("Authenticated")
        return True

    @retry_if_exception(
        exceptions=(
            StaleElementReferenceException,
            ElementClickInterceptedException,
            TimeoutException,
        ),
        failure_return=[
            (None, None, "Download Error"),
        ],
    )
    def download(self, site_url: str, file_pattern: str) -> List[Tuple[str | None, str, str]]:
        self.logger.info(f"Search {file_pattern} - {site_url}")
        result = []
        self.navigate(site_url)
        # Folder
        folder_found = False
        folders: List[str] = file_pattern.split("/")[:-1]
        for step in folders:
            time.sleep(self.retry_interval)
            gridcells = self.wait.until(
                EC.any_of(
                    EC.presence_of_all_elements_located(
                        (By.CSS_SELECTOR, "div[role='gridcell'][data-automationid='field-LinkFilename']")
                    ),
                    EC.presence_of_all_elements_located(
                        (By.CSS_SELECTOR, "div[role='gridcell'][data-automation-key^='displayNameColumn']")
                    ),
                )
            )
            for gridcell in gridcells:
                text = re.sub(r"[\ue000-\uf8ff]|\n|Press C to open file hover card", "", gridcell.text)
                if re.match(step, text):
                    button = gridcell.find_element(By.XPATH, './/button | .//span[@role="button"]')
                    self.wait.until(EC.element_to_be_clickable(button)).click()
                    time.sleep(self.timeout)
                    folder_found = True
                    break
            time.sleep(self.retry_interval)
        if not folder_found:
            raise LookupError(f"Folder Not Found: {file_pattern}")
        # File
        file: str = file_pattern.split("/")[-1]
        gridcells = self.wait.until(
            EC.any_of(
                EC.presence_of_all_elements_located(
                    (By.CSS_SELECTOR, "div[role='gridcell'][data-automationid='field-LinkFilename']")
                ),
                EC.presence_of_all_elements_located(
                    (By.CSS_SELECTOR, "div[role='gridcell'][data-automation-key^='displayNameColumn']")
                ),
            )
        )
        for gridcell in gridcells:
            file_name = re.sub(r"[\ue000-\uf8ff]|\n|Press C to open file hover card", "", gridcell.text)
            if re.match(file, file_name):
                button = gridcell.find_element(By.XPATH, './/button | .//span[@role="button"]')
                self.wait.until(EC.element_to_be_clickable(button))
                time.sleep(self.retry_interval)
                # Copy Link
                link = None
                # Download File
                ActionChains(self.browser).context_click(button).perform()
                download_btn = self.wait.until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "button[data-automationid='downloadCommand']"))
                )
                self.wait.until(EC.element_to_be_clickable(download_btn)).click()
                time.sleep(2.5)
                file_path, status = self.wait_for_download_to_finish()
                self.logger.info(f"Download {file_name}: {file_path if not status else status},")
                result.append((link, os.path.join(self.download_directory, file_path), status))
        return result


__all__ = ["SharePoint"]
