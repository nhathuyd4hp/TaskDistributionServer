import os
import re
import time
from urllib.parse import urljoin

import pandas as pd
from playwright._impl._errors import TimeoutError
from playwright.sync_api import sync_playwright

class WebAccess:
    def __init__(
        self,
        domain: str,
        username: str,
        password: str,
        context=None,
        browser=None,
        playwright=None,
        logger=None,
        headless: bool = False,
        timeout: float = 5000,
    ):
        self._external_context = context is not None
        self._external_browser = browser is not None
        self._external_pw = playwright is not None

        if playwright:
            self._pw = playwright
        else:
            self._pw = sync_playwright().start()

        if browser:
            self.browser = browser
        else:
            self.browser = self._pw.chromium.launch(
                headless=headless,
                timeout=timeout,
                args=["--start-maximized"],
            )

        if context:
            self.context = context
        else:
            self.context = self.browser.new_context(
                no_viewport=True,
            )

        self.page = self.context.new_page()
        self.domain = domain
        self.username = username
        self.password = password
        self.timeout = timeout
        self.logger = logger
        if not self.__authentication():
            raise PermissionError("Authentication failed")

    def __authentication(self) -> bool:
        try:
            self.page.goto(self.domain, wait_until="domcontentloaded")
            self.page.wait_for_selector("input[type='text']").fill(self.username)
            self.page.wait_for_selector("input[type='password']").fill(self.password)
            self.page.wait_for_selector("button[type='submit'][class='btn login']").click()
            try:
                account: str = self.page.locator("div[id='f-menus'] > div[class='name']").text_content().split("：")[-1]
                self.logger.info(f"Logged in as: {account}")
                return True
            except TimeoutError:
                return False
        except TimeoutError:
            self.logger.error("RETRY __authentication")
            return self.__authentication()

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        if not self._external_context:
            self.context.close()
        if not self._external_browser:
            self.browser.close()
        if not self._external_pw:
            self._pw.stop()

    def download_data(
        self,
        building: str,
    ) -> pd.DataFrame:
        self.logger.info(f"Download data: ビルダー名部分一致 = {building} | 図面 = ['作図済', '送付済', 'CBUP済', 'CB送付済', '図面確定']")
        try:
            self.page.bring_to_front()
            with self.page.expect_navigation(wait_until="domcontentloaded"):
                self.page.wait_for_selector("a[class='fa fa-industry']").click()
            self.page.locator("button[type='reset']").click()
            self.page.locator("input[name='search_builder_name_like']").fill(building)
            time.sleep(0.5)
            self.page.locator("button[id='search_drawing_type_ms']").click()
            for drawing in ["作図済", "送付済", "CBUP済", "CB送付済", "図面確定"]:
                self.page.locator(
                    "label[for^='ui-multiselect-5-search_drawing_type-']",
                    has=self.page.locator("span", has_text=drawing),
                ).first.click()
                time.sleep(0.1)
            self.page.locator("button[id='search_drawing_type_ms']").click()

            with self.page.expect_navigation(wait_until="domcontentloaded"):
                self.page.locator("button[class='search fa fa-search']").click(force=True)
            with self.page.expect_download() as download_info:
                self.page.locator("a[class='button fa fa-download']").first.click()
                download = download_info.value
                save_path = os.path.abspath(download.suggested_filename)
                os.makedirs(os.path.dirname(save_path), exist_ok=True)
                download.save_as(save_path)
                try:
                    orders = pd.read_csv(save_path, encoding="cp932")
                except UnicodeDecodeError:
                    orders = pd.read_csv(save_path)
                os.remove(save_path)
                return orders
        except TimeoutError:
            self.logger.error("RETRY download_data")
            return self.download_data(building=building)

    def update(self, case: str, area: str) -> bool:
        self.logger.info(f"Update {case} [延床平米 -> {area}]")
        try:
            self.page.bring_to_front()
            with self.page.expect_navigation(wait_until="domcontentloaded"):
                self.page.wait_for_selector("a[class='fa fa-industry']").click()
            self.page.locator("button[type='reset']").click()
            self.page.locator("input[name='search_project_cd']").fill(case)
            with self.page.expect_navigation(wait_until="domcontentloaded"):
                self.page.locator("button[class='search fa fa-search']").click(force=True)
            time.sleep(0.5)
            hrefs: list = []
            for i in range(self.page.locator("input[type='button'][value='参照']").count()):
                href = self.page.locator("input[type='button'][value='参照']").nth(i).get_attribute("onclick")
                hrefs.append(re.search(r"href='([^']+)'", href).group(1))
            hrefs = list(set(hrefs))
            for href in hrefs:
                self.page.goto(
                    urljoin(self.domain, href),
                    wait_until="domcontentloaded",
                )
                square_metre = self.page.locator("input[name='project_articles[0][square_metre]']")
                if square_metre.get_attribute("value") == area:
                    return True
                square_metre.fill(area)
                square_metre.press("Tab")
                try:
                    with self.page.expect_navigation(wait_until="domcontentloaded"):
                        self.page.locator("button[id='order_update']").first.click()
                    notification: str = self.page.locator("div[class='f-content'] > div[class='data']").text_content().strip()
                    self.logger.info(notification)
                    return True
                except TimeoutError:
                    return False
        except TimeoutError:
            self.logger.error("RETRY update")
            return self.update(case, area)
