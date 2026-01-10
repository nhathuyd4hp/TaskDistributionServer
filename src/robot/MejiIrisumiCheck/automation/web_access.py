import logging
import os

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
        logger: logging.Logger = logging.getLogger("WebAccess"),
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
                timeout=self.timeout,
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

    def download_orders(
        self,
        from_date: str,
        to_date: str,
        save_to: str | None = None,
    ) -> pd.DataFrame:
        try:
            with self.page.expect_navigation(wait_until="domcontentloaded"):
                self.page.wait_for_selector("a[class='fa fa-industry']").click()
            self.page.locator("button[type='reset']").click()
            self.page.locator("input[name='search_fix_deliver_date_from']").fill(from_date)
            self.page.locator("input[name='search_fix_deliver_date_to']").fill(to_date)
            with self.page.expect_navigation(wait_until="domcontentloaded"):
                self.page.locator("button[class='search fa fa-search']").click(force=True)
            with self.page.expect_download() as download_info:
                self.page.locator("a[class='button fa fa-download']").first.click()
                download = download_info.value
                if save_to:
                    save_path = os.path.join(save_to, download.suggested_filename)
                else:
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
            return self.download_orders(from_date, to_date, save_to)
