import os
import re
from typing import List
from urllib.parse import urlparse

from playwright._impl._errors import TimeoutError
from playwright.sync_api import Locator, sync_playwright


class SharePoint:
    def __init__(
        self,
        domain: str,
        email: str,
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
        self.email = email
        self.password = password
        self.timeout = timeout
        self.logger = logger
        if not self.__authentication():
            raise PermissionError("Authentication failed")

    def __authentication(self) -> bool:
        try:
            self.page.bring_to_front()
            self.page.goto(self.domain, wait_until="domcontentloaded")
            if urlparse(self.page.url).netloc != urlparse(self.domain).netloc:
                try:
                    self.page.wait_for_selector("input[type='email']", state="visible", timeout=10000).fill(self.email)
                    self.page.locator(
                        selector="input[type='submit']",
                        has_text="Next",
                    ).click(timeout=self.timeout)
                    self.page.wait_for_selector("input[type='password']", state="visible", timeout=10000).fill(self.password)
                    self.page.locator(
                        selector="input[type='submit']",
                        has_text="Sign in",
                    ).click(timeout=self.timeout)
                    with self.page.expect_navigation(
                        url=re.compile(f"^{self.domain}"),
                        wait_until="load",
                        timeout=30000,
                    ):
                        self.page.locator("input[type='submit']", has_text="Yes").click()
                except TimeoutError:
                    error: Locator = self.page.wait_for_selector(
                        "div#usernameError, div#passwordError",
                    )
                    self.logger.error(error.text_content())
                    return False
            while True:
                try:
                    self.page.wait_for_selector("div[id='HeaderButtonRegion']", state="visible", timeout=10000)
                    self.page.wait_for_selector(selector="div[id='O365_MainLink_MePhoto']", timeout=300, state="visible").click()
                    currentAccount_primary = self.page.wait_for_selector(
                        selector="div[id='mectrl_currentAccount_primary']", timeout=300, state="visible"
                    ).text_content()
                    currentAccount_secondary = self.page.wait_for_selector(
                        selector="div[id='mectrl_currentAccount_secondary']", timeout=300, state="visible"
                    ).text_content()
                    self.logger.info(f"Logged in as: {currentAccount_primary} ({currentAccount_secondary})")
                    break
                except TimeoutError:
                    self.page.reload(wait_until="domcontentloaded")
            return True
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

    def download_files(
        self,
        url: str,
        file: re.Pattern | str,
        steps: List[str | re.Pattern] | None = None,
        save_to: str | None = None,
    ) -> List[str]:
        try:
            self.page.bring_to_front()
            downloads = []
            self.page.goto(url=url)
            self.page.wait_for_selector(
                selector="div[class='app-container']",
                timeout=10000,
                state="visible",
            )
            for step in steps or []:
                span: Locator = self.page.locator(
                    selector="span[role='button']",
                    has_text=step,
                )
                try:
                    span.first.wait_for(timeout=self.timeout, state="visible")
                except TimeoutError:
                    return []
                if span.count() != 1:
                    return []
                text = span.text_content()
                span.click()
                self.page.locator(
                    selector="div[type='button'][data-automationid='breadcrumb-crumb']:visible",
                    has_text=text,
                ).wait_for(
                    timeout=self.timeout,
                    state="visible",
                )
            self.page.locator(
                selector="span[role='button'][data-id='heroField']",
                has_text=file,
            ).first.wait_for(timeout=5000, state="visible")
            items: Locator = self.page.locator(
                selector="span[role='button'][data-id='heroField']",
                has_text=file,
            )
            for i in range(items.count()):
                item: Locator = items.nth(i)
                item.click(button="right")
                download_btn = self.page.locator("button[data-automationid='downloadCommand'][role='menuitem']:not([type='button'])")
                download_btn.wait_for(state="visible", timeout=self.timeout)
                with self.page.expect_download() as download_info:
                    download_btn.click()
                download = download_info.value
                if save_to:
                    save_path = os.path.join(save_to, download.suggested_filename)
                else:
                    save_path = os.path.abspath(download.suggested_filename)
                os.makedirs(os.path.dirname(save_path), exist_ok=True)
                download.save_as(save_path)
                downloads.append(save_path)
                self.logger.info(f"Download {download.suggested_filename}")
            return downloads
        except TimeoutError:
            if self.page.locator("div[id='ms-error-content']").count() == 1:
                notification: str = self.page.locator("div[id='ms-error-content']").text_content().strip().split("\n")[0]
                self.logger.error(notification)
                if notification == "このドキュメントにアクセスできません。このドキュメントを共有するユーザーに連絡してください。":
                    return []
            self.logger.error("RETRY download_files")
            return self.download_files(url, file, steps, save_to)
