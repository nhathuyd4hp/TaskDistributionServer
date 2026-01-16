import logging
import time
from contextlib import suppress

import pandas as pd
from playwright._impl._errors import TimeoutError
from playwright.sync_api import sync_playwright


class MailDealer:
    def __init__(
        self,
        domain: str,
        username: str,
        password: str,
        context=None,
        browser=None,
        playwright=None,
        logger: logging.Logger = logging.getLogger("MailDealer"),
        headless: bool = False,
        timeout: float = 5000,
    ):
        self.logger = logger
        self._external_context = context is not None
        self._external_browser = browser is not None
        self._external_pw = playwright is not None
        self.timeout = timeout
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
        if not self.__authentication():
            raise PermissionError("Authentication failed")

    def __authentication(self) -> bool:
        try:
            self.page.bring_to_front()
            self.page.goto(self.domain, wait_until="domcontentloaded")
            self.page.locator(selector="input[id='fUName']").fill(self.username)
            self.page.locator(selector="input[id='fPassword']").fill(self.password)
            with self.page.expect_navigation(
                wait_until="networkidle",
                timeout=30000,
            ):
                self.page.locator(selector="input[type='submit'][value='ログイン']").click()
            self.page.locator("button[title='設定']").click(click_count=2, delay=0.25)
            with suppress(TimeoutError):
                self.page.wait_for_selector("div[id='md_dialog']", timeout=5000)
                self.page.locator("input[type='button'][id='md_dialog_submit']").click()
            return True
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

    def mail_box(self, mailbox: str) -> pd.DataFrame:
        try:
            # Switch Mailbox
            self.page.bring_to_front()
            self.page.frame(name="side").locator(f"span[title='{mailbox}']").click()
            # Get Table
            self.page.frame(name="main").wait_for_selector("table", state="visible")
            columns = self.page.frame(name="main").locator("table thead")
            columns: list[str] = [
                columns.locator("th").nth(i).text_content() for i in range(columns.locator("th").count())
            ]
            Rows = self.page.frame(name="main").locator("table tbody")
            data = [
                [Rows.nth(i).locator("td").nth(j).text_content() for j in range(Rows.nth(i).locator("td").count())]
                for i in range(Rows.count())
            ]
            if self.page.frame(name="main").locator("div[class='olv-p-maillist__no-data']").count() == 1:
                self.logger.warning("条件に一致するデータがありません。")
                return pd.DataFrame(columns=columns)
            data = pd.DataFrame(data=data, columns=columns)
            if not (data[" フォルダ "] == mailbox).all():
                return self.mail_box(mailbox)
            return data
        except TimeoutError:
            return self.mail_box(mailbox)

    def update_mail(self, mail_id: str, label: str, fMatterID: str, comment: str | None = None) -> bool | str:
        self.logger.info(f"Update mail: {mail_id}")
        try:
            self.page.bring_to_front()
            self.page.locator("input[name='fDQuery[B]']").clear()
            self.page.locator("input[name='fDQuery[B]']").fill(mail_id)
            self.page.locator("button[title='検索']").click()
            self.page.wait_for_selector("div[class='loader']", state="visible", timeout=10000)
            self.page.wait_for_selector("div[class='loader']", state="hidden", timeout=30000)
            # ---- #
            while True:
                if (
                    self.page.frame(name="main")
                    .locator("div[class='olv-p-mail-ops__act-status'] > div[class^='dropdown']")
                    .count()
                    == 4
                ):
                    break
                continue
            status = [
                self.page.frame(name="main")
                .locator("div[class='olv-p-mail-ops__act-status'] > div[class^='dropdown']")
                .nth(i)
                .text_content()
                for i in range(4)
            ]
            if "担当者指定なし" not in status:
                return "Đã có người làm"
            self.page.frame(name="main").locator(
                "div[class='dropdown__text is-default']", has_text="担当者指定なし"
            ).click()
            self.page.frame(name="main").locator("input[class='list__filter-input']").fill(label)
            time.sleep(0.5)
            if self.page.frame(name="main").locator("li[class^='list__item']").count() != 1:
                return "Lỗi gắn người phụ trách"
            self.page.frame(name="main").locator("li[class^='list__item']").click()
            self.page.frame(name="main").wait_for_selector("div[class^='snackbar is-success']", state="visible")
            # ---- #
            self.page.frame(name="main").locator("button[title='一括操作']").click()
            self.page.frame(name="main").locator("div[class='pop-panel__content'] input[id='fMatterID_add']").fill(
                fMatterID
            )
            self.page.frame(name="main").locator(
                "div[class='pop-panel__content'] input[name='fAddMatterRelByMGID'] + div.checkbox__indicator"
            ).check()
            if comment:
                self.page.frame(name="main").locator(
                    "div[class='pop-panel__content'] input[id='fMatterID_add'] + button", has_text="関連付ける"
                ).click()
                time.sleep(2.5)
                self.page.frame(name="main").locator("button[class='olv-p-mail-view-header__ops-comment']").click()
                self.page.frame(name="main").locator("textarea[name='fComment']").fill(comment)
                self.page.frame(name="main").locator("button", has_text="登録").click()
                time.sleep(2.5)
                return True
            else:
                self.page.frame(name="main").locator(
                    "div[class='pop-panel__content'] input[id='fMatterID_add'] + button", has_text="関連付ける"
                ).click()
                time.sleep(2.5)
                return True
        except TimeoutError:
            if self.page.frame(name="main").locator("div[class='alert alert-error']").count() == 1:
                notification = self.page.frame(name="main").locator("div[class='alert alert-error']").text_content()
                time.sleep(2.5)
                return self.update_mail(mail_id, label, fMatterID, comment)
            if self.page.frame(name="main").locator("div[class^='mailList-no-tab']").count() == 0:
                return self.update_mail(mail_id, label, fMatterID, comment)
            notification = self.page.frame(name="main").locator("div[class^='mailList-no-tab']").text_content().strip()
            if notification == "検索条件に一致するデータがありません。":
                return False
            return self.update_mail(mail_id, label, fMatterID, comment)
