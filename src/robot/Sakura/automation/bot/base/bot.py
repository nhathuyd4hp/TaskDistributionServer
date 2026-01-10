import logging
import os
import time
from abc import ABC, abstractmethod

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support.ui import WebDriverWait

from .meta import IBotMeta


class IBot(ABC, metaclass=IBotMeta):
    def __init__(
        self,
        options: Options = webdriver.ChromeOptions(),
        service: Service = webdriver.ChromeService(),
        keep_alive: bool = False,
        timeout: float = 5,
        retry_interval: float = 0.5,
        logger: logging.Logger = None,
    ):
        self.timeout = timeout
        self.retry_interval = retry_interval
        self.options = options
        self.service = service
        self.browser = webdriver.Chrome(
            options=self.options,
            service=self.service,
            keep_alive=keep_alive,
        )
        self.browser.maximize_window()
        self.wait = WebDriverWait(
            driver=self.browser,
            timeout=self.timeout,
            poll_frequency=self.retry_interval,
        )
        self.logger = logger
        self.root_window = self.browser.window_handles[0]
        self.authenticated = False
        self.download_directory = self.options.experimental_options.get("prefs", {}).get(
            "download.default_directory",
            os.path.join(os.path.expanduser("~"), "downloads"),
        )
        os.makedirs(self.download_directory, exist_ok=True)

    def __dell__(self):
        try:
            self.browser.quit()
        except Exception:
            pass

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.browser.quit()

    def navigate(self, url, wait_for_complete: bool = True):
        time.sleep(self.retry_interval)
        self.browser.execute_script("window.stop()")
        time.sleep(self.retry_interval)
        self.browser.get(url)
        time.sleep(self.retry_interval)
        if wait_for_complete:
            while self.browser.execute_script("return document.readyState") != "complete":
                time.sleep(self.retry_interval)
        time.sleep(self.retry_interval)

    @abstractmethod
    def _authentication(self, **kwargs) -> bool:
        """Abstract method to implement authentication"""
        pass

    def open_new_tab(self) -> int:
        before_windows = self.browser.window_handles
        self.browser.execute_script("window.open('');")
        after_windows = self.browser.window_handles
        return list(set(after_windows) - set(before_windows))[0]

    def wait_for_download_to_start(self) -> list[str]:
        while True:
            downloading_files = [
                filename for filename in os.listdir(self.download_directory) if filename.endswith(".crdownload")
            ]
            if downloading_files:
                return downloading_files

    def wait_for_download_to_finish(self) -> tuple[str, str]:
        window_id = self.open_new_tab()
        self.browser.switch_to.window(window_id)
        self.navigate("chrome://downloads")
        download_items: list[WebElement] = self.browser.execute_script(
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
