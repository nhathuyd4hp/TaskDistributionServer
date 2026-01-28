import logging

from playwright.sync_api import Page
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError

DANDOLI_URL = "https://www.dandoli.jp"
USERNAME = "kantou@nsk-cad.com"
PASSWORD = "nsk00426"

logger = logging.getLogger(__name__)


def login_dandoli(page: Page):
    """
    Logs into Dandoli using the provided Playwright page.
    Does NOT create or close browser/context.
    Relies on global logging setup from Main.
    """

    try:
        logger.info("🔐 Opening Dandoli login page")
        page.goto(DANDOLI_URL, wait_until="domcontentloaded")

        logger.info("⌛ Waiting for login form")
        page.wait_for_selector('input[name="username"]', timeout=15000)

        logger.info("✍ Entering credentials")
        page.fill('input[name="username"]', USERNAME)
        page.fill('input[name="password"]', PASSWORD)

        logger.info("➡ Submitting login")
        page.click('button[type="submit"]')

        logger.info("⌛ Waiting for dashboard / place selector")
        page.wait_for_selector("a.placeSwitchButton__currentPlace", timeout=20000)

        logger.info("✅ Dandoli login successful")

    except PlaywrightTimeoutError as e:
        screenshot_path = "login_failed_dandoli.png"
        page.screenshot(path=screenshot_path)

        logger.error(f"❌ Dandoli login failed or blocked. Screenshot saved: {screenshot_path}", exc_info=True)
        raise RuntimeError("Dandoli login failed") from e

    except Exception as e:
        logger.error("🔥 Unexpected error during Dandoli login", exc_info=True)
        raise e
