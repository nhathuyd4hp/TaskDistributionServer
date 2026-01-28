import logging
import re
import unicodedata

from playwright.sync_api import Page

logger = logging.getLogger(__name__)

PLACE_SWITCH_BUTTON = "a.placeSwitchButton__currentPlace"
PLACE_LIST = "ul.popover__switchablePlaceList"
PLACE_ITEMS = "li.popover__switchablePlaceListItem"
PLACE_NAME = "div.popover__switchablePlaceListName"
COMPANY_NAME = "div.popover__switchableCompanyName"
CURRENT_TAG = "span.popover__switchablePlaceListCurrent"
PLACE_LINK = "a.popover__switchablePlaceListLink"


def normalize_jp(text: str) -> str:
    """
    Normalizes Japanese strings for safe comparison:
    - Unicode normalization
    - Collapse all whitespace
    """
    text = unicodedata.normalize("NFKC", text)
    text = re.sub(r"\s+", "", text)
    return text


def get_current_place_from_header(page: Page) -> str:
    """
    Reads current place from the header (authoritative after switch)
    """
    container = page.locator(".placeSwitchButton__placeName")
    text = container.inner_text().strip()
    return text


def ensure_home_screen(page: Page):
    """
    Confirms we are on Dandoli home/dashboard.
    """
    logger.info("⌛ Verifying Dandoli home screen")
    page.wait_for_selector(PLACE_SWITCH_BUTTON, timeout=15000)
    logger.info("✅ Home screen confirmed")


def get_current_place_from_dropdown(page: Page):
    logger.debug("⌛ Reading current place from dropdown")
    page.wait_for_selector(PLACE_LIST, timeout=10000)

    current = page.locator(f"{PLACE_ITEMS}:has({CURRENT_TAG})")

    if current.count() != 1:
        logger.error("❌ Cannot uniquely identify current place in dropdown")
        raise RuntimeError("Cannot uniquely identify current place")

    builder = current.locator(PLACE_NAME).inner_text().strip()
    company = current.locator(COMPANY_NAME).inner_text().strip()

    logger.info(f"📍 Current place (dropdown): {builder} / {company}")
    return builder, company


def find_place_option(page: Page, expected_builder: str, expected_company: str | None):
    items = page.locator(PLACE_ITEMS)

    logger.debug("🔍 Scanning place list for target")

    for i in range(items.count()):
        item = items.nth(i)

        builder = item.locator(PLACE_NAME).inner_text().strip()
        company = item.locator(COMPANY_NAME).inner_text().strip()

        if normalize_jp(expected_builder) not in normalize_jp(builder):
            continue

        if expected_company:
            if normalize_jp(expected_company) not in normalize_jp(company):
                continue

        logger.info(f"🎯 Target place found: {builder} / {company}")
        return item

    logger.warning(f"⚠ Target place not found in list: {expected_builder}")
    return None


def ensure_place(page: Page, expected_builder: str, expected_company: str | None = None):
    """
    Ensures the expected builder/place is selected in Dandoli.
    SAFE against pages where place switcher is not rendered.
    """

    # 🔐 HARD GUARD: place switcher may not exist on Genba pages
    if page.locator(PLACE_SWITCH_BUTTON).count() == 0:
        logger.info("ℹ Place switch button not present – assuming place unchanged")
        return

    logger.info("🔄 Opening place switcher")
    page.click(PLACE_SWITCH_BUTTON)
    page.wait_for_selector(PLACE_LIST, timeout=10000)

    current_builder, current_company = get_current_place_from_dropdown(page)

    # Early exit if builder already matches
    if normalize_jp(expected_builder) in normalize_jp(current_builder):
        logger.info("✅ Correct builder already selected")
        page.keyboard.press("Escape")
        return

    target = find_place_option(page, expected_builder, expected_company)

    if not target:
        screenshot_path = "place_not_found.png"
        page.screenshot(path=screenshot_path)
        logger.error(f"❌ Target place not found: {expected_builder}. Screenshot: {screenshot_path}")
        raise RuntimeError(f"Target place not found: {expected_builder}")

    logger.info(f"▶ Switching to place: {expected_builder}")
    target.locator(PLACE_LINK).click()

    # Allow SPA rebind
    page.wait_for_timeout(500)

    logger.info("⌛ Verifying place switch via header")
    page.wait_for_function(
        """
        (expected) => {
            const el = document.querySelector('.placeSwitchButton__placeName');
            return el && el.innerText.includes(expected);
        }
        """,
        arg=expected_builder,
        timeout=20000,
    )

    header_text = get_current_place_from_header(page)

    if normalize_jp(expected_builder) not in normalize_jp(header_text):
        screenshot_path = "builder_header_mismatch.png"
        page.screenshot(path=screenshot_path)
        logger.error(
            f"❌ Builder mismatch after place switch. Header: {header_text}. " f"Screenshot: {screenshot_path}"
        )
        raise RuntimeError("Builder mismatch after place switch (header check)")

    logger.info("✅ Place switch verified (header)")
    page.keyboard.press("Escape")
