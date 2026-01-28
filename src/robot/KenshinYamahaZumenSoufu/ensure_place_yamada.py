import logging

from dandoli_state import get_current_place_from_header, normalize_jp
from playwright.sync_api import Page

logger = logging.getLogger(__name__)

PLACE_SWITCH_BUTTON = "a.placeSwitchButton__currentPlace"
PLACE_LIST = "ul.popover__switchablePlaceList"
PLACE_ITEMS = "li.popover__switchablePlaceListItem"
PLACE_NAME = "div.popover__switchablePlaceListName"
COMPANY_NAME = "div.popover__switchableCompanyName"
PLACE_LINK = "a.popover__switchablePlaceListLink"
HOME_LOGO = "a.header__logo.js-move-to-dashboard"

YAMADA_SHITENS = [
    "★YH_不動産_第一営業部",
    "★YH_首都圏支店",
    "★YH_千葉支店",
    "★YH_中部東支店",
    "★YH_関西南支店",
    "★YH_神奈川東支店",
    "★YH_南東北支店",
    "★YH_群馬支店",
    "★YH_茨城支店",
    "★YH_九州北支店",
    "★YH_埼玉支店",
    "★YH_北東北支店",
    "★YH_九州南支店",
    "★YH_京滋支店",
    "★YH_北陸支店",
]


def force_true_home(page: Page):
    logger.info("🏠 Clicking logo to force true HOME")

    page.locator("a.header__logo.js-move-to-dashboard").click()

    # Just wait for the dropdown button itself
    page.wait_for_selector("a.placeSwitchButton__currentPlace", state="visible", timeout=20000)

    logger.info("✅ True HOME confirmed (place switch visible)")


def ensure_place_yamada(page: Page, shiten_name: str) -> bool:
    logger.info(f"🏢 Ensuring Yamada place | 支店名: {shiten_name}")

    force_true_home(page)

    page.locator(PLACE_SWITCH_BUTTON).click()
    page.wait_for_selector(PLACE_LIST, timeout=10000)

    items = page.locator(PLACE_ITEMS)
    count = items.count()
    logger.info(f"📋 Visible places: {count}")

    if not shiten_name or str(shiten_name).lower() == "nan":
        logger.warning("⚠ 支店名 is empty/nan – skipping Yamada switch")
        return False

    for configured in YAMADA_SHITENS:
        if normalize_jp(shiten_name) not in normalize_jp(configured):
            continue

        logger.info(f"🎯 Target Yamada config: {configured}")

        for i in range(count):
            item = items.nth(i)

            builder = item.locator(PLACE_NAME).inner_text().strip()
            company = item.locator(COMPANY_NAME).inner_text().strip()

            if normalize_jp(configured) not in normalize_jp(builder):
                continue

            logger.info(f"▶ Switching to Yamada place: {builder} / {company}")

            item.locator(PLACE_LINK).click()

            page.wait_for_function(
                """
                (expected) => {
                    const el = document.querySelector('.placeSwitchButton__placeName');
                    return el && el.innerText.includes(expected);
                }
                """,
                arg=builder,
                timeout=20000,
            )

            header = get_current_place_from_header(page)
            if normalize_jp(builder) not in normalize_jp(header):
                logger.warning(f"⚠ Header mismatch after switch: {header}")
                continue

            logger.info(f"✅ Yamada place selected: {builder}")
            page.keyboard.press("Escape")
            return True

    page.keyboard.press("Escape")
    logger.error(f"❌ Failed to resolve Yamada 支店: {shiten_name}")
    return False
