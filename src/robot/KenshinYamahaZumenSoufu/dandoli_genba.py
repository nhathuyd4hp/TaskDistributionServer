import logging

from dandoli_state import normalize_jp
from playwright.sync_api import Page
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError

logger = logging.getLogger(__name__)

GENBA_MENU = "#nav-sites"
SEARCH_INPUT = "input.pageContentSite__searchItemKeyword[name='keyword']"
RESULT_ROWS = "table.pageContentSite__table tbody tr"


# -------------------------------------------------
# LOADER (single source of truth)
# -------------------------------------------------
def wait_for_loader_to_disappear(page: Page, timeout: int = 20000):
    logger.debug("⌛ Waiting for loader overlay to disappear")

    page.wait_for_function(
        """
        () => {
            const loader = document.querySelector('.loader-container');
            return !loader || loader.offsetParent === null;
        }
        """,
        timeout=timeout,
    )

    logger.debug("✅ Loader overlay gone")


# -------------------------------------------------
# GENBA HOME
# -------------------------------------------------
def ensure_genba_kanri_home(page: Page):
    """
    Ensures we are REALLY on 現場管理 (Genba Kanri) list screen.
    """
    logger.info("🏠 Navigating to Genba Kanri home")
    page.locator(GENBA_MENU).click()

    try:
        logger.info("⌛ Waiting for search input")
        page.wait_for_selector(SEARCH_INPUT, state="visible", timeout=20000)
        wait_for_loader_to_disappear(page)

    except PlaywrightTimeoutError as e:
        screenshot_path = "genba_home_failed.png"
        page.screenshot(path=screenshot_path)
        logger.error(f"❌ Failed to reach Genba Kanri. Screenshot: {screenshot_path}", exc_info=True)
        raise RuntimeError("Failed to reach Genba Kanri home") from e

    logger.info("✅ Genba Kanri home confirmed")


# -------------------------------------------------
# SEARCH
# -------------------------------------------------
def search_anken_by_name(page: Page, anken_name: str):
    logger.info(f"🔍 Searching by 現場名: {anken_name}")

    search = page.locator(SEARCH_INPUT)
    search.wait_for(state="visible", timeout=10000)

    search.click()
    search.fill("")
    search.type(anken_name, delay=30)
    search.press("Enter")


def wait_for_search_results(page: Page, keyword: str, timeout: int = 30000):
    """
    Waits until the table REFLECTS the search:
    - exactly ONE row
    - row text contains the keyword
    """
    logger.info("⌛ Waiting for Dandoli search results to reflect")

    keyword_norm = normalize_jp(keyword)

    page.wait_for_function(
        """
        (kw) => {
            const rows = document.querySelectorAll(
                'table.pageContentSite__table tbody tr'
            );
            if (rows.length !== 1) return false;

            const text = rows[0].innerText
                .replace(/\\s+/g, '')
                .normalize('NFKC');

            return text.includes(kw);
        }
        """,
        arg=keyword_norm,
        timeout=timeout,
    )

    wait_for_loader_to_disappear(page)
    logger.info("✅ Search results reflected in table")


# -------------------------------------------------
# ENTER ANKEN
# -------------------------------------------------
def enter_anken(page: Page):
    rows = page.locator(RESULT_ROWS)
    count = rows.count()

    if count != 1:
        raise RuntimeError(f"❌ Aborting enter_anken: expected 1 row after search, found {count}")

    row = rows.first
    name_cell = row.locator("td.pageContentSite__nameCol")

    logger.info("🎯 Clicking anken name cell (avoids map click)")
    name_cell.scroll_into_view_if_needed()
    name_cell.click()

    page.wait_for_url("**/sites/**", timeout=20000)
    wait_for_loader_to_disappear(page)

    logger.info("✅ Inside correct anken")


# -------------------------------------------------
# FORCE RETURN
# -------------------------------------------------
def force_return_to_genba_kanri(page: Page):
    logger.info("🔄 Returning to Genba Kanri (force reset)")

    page.locator(GENBA_MENU).click()
    page.wait_for_selector(SEARCH_INPUT, timeout=20000)
    wait_for_loader_to_disappear(page)

    logger.info("✅ Genba Kanri search page ready")
