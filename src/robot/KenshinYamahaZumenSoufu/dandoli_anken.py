import logging
import os

from playwright.sync_api import Page
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError

logger = logging.getLogger(__name__)

GENBA_SHIRYO_TAB = (
    "li.site-document-tab "
    "a.pageContentSiteDetail__menuLink:has("
    "span.pageContentSiteDetail__menuText:text('現場資料'))"
)


def go_to_genba_shiryo(page: Page):
    """
    Clicks 現場資料 tab inside an anken and waits until
    the 現場資料 content is actually loaded.
    """
    logger.info("📂 Opening 現場資料 tab")

    tab = page.locator("a.pageContentSiteDetail__menuLink:has-text('現場資料')")
    tab.wait_for(state="visible", timeout=10000)
    tab.click(force=True)

    try:
        page.wait_for_function(
            """
            () => {
                return (
                    document.querySelector('.site-document')
                    || document.querySelector('.pageContentSiteDocument')
                    || document.body.innerText.includes('アップロード')
                    || document.body.innerText.includes('ファイル')
                );
            }
            """,
            timeout=20000,
        )
    except PlaywrightTimeoutError as e:
        screenshot_path = "genba_shiryo_not_loaded.png"
        page.screenshot(path=screenshot_path)
        logger.error(f"❌ 現場資料 tab did not load. Screenshot: {screenshot_path}", exc_info=True)
        raise RuntimeError("現場資料 tab click did not load content") from e

    logger.info("✅ 現場資料 tab opened and verified")


def open_bulk_upload_single_type(page: Page):
    """
    Clicks 「１つの種類で一括登録」 and waits for the upload modal to appear.
    """
    logger.info("🧾 Opening 「１つの種類で一括登録」 modal")

    button = page.locator("button:has-text('１つの種類で一括登録')")
    button.wait_for(state="visible", timeout=10000)
    button.click(force=True)

    try:
        page.wait_for_function(
            """
            () => {
                return (
                    document.body.innerText.includes('ファイルを選択')
                    || document.body.innerText.includes('ドラッグ')
                    || document.querySelector('.modal')
                );
            }
            """,
            timeout=20000,
        )
    except PlaywrightTimeoutError as e:
        screenshot_path = "bulk_upload_modal_not_opened.png"
        page.screenshot(path=screenshot_path)
        logger.error(f"❌ Bulk upload modal did not open. Screenshot: {screenshot_path}", exc_info=True)
        raise RuntimeError("一括登録 modal did not open") from e

    logger.info("✅ 「１つの種類で一括登録」 modal opened")


def select_upload_type_shosetsu_kensetsu(page: Page):
    """
    Selects the correct upload type:
    - Prefer 「住設・建材 承認図」
    - Fallback to 「軽天割付図」 (Yamada Homes)
    """
    logger.info("🔽 Selecting upload type (auto-detect)")

    modal = page.locator("#sites_document-modal")
    modal.wait_for(state="visible", timeout=10000)

    select = modal.locator("select[name='type_id']")
    select.wait_for(state="visible", timeout=10000)

    # Get all options text
    options = select.locator("option")
    option_count = options.count()

    found_value = None
    # found_label = None

    for i in range(option_count):
        opt = options.nth(i)
        label = opt.inner_text().strip()

        if "住設・建材" in label:
            found_value = opt.get_attribute("value")
            # found_label = label
            break

    # Fallback for Yamada Homes
    if not found_value:
        for i in range(option_count):
            opt = options.nth(i)
            label = opt.inner_text().strip()

            if "軽天割付図" in label:
                found_value = opt.get_attribute("value")
                # found_label = label
                break

    if not found_value:
        screenshot_path = "upload_type_not_found.png"
        modal.screenshot(path=screenshot_path)
        logger.error("❌ No suitable upload type found (住設・建材 / 軽天割付図). " f"Screenshot: {screenshot_path}")
        raise RuntimeError("No valid upload type found")

    select.select_option(value=found_value)

    selected_text = select.locator("option:checked").inner_text().strip()

    logger.info(f"✅ Upload type selected: {selected_text}")


def upload_single_pdf(page: Page, pdf_path: str):
    if not os.path.exists(pdf_path):
        logger.error(f"❌ PDF file not found: {pdf_path}")
        raise RuntimeError(f"File not found: {pdf_path}")

    filename = os.path.basename(pdf_path)
    logger.info(f"📎 Uploading file: {filename}")

    modal = page.locator("#sites_document-modal")
    modal.wait_for(state="visible", timeout=10000)

    file_input = modal.locator("input.file-input[type='file']")

    if file_input.count() != 1:
        screenshot_path = "file_input_ambiguous.png"
        modal.screenshot(path=screenshot_path)
        logger.error(f"❌ File input not found or ambiguous. Screenshot: {screenshot_path}")
        raise RuntimeError("File input not found or ambiguous")

    file_input.set_input_files(pdf_path)

    page.wait_for_function(
        """
        (name) => {
          const modal = document.querySelector('#sites_document-modal');
          return modal && modal.innerText.includes(name);
        }
        """,
        arg=filename,
        timeout=20000,
    )

    logger.info("✅ File attached and visible in upload list")


def enter_file_description(page: Page, note: str):
    logger.info("📝 Entering file description")

    modal = page.locator("#sites_document-modal")
    modal.wait_for(state="visible", timeout=10000)

    desc_input = modal.locator("input[name='desc']")

    if desc_input.count() != 1:
        screenshot_path = "desc_input_not_unique.png"
        modal.screenshot(path=screenshot_path)
        logger.error(f"❌ Description input not unique. Screenshot: {screenshot_path}")
        raise RuntimeError("File description input not found or ambiguous")

    desc_input.fill("")
    desc_input.type(note, delay=50)

    page.wait_for_function(
        """
        (value) => {
          const el = document.querySelector(
            '#sites_document-modal input[name="desc"]'
          );
          return el && el.value === value;
        }
        """,
        arg=note,
        timeout=5000,
    )

    logger.info(f"✅ File description set: {note}")


def submit_upload(page: Page):
    logger.info("💾 Submitting upload (編集を実行)")

    page.locator("button.update").click()

    page.locator("div.modal-content").wait_for(state="hidden", timeout=20000)

    page.wait_for_selector("button.js-btn-open-content:has-text('通知する')", timeout=20000)

    logger.info("✅ Upload submitted, modal closed")


def confirm_notification(page: Page):
    logger.info("📣 Confirming 通知する")

    confirm = page.locator("div.confirm-content")
    confirm.wait_for(state="visible", timeout=20000)

    confirm.locator("button.js-btn-open-content").click()
    confirm.wait_for(state="hidden", timeout=15000)

    logger.info("✅ Notification confirmed")


def select_all_except_nsk(page: Page):
    logger.info("👥 Selecting all users except NSK")

    participant_panel = page.locator("#sites_change_notification_user_list-page-layout")
    participant_panel.wait_for(state="attached", timeout=15000)

    participant_panel.locator("button.js-select-all-btn[data-check-all='true']").click()

    page.wait_for_timeout(500)

    nsk_row = participant_panel.locator("li.user-list-item.is-parent:has-text('エヌ・エス・ケー工業㈱')")

    if nsk_row.count() != 1:
        screenshot_path = "nsk_parent_not_found.png"
        page.screenshot(path=screenshot_path)
        logger.error(f"❌ NSK parent row not uniquely found. Screenshot: {screenshot_path}")
        raise RuntimeError("NSK parent row not uniquely found")

    checkbox = nsk_row.locator("input.list-checkbox")

    if checkbox.is_checked():
        logger.info("🚫 Excluding NSK recipients")
        checkbox.click()
    else:
        logger.info("ℹ NSK already unselected")

    if checkbox.is_checked():
        raise RuntimeError("NSK checkbox still selected after exclusion")

    logger.info("✅ Notification recipients set")


def move_users_to_receiver(page: Page):
    logger.info("➡️ Moving selected users to 宛先")

    add_btn = page.locator("button.js-add-users-btn:has-text('宛先に追加')")
    add_btn.wait_for(state="visible", timeout=10000)

    page.wait_for_function("() => !document.querySelector('button.js-add-users-btn').disabled", timeout=10000)

    add_btn.click()
    logger.info("✅ Users moved to 宛先 panel")


def enter_notification_comment(page: Page):
    logger.info("📝 Entering notification comment")

    comment_text = (
        "お世話になっております。\n" "軽天割付図をUPしましたのでご確認お願い致します。\n" "宜しくお願い致します。"
    )

    comment_box = page.locator("div.comment.js-comment")
    comment_box.wait_for(state="visible", timeout=10000)

    comment_box.click()
    comment_box.type(comment_text, delay=20)

    page.wait_for_timeout(300)
    entered_text = comment_box.inner_text()

    if "軽天割付図" not in entered_text:
        screenshot_path = "comment_not_entered.png"
        page.screenshot(path=screenshot_path)
        logger.error(f"❌ Notification comment not entered correctly. Screenshot: {screenshot_path}")
        raise RuntimeError("Notification comment not entered correctly")

    logger.info("✅ Notification comment entered successfully")


def send_notification(page: Page):
    logger.info("🚀 Sending notification")

    send_btn = page.locator("button.js-send-site-notification-btn:has-text('送信')")
    send_btn.wait_for(state="visible", timeout=10000)

    page.wait_for_function(
        "() => !document.querySelector('button.js-send-site-notification-btn').disabled", timeout=10000
    )

    send_btn.click()
    page.wait_for_timeout(1000)

    logger.info("✅ Notification sent successfully")
