import logging
import os

# ==================================================
# LOAD XPATH CONFIG
# ==================================================
from config_access_token import token_file  # noqa
from Nasiwak import create_json_config
from playwright.sync_api import Page
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError

# Replace with your actual file path
file_path = os.path.join(os.getcwd(), "Access_token", "Access_token.txt")
# logging.info(f"file path for text file is: {file_path}")
# Open and read the file
with open(file_path, "r", encoding="utf-8") as file:
    content = file.read()
# logging.info(f"Extracted text from .txt file is: {content}")


ACCESS_TOKEN = content

MailDealer_config_url = "https://raw.githubusercontent.com/Nasiwak/Nasiwak-jsons/refs/heads/main/MailDealer.json"
WebAccess_config_url = "https://raw.githubusercontent.com/Nasiwak/Nasiwak-jsons/refs/heads/main/webaccess.json"
Kizuku_config_url = "https://raw.githubusercontent.com/Nasiwak/Nasiwak-jsons/refs/heads/main/kizuku.json"
A1_config_url = "https://raw.githubusercontent.com/Nasiwak/Nasiwak-jsons/refs/heads/main/A1.json"

try:
    Maildealer_Data = create_json_config(MailDealer_config_url, ACCESS_TOKEN)
    Webaccess_Data = create_json_config(WebAccess_config_url, ACCESS_TOKEN)
    Kizuku_Data = create_json_config(Kizuku_config_url, ACCESS_TOKEN)
    A1_Data = create_json_config(A1_config_url, ACCESS_TOKEN)
    logging.info("✅ Configs loaded successfully.")
except Exception as e:
    logging.error(f"❌ Failed to load configs: {e}")


# ==================================================
# LOGIN (Selenium-equivalent, explicit)
# ==================================================
def Accesslogin(page: Page) -> bool:
    """
    Explicit WebAccess login.
    Mirrors Selenium Accesslogin exactly.
    """

    try:
        x = Webaccess_Data["xpaths"]["ログイン_xpaths"]

        page.goto(Webaccess_Data["webaccess_url"], wait_until="domcontentloaded")

        page.locator(f"xpath={x['ログインID']}").fill("NasiwakRobot")
        page.locator(f"xpath={x['パスワード']}").fill("159753")
        page.locator(f"xpath={x['ログイン']}").click()

        # Authoritative post-login proof
        page.wait_for_selector(f"xpath={Webaccess_Data['xpaths']['受注一覧']}", timeout=15000)

        logging.info("✅ WebAccess login complete.")
        return True

    except Exception as e:
        logging.error(f"❌ WebAccess login failed: {e}")
        return False


# ==================================================
# MAIN UPDATE FUNCTION
# ==================================================
def webaccess_update_drawing_status(page: Page, 案件番号: str) -> str:
    """
    Updates 図面 status in WebAccess.

    Returns:
        UPDATED   -> Status changed successfully
        NO_CHANGE -> Already final state
        FAILED    -> Attempted but failed
    """

    try:
        x = Webaccess_Data["xpaths"]

        # --------------------------------------------------
        # Ensure we are on WebAccess (tab is persistent)
        # --------------------------------------------------
        page.goto(Webaccess_Data["webaccess_url"], wait_until="domcontentloaded")

        # --------------------------------------------------
        # Go to 受注一覧
        # --------------------------------------------------
        page.locator(f"xpath={x['受注一覧']}").click()
        page.wait_for_timeout(500)

        # --------------------------------------------------
        # Reset filters
        # --------------------------------------------------
        page.locator(f"xpath={x['受注一覧_xpaths']['リセット']}").click()
        page.wait_for_timeout(300)

        # --------------------------------------------------
        # Enter 案件番号
        # --------------------------------------------------
        anken_input = page.locator(f"xpath={x['受注一覧_xpaths']['案件番号']}")
        anken_input.fill("")
        anken_input.fill(str(案件番号))

        page.locator(f"xpath={x['受注一覧_xpaths']['検索']}").click()

        # --------------------------------------------------
        # Click 参照
        # --------------------------------------------------
        try:
            page.locator(f"xpath={x['受注一覧_xpaths']['参照']}").wait_for(timeout=8000)

            page.locator(f"xpath={x['受注一覧_xpaths']['参照']}").click()

        except PlaywrightTimeoutError:
            logging.warning("❌ WebAccess: 参照 button not found")
            return "FAILED"

        # --------------------------------------------------
        # Read current 図面 status
        # --------------------------------------------------
        drawing_select = page.locator(f"xpath={x['案件詳細_xpaths']['図面']}")

        drawing_select.wait_for(state="visible", timeout=10000)

        # current_value = drawing_select.input_value()

        current_text = drawing_select.locator("option:checked").inner_text().strip()

        logging.info(f"Current 図面 status: {current_text}")

        # --------------------------------------------------
        # Decide next state
        # --------------------------------------------------
        target_value = None

        if current_text == "作図済":
            target_value = "7"  # 送付済
        elif current_text == "CBUP済":
            target_value = "8"  # CB送付済
        else:
            logging.info("ℹ WebAccess: no status change needed")
            return "NO_CHANGE"

        # --------------------------------------------------
        # Change status SAFELY
        # --------------------------------------------------
        drawing_select.select_option(value=target_value)

        # Verify change
        page.wait_for_function(
            """
            (select, val) => select.value === val
            """,
            arg=(drawing_select, target_value),
            timeout=5000,
        )

        # --------------------------------------------------
        # Save
        # --------------------------------------------------
        page.locator(f"xpath={x['案件詳細_xpaths']['案件情報を更新する']}").click()

        # Success message
        page.wait_for_selector(f"xpath={x['案件詳細_xpaths']['案件情報を更新しました']}", timeout=10000)

        logging.info("✅ WebAccess status updated successfully")

        return "UPDATED"

    except Exception as e:
        logging.error(f"🔥 WebAccess update failed: {e}")
        return "FAILED"
