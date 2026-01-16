# --- excel_check.py ---
import logging
import os
import re
import time

import xlwings as xw
from builder_downloader import Builder_SharePoint_GraphAPI
from config_access_token import token_file  # noqa
from excel_conditions import ExcelConditionApplier
from graph_downloader import download_folder_by_anken
from logging_setup import setup_logging
from Nasiwak import Webaccess, create_driver, create_json_config
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

# Replace with your actual file path
file_path = os.path.join(os.getcwd(), "Access_token", "Access_token.txt")
logging.info(f"file path for text file is: {file_path}")
# Open and read the file
with open(file_path, "r", encoding="utf-8") as file:
    content = file.read()
logging.info(f"Extracted text from .txt file is: {content}")


# === CONFIG ===
ACCESS_TOKEN = content
WEBACCESS_JSON_URL = "https://raw.githubusercontent.com/Nasiwak/Nasiwak-jsons/refs/heads/main/webaccess.json"

# 🚀 Setup Logging
setup_logging()


class Excel_check:

    def __init__(self):
        self.webaccess_config = create_json_config(WEBACCESS_JSON_URL, ACCESS_TOKEN)
        self.driver = create_driver()
        self.wb = Webaccess(self.webaccess_config)
        self.previous_builder_id = ""
        self.wb.WebAccess_login(self.driver)
        self.condition_applier = ExcelConditionApplier()

        logging.info("✅ WebAccess login successful")

    def data_fetching(self, anken_bango, builder_id):
        self.bango = anken_bango
        self.builder_id = builder_id

        self.fetch_from_webaccess()
        self.extract_files(self.bango, self.builder_id)

    def fetch_from_webaccess(self):
        try:
            driver = self.driver
            driver.switch_to.window(driver.window_handles[0])

            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, self.webaccess_config["xpaths"]["受注一覧"]))
            )
            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, self.webaccess_config["xpaths"]["受注一覧"]))
            ).click()

            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, self.webaccess_config["xpaths"]["受注一覧_xpaths"]["リセット"])
                )
            )
            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, self.webaccess_config["xpaths"]["受注一覧_xpaths"]["リセット"]))
            ).click()

            driver.execute_script("scroll(0, 0);")

            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, self.webaccess_config["xpaths"]["受注一覧_xpaths"]["確定納品日_1"])
                )
            )
            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable(
                    (By.XPATH, self.webaccess_config["xpaths"]["受注一覧_xpaths"]["確定納品日_1"])
                )
            ).clear()

            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, self.webaccess_config["xpaths"]["受注一覧_xpaths"]["案件番号"])
                )
            )
            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, self.webaccess_config["xpaths"]["受注一覧_xpaths"]["案件番号"]))
            ).send_keys(self.bango)
            logging.info(f"🔍 Searching for {self.bango}")

            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, self.webaccess_config["xpaths"]["受注一覧_xpaths"]["検索"]))
            )
            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, self.webaccess_config["xpaths"]["受注一覧_xpaths"]["検索"]))
            ).click()
            logging.info(f"✅ Search completed for {self.bango}")
            time.sleep(3)

            # try:
            #     table = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "orderlist")))
            #     last_height = -1

            #     while True:
            #         rows = table.find_elements(By.CSS_SELECTOR, "tbody tr")
            #         if not rows:
            #             break
            #         last_row = rows[-1]

            #         # Scroll to last row
            #         driver.execute_script("arguments[0].scrollIntoView({block: 'end'});", last_row)
            #         time.sleep(0.5)

            #         new_height = len(rows)
            #         if new_height == last_height:
            #             break
            #         last_height = new_height
            # except Exception as e:
            #     logging.info(f"⚠️ Table scroll failed: {e}")

            time.sleep(2)

            # Ensure the row is loaded
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#orderlist tbody tr")))
            first_row = driver.find_element(By.CSS_SELECTOR, "#orderlist tbody tr")
            cells = first_row.find_elements(By.TAG_NAME, "td")

            # Column map: adjust based on exact DOM inspection
            self.Builder_name = cells[8].text.strip()  # Builder
            self.Address = cells[20].text.strip()  # Address
            self.shinki_status = cells[13].text.strip()  # Status

            logging.info(f"✅ Builder: {self.Builder_name}")
            logging.info(f"✅ Address: {self.Address}")
            logging.info(f"✅ Status: {self.shinki_status}")

            if self.shinki_status == "新規":
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located(
                        (By.XPATH, self.webaccess_config["xpaths"]["受注一覧_xpaths"]["参照_目地1"])
                    )
                )
                self.目地 = (
                    WebDriverWait(driver, 10)
                    .until(
                        EC.element_to_be_clickable(
                            (By.XPATH, self.webaccess_config["xpaths"]["受注一覧_xpaths"]["参照_目地1"])
                        )
                    )
                    .text
                    or "0"
                )
                logging.info(f"✅ 目地: {self.目地}")

                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located(
                        (By.XPATH, self.webaccess_config["xpaths"]["受注一覧_xpaths"]["参照_入隅1"])
                    )
                )
                self.入隅 = (
                    WebDriverWait(driver, 10)
                    .until(
                        EC.element_to_be_clickable(
                            (By.XPATH, self.webaccess_config["xpaths"]["受注一覧_xpaths"]["参照_入隅1"])
                        )
                    )
                    .text
                    or "0"
                )
                logging.info(f"✅ 入隅: {self.入隅}")

                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located(
                        (By.XPATH, self.webaccess_config["xpaths"]["受注一覧_xpaths"]["参照_配送時特記事項1"])
                    )
                )
                self.配送時特記事項 = (
                    WebDriverWait(driver, 10)
                    .until(
                        EC.element_to_be_clickable(
                            (By.XPATH, self.webaccess_config["xpaths"]["受注一覧_xpaths"]["参照_配送時特記事項1"])
                        )
                    )
                    .text
                )
                logging.info(f"✅ 配送時特記事項1: {self.配送時特記事項}")

            elif self.shinki_status == "先行":
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located(
                        (By.XPATH, self.webaccess_config["xpaths"]["受注一覧_xpaths"]["参照_目地2"])
                    )
                )
                self.目地 = (
                    WebDriverWait(driver, 10)
                    .until(
                        EC.element_to_be_clickable(
                            (By.XPATH, self.webaccess_config["xpaths"]["受注一覧_xpaths"]["参照_目地2"])
                        )
                    )
                    .text
                    or "0"
                )
                logging.info(f"✅ 新規目地: {self.目地}")

                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located(
                        (By.XPATH, self.webaccess_config["xpaths"]["受注一覧_xpaths"]["参照_入隅2"])
                    )
                )
                self.入隅 = (
                    WebDriverWait(driver, 10)
                    .until(
                        EC.element_to_be_clickable(
                            (By.XPATH, self.webaccess_config["xpaths"]["受注一覧_xpaths"]["参照_入隅2"])
                        )
                    )
                    .text
                    or "0"
                )
                logging.info(f"✅ 新規入隅: {self.入隅}")

                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located(
                        (By.XPATH, self.webaccess_config["xpaths"]["受注一覧_xpaths"]["参照_配送時特記事項2"])
                    )
                )
                self.配送時特記事項 = (
                    WebDriverWait(driver, 10)
                    .until(
                        EC.element_to_be_clickable(
                            (By.XPATH, self.webaccess_config["xpaths"]["受注一覧_xpaths"]["参照_配送時特記事項2"])
                        )
                    )
                    .text
                )
                logging.info(f"✅ 配送時特記事項2: {self.配送時特記事項}")

            if str(self.builder_id).strip() == "022001":
                ad_translation = {
                    "大阪工場引取 本社": "引取 (本社)",
                    "大阪工場引取 高槻倉庫": "引取 (高槻倉庫)",
                    "大阪工場引取 西宮倉庫": "引取 (西宮倉庫)",
                }
                matched = False
                for key in ad_translation.keys():
                    if key in self.配送時特記事項:
                        self.Address = ad_translation[key]
                        matched = True
                        break
                if not matched:
                    logging.warning("⚠️ No matching special address found for 022001 builder")

            logging.info(f"✅ Final Address: {self.Address}")

        except Exception as e:
            logging.warning(f"⚠️ Error fetching data: {e}")
            raise Exception("Skip to next 案件") from e

    def extract_files(self, id, builder_id):
        folder_path = os.path.join(os.getcwd(), "Ankens", f"{self.builder_id} {self.Builder_name}", id)
        os.makedirs(folder_path, exist_ok=True)
        try:
            download_folder_by_anken(self.bango, folder_path)
            # 🔍 NEW: sanity check
            files_after_download = os.listdir(folder_path)
            # logging.debug(f"📂 Files after download: {files_after_download}")

            if not files_after_download:
                raise Exception("❌ Downloaded folder is empty")

            logging.info(f"✅ Downloaded folder: {folder_path}")

            for file in os.listdir(folder_path):
                if file.endswith(".pdf") and not file.startswith("★"):
                    old_path = os.path.join(folder_path, file)
                    new_path = os.path.join(folder_path, f"★{file}")
                    os.rename(old_path, new_path)
                    logging.info(f"⭐ Renamed PDF: {file} ➡️ ★{file}")

        except Exception as e:
            logging.warning(f"⚠️ Could not download folder: {e}")
            self.delete_folder(id)
            raise Exception("Skip to next 案件") from e

        excel_files = [
            f
            for f in os.listdir(folder_path)
            if f.endswith(".xls")
            and (
                f.endswith(
                    (
                        "階.xls",
                        "屋.xls",
                        "のみ.xls",
                        "1).xls",
                        "2).xls",
                        "3).xls",
                        "4).xls",
                        "5).xls",
                        "6).xls",
                        "7).xls",
                        "8).xls",
                        "9).xls",
                        "①.xls",
                        "②.xls",
                        "③.xls",
                        "④.xls",
                        "⑤.xls",
                        "⑥.xls",
                        "⑦.xls",
                        "⑧.xls",
                        "⑨.xls",
                        "⑩.xls",
                        "⑪.xls",
                        "⑫.xls",
                        "⑬.xls",
                        "⑭.xls",
                        "⑮.xls",
                        "小物.xls",
                        "原本.xls",
                        "(A).xls",
                        "(B).xls",
                        "(C).xls",
                        "(D).xls",
                        "(E).xls",
                        "(F).xls",
                        "(G).xls",
                        "(H).xls",
                        "(I).xls",
                    )
                )
                or re.search(r"\(\d{3}\)\.xls$", f)
            )
        ]
        logging.debug(f"✅ Matched Excel files: {excel_files}")

        if not excel_files:
            logging.warning("⚠️ No Excel files found inside folder")
            return

        if self.previous_builder_id != self.builder_id:
            self.builder_id_drive()
            self.previous_builder_id = self.builder_id

        self.builder_copy(id)

    def detect_max_floor_for_builder_copy(self, excel_files):
        # Pick highest floor number
        if any("3階" in name for name in excel_files):
            return 3
        if any("2階" in name for name in excel_files):
            return 2
        if any("2階のみ" in name for name in excel_files):
            return 2
        return 1

    def builder_id_drive(self):
        destination_folder = os.path.join(os.getcwd(), "Ankens", f"{self.builder_id} {self.Builder_name}")
        os.makedirs(destination_folder, exist_ok=True)
        Builder_SharePoint_GraphAPI(builder_file_name=f"{self.builder_id}.xlsx", local_folder_path=destination_folder)
        logging.info(f"✅ Builder Excel {self.builder_id}.xlsx downloaded")

    def builder_copy(self, id):
        app = xw.App(visible=False)
        try:
            excel_folder = os.path.join("Ankens", f"{self.builder_id} {self.Builder_name}", id)
            source_wb_path = os.path.join("Ankens", f"{self.builder_id} {self.Builder_name}", f"{self.builder_id}.xlsx")
            source_wb = app.books.open(source_wb_path)

            # excel_files = [f for f in os.listdir(excel_folder) if f.endswith(('階.xls', '屋.xls', 'のみ.xls', '1).xls', '2).xls', '3).xls', '4).xls', '5).xls', '6).xls', '7).xls', '8).xls', '9).xls', '①.xls', '②.xls', '③.xls', '④.xls', '⑤.xls', '⑥.xls', '⑦.xls', '⑧.xls', '⑨.xls', '⑩.xls', '⑪.xls', '⑫.xls', '⑬.xls', '⑭.xls', '⑮.xls', '小物.xls', '原本.xls'))]
            excel_files = [
                f
                for f in os.listdir(excel_folder)
                if f.endswith(".xls")
                and (
                    f.endswith(
                        (
                            "階.xls",
                            "屋.xls",
                            "のみ.xls",
                            "1).xls",
                            "2).xls",
                            "3).xls",
                            "4).xls",
                            "5).xls",
                            "6).xls",
                            "7).xls",
                            "8).xls",
                            "9).xls",
                            "①.xls",
                            "②.xls",
                            "③.xls",
                            "④.xls",
                            "⑤.xls",
                            "⑥.xls",
                            "⑦.xls",
                            "⑧.xls",
                            "⑨.xls",
                            "⑩.xls",
                            "⑪.xls",
                            "⑫.xls",
                            "⑬.xls",
                            "⑭.xls",
                            "⑮.xls",
                            "小物.xls",
                            "原本.xls",
                        )
                    )
                    or re.search(r"\(\d{3}\)\.xls$", f)
                )
            ]

            # 🔍 Detect floor count once for all files in the folder
            builder_floor = self.detect_max_floor_for_builder_copy(excel_files)

            for file_name in excel_files:
                dest_wb_path = os.path.join(excel_folder, file_name)
                destination_wb = app.books.open(dest_wb_path)

                try:
                    # 🛠 Detect floor individually for each file
                    if "1階" in file_name:
                        floor = 1
                    elif "2階" in file_name:
                        floor = 2
                    elif "3階" in file_name:
                        floor = 3
                    else:
                        floor = 1  # fallback

                    logging.info(f"📄 Processing {file_name} | Detected Floor: {floor}")

                    destination_sheet = destination_wb.sheets["その他"]
                    source_sheet = source_wb.sheets[f"{builder_floor}"]

                    destination_sheet.range("B5").value = source_sheet.range("A1:H19").value
                    if floor == 1:
                        destination_sheet["J9"].value = self.目地
                        destination_sheet["J8"].value = self.入隅

                    stock_sheet = destination_wb.sheets["野縁"]
                    stock_sheet["AE3"].value = self.Address
                    stock_sheet["AE5"].value = self.Builder_name

                    # ✨ Apply floor-wise conditions
                    self.condition_applier.apply_conditions(destination_sheet, stock_sheet, self.builder_id, floor)
                    self.condition_applier.apply_first_free_cell_conditions(
                        destination_sheet, stock_sheet, self.builder_id, floor
                    )

                    self.run_macro2(destination_wb)
                    self.run_sort_macro(destination_wb)

                    destination_wb.save()
                    logging.info(f"✅ Processed and updated {file_name}")

                finally:
                    if destination_wb:
                        destination_wb.close()

                # ⭐ Rename processed file
                old_path = dest_wb_path
                new_path = os.path.join(excel_folder, f"★{file_name}")

                if os.path.exists(old_path):
                    os.rename(old_path, new_path)
                    logging.info(f"✅ Renamed {file_name} ➡️ ★{file_name}")

        except Exception as e:
            logging.error(f"❌ builder_copy critical error: {e}")
            raise Exception("Builder copy failed") from e
        finally:
            source_wb.close()
            app.quit()

    def run_macro2(self, workbook):
        sheet_main = workbook.sheets["野縁"]
        sheet_data = workbook.sheets["製作用データ"]
        sheet_main.range("U11").value = sheet_main.range("Q11:R110").value
        sorted_data = sorted(sheet_main.range("U11:V110").value, reverse=True)
        sheet_main.range("U11:V110").value = sorted_data
        sheet_data.range("A1").clear_contents()
        sheet_data.range("A3:D73").clear_contents()
        sheet_data.range("A1").value = sheet_main.range("E3").value
        dynamic_range = sheet_main.range((11, 30), (11 + int(sheet_main.range((111, 41)).value) - 1, 33))
        sheet_data.range("A3").value = dynamic_range.value
        workbook.save()

    def run_sort_macro(self, workbook):
        sheet1 = workbook.sheets["その他"]
        last_cell = sheet1.range("Q" + str(sheet1.cells.last_cell.row)).end("up").row
        sort_range = sheet1.range(f"Q5:T{last_cell}")
        sort_range.value = sorted(sort_range.value, reverse=True)
        workbook.save()

    def delete_folder(self, id):
        folder = os.path.join(os.getcwd(), "Ankens", f"{self.builder_id} {self.Builder_name}", id)
        if os.path.exists(folder):
            for file in os.listdir(folder):
                os.remove(os.path.join(folder, file))
            os.rmdir(folder)


# === MAIN ===
if __name__ == "__main__":
    obj = Excel_check()
    obj.data_fetching("481416", "016600")
