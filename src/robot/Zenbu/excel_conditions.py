# --- excel_conditions.py (Final Corrected Version) ---

import logging

import pandas as pd


class ExcelConditionApplier:
    def __init__(self):
        try:
            self.conditions_df = pd.read_excel("demo.xlsx", dtype=str)
            logging.info("✅ Demo Excel loaded for conditions")
        except Exception as e:
            logging.error(f"❌ Failed to load demo.xlsx: {e}")
            raise

    def _floor_matches(self, floor_condition, floor_no):
        try:
            floor_condition = floor_condition.replace(" ", "")  # 🧹 Remove all spaces
            floor_no_str = f"{floor_no}階"

            if floor_condition == "Allfloors":
                return True
            if floor_condition == "1.2FcaseAll" and floor_no in [1, 2]:
                return True
            if floor_condition == "1.2.3FcaseAll" and floor_no in [1, 2, 3]:
                return True
            if floor_condition.startswith("ONLY"):
                # ✨ Safely match ONLY floor
                expected_floor = floor_condition.replace("ONLY", "")
                return expected_floor == floor_no_str
            return False
        except Exception as e:
            logging.error(f"⚠️ Error matching floor: {floor_condition} - {e}")
            return False

    def apply_conditions(self, destination_sheet, _, builder_id, floor_no):
        try:
            builder_conditions = self.conditions_df[
                self.conditions_df["Builder Code"].astype(str).str.strip() == str(builder_id).strip()
            ]
            # 🚀 Filter floors first
            builder_conditions = builder_conditions[
                builder_conditions["Floor value"].apply(lambda x: self._floor_matches(x, floor_no))
            ]

            logging.info(f"✅ Normal Conditions after floor filter: {builder_conditions.shape[0]} rows")

            for _, row in builder_conditions.iterrows():
                location = row["Location"]
                action = row["What to do"]

                if "first free cell available" in location.lower():
                    continue  # handled separately

                logging.info(f"🔎 Normal Condition: {location} | {action}")

                sheetname, cell = self._parse_location(location)
                sheet = destination_sheet.book.sheets[sheetname.strip()]
                self._apply_action(sheet, destination_sheet, cell, action)
        except Exception as e:
            logging.error(f"❌ Error in apply_conditions: {e}")
            raise

    def apply_first_free_cell_conditions(self, destination_sheet, stock_sheet, builder_id, floor_no):
        try:
            builder_conditions = self.conditions_df[
                self.conditions_df["Builder Code"].astype(str).str.strip() == str(builder_id).strip()
            ]
            logging.info(
                f"Checking First Free Cell Rows for Builder {builder_id}: {builder_conditions.shape[0]} rows (before filtering)"  # noqa
            )

            # 🛡 FILTER ONLY MATCHING FLOOR FIRST
            builder_conditions = builder_conditions[
                builder_conditions["Floor value"].apply(lambda x: self._floor_matches(x, floor_no))
            ]
            logging.info(f"✅ Matching Conditions after floor filter: {builder_conditions.shape[0]} rows")

            for _, row in builder_conditions.iterrows():
                location = row["Location"]
                action = row["What to do"]

                if "first free cell available" not in location.lower():
                    continue

                logging.info(f"🔎 First Free Cell Matched: {location} | {action}")

                sheetname, _ = self._parse_location(location)
                sheetname = sheetname.strip()
                try:
                    sheet_obj = destination_sheet.book.sheets[sheetname]
                except Exception as e:
                    logging.error(f"❌ Sheet '{sheetname}' not found in workbook: {e}")
                    continue

                logging.info(f"📄 Target Sheet: '{sheet_obj.name}' (Excel sheet name from xlwings)")
                logging.info(
                    f"🧪 Checking if '{sheet_obj.name}' == '野縁'? {'YES' if sheet_obj.name == '野縁' else 'NO'}"
                )

                column = self._extract_column_from_location_text(location)
                value = self._extract_value_from_action(action)

                self._put_in_first_free_cell(sheet_obj, sheetname, column, value)

        except Exception as e:
            logging.error(f"❌ Error in apply_first_free_cell_conditions: {e}")
            raise

    def _parse_location(self, location_text):
        try:
            location_text = location_text.replace("、", "")  # fix Japanese comma
            sheetname = ""
            cell = ""

            if "Sheetname:" in location_text:
                sheetname = location_text.split("Sheetname:")[1].split("Range:")[0].strip()
                sheetname = sheetname.strip(" 、")  # strip spaces and Japanese comma

            if "Range:" in location_text:
                cell = location_text.split("Range:")[1].strip()

            return sheetname, cell
        except Exception as e:
            logging.error(f"⚠️ Error parsing location: {location_text} ({e})")
            return "", ""

    def _extract_column_from_cell(self, cell_text):
        try:
            if "in" in cell_text:
                # Example: 'first free cell available in Q column'
                column_part = cell_text.split("in")[-1]
                logging.info(column_part)
                column_letter = "".join(filter(str.isalpha, column_part)).upper()
                logging.info(column_letter)
                return column_letter
            else:
                return "".join(filter(str.isalpha, cell_text))
        except Exception as e:
            logging.error(f"⚠️ Error extracting column from cell: {cell_text} ({e})")
            return "M"  # fallback

    def _extract_value_from_action(self, action_text):
        value = action_text.split("Put")[-1].strip().strip("'")
        return value

    def _extract_column_from_location_text(self, location_text):
        try:
            if "in" in location_text:
                column_part = location_text.split("in")[-1].strip()
                column_letter = "".join(filter(str.isalpha, column_part)).upper()

                # ⚡ Extra: remove COLUMN word
                if column_letter.endswith("COLUMN"):
                    column_letter = column_letter.replace("COLUMN", "")

                logging.info(f"📏 Extracted column for first free cell: {column_letter}")
                return column_letter
        except Exception as e:
            logging.error(f"⚠️ Error extracting column from location text: {location_text} ({e})")
            return "M"  # fallback

    def _put_in_first_free_cell(self, sheet, sheetname, column, value):
        try:
            clean_name = sheetname.replace(" ", "").replace("　", "")

            if clean_name == "野縁":
                start_row = 11
            else:
                start_row = 5

            end_row = 50

            for row in range(start_row, end_row + 1):
                cell_value = sheet.range(f"{column}{row}").value

                # 📢 Important: Empty cell must be REALLY empty (no text, no 0)
                if cell_value is None or str(cell_value).strip() in ["", "0"]:
                    sheet.range(f"{column}{row}").value = int(value) if str(value).isdigit() else value
                    logging.info(f"📝 First Free Cell: Put {value} in {column}{row}")
                    return
            else:
                logging.warning(f"⚠️ No empty cell found in {column} within rows {start_row}-{end_row}")
        except Exception as e:
            logging.error(f"❌ Error during first free cell insertion in {column}: {e}")

    def _apply_action(self, sheet, stock_sheet, cell, action):
        try:
            if action.startswith("Overwrite"):
                value = action.split("'")[1]
                sheet[cell].value = value
                logging.info(f"✏️ Overwrote {cell} with '{value}'")

            elif action.startswith("Put"):
                value = action.split("Put")[-1].strip().strip("'")
                sheet[cell].value = value if not value.isdigit() else int(value)
                logging.info(f"➕ Put {value} in {cell}")

            elif action.startswith("Add"):
                increment = int(action.split("+")[-1])
                sheet[cell].value = (sheet[cell].value or 0) + increment
                logging.info(f"➕ Added {increment} to {cell}")

            elif action.startswith("Shoumeifukku"):
                value = action.split("'")[1]
                sheet[cell].value = value
                logging.info(f"💡 Shoumeifukku '{value}' in {cell}")

            elif action.startswith("Special_Shoumeifukku"):
                部屋番号 = pd.Series(stock_sheet.range("F11:F46").value)
                部屋番号.dropna(inplace=True)
                unique_count = len(set(部屋番号))
                sheet[cell].value = (unique_count * 2) + 2
                logging.info(f"💡 Special_Shoumeifukku set {sheet[cell].value} in {cell}")

            elif action.startswith("Neji"):
                min_neji = (sheet["J19"].value or 0) * 2
                neji = 1
                while 30 * neji <= min_neji:
                    neji += 1
                sheet[cell].value = neji
                logging.info(f"🔨 Neji set {neji} in {cell}")

            elif cell == "C21":
                sheet[cell].value = "吊元セット～【L＝250】"
                logging.info("🌇 Set C21")

            elif cell == "C22":
                sheet[cell].value = "吊元セット～【L＝150】"
                logging.info("🌇 Set C22")

            elif cell == "J12":
                sheet[cell].value = 2
                logging.info("🔈 Set J12=2")

        except Exception as e:
            logging.error(f"⚠️ Error applying direct action on {cell}: {e}")
