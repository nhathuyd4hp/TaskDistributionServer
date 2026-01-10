# config.py
import logging
import os

from Nasiwak import create_json_config

from src.core.config import settings

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
    logging.info("Configs loaded successfully.")
except Exception as e:
    logging.error(f"Failed to load configs: {e}")

# 📚 Microsoft Graph API Credentials
CLIENT_ID = settings.API_SHAREPOINT_CLIENT_ID
CLIENT_SECRET = settings.API_SHAREPOINT_CLIENT_SECRET
TENANT_ID = settings.API_SHAREPOINT_TENANT_ID

# 🌎 Microsoft Graph API URLs
BASE_URL = "https://graph.microsoft.com/v1.0"
SEARCH_URL = f"{BASE_URL}/search/query"

# 📂 Drives to Search (Multiple Sites)
DRIVE_IDS = [
    # Site: 2021
    "b!XKuZyeFTlkSp4cWfSxd10AfxU7PGA3xBi27uhOfFFMgqFmxUFSTDS5xlMmIATcY_",  # か行
    "b!XKuZyeFTlkSp4cWfSxd10AfxU7PGA3xBi27uhOfFFMhUTAdjY1jnSK7YfX-IfcQs",  # さ行
    "b!XKuZyeFTlkSp4cWfSxd10AfxU7PGA3xBi27uhOfFFMgxXBpRecuxQp40qCQ96qCw",  # た行
    "b!XKuZyeFTlkSp4cWfSxd10AfxU7PGA3xBi27uhOfFFMhn9jPDMigKTKcQq4biVQTp",  # ドキュメント・2
    "b!XKuZyeFTlkSp4cWfSxd10AfxU7PGA3xBi27uhOfFFMgLltFBcoLeSJyVrkeRVc-u",  # な行
    "b!XKuZyeFTlkSp4cWfSxd10AfxU7PGA3xBi27uhOfFFMiGv-OjvRfuSIZjru4KPfrt",  # ドキュメント
    "b!XKuZyeFTlkSp4cWfSxd10AfxU7PGA3xBi27uhOfFFMje-cYmil_oQ6oMx_OlS8au",  # ま行
    "b!XKuZyeFTlkSp4cWfSxd10AfxU7PGA3xBi27uhOfFFMi2WzsMrIhERLjKzloPS0YK",  # や・ら・わ行
    "b!XKuZyeFTlkSp4cWfSxd10AfxU7PGA3xBi27uhOfFFMihKuZWYmqkTqqy3R9t3aff",  # は行
    "b!XKuZyeFTlkSp4cWfSxd10AfxU7PGA3xBi27uhOfFFMhYgcNNyw5IQKep4L6_VFIk",  # あ行
    # Site: Kantou
    "b!CGMwpFZqO0aR13-uULpoA739OTZDETFKpDsa-PGqFCBe0TiC03OyTLyZUjcaE8e9",  # ドキュメント
    "b!CGMwpFZqO0aR13-uULpoA739OTZDETFKpDsa-PGqFCBMszIEG92nQ76ejmAOfnzy",  # 新ドキュメント(関東)
    "b!CGMwpFZqO0aR13-uULpoA739OTZDETFKpDsa-PGqFCCdvvkNEUAUTb8Gxjm9oin3",  # 検索構成リスト
    # Site: 2019
    "b!sCgCnWR2UkGKdRInfBWzdlcnAGNMtfdEjamzCOTJHvCO1eFDmXWzRpY7g3QpUVA-",  # Documents
    "b!sCgCnWR2UkGKdRInfBWzdlcnAGNMtfdEjamzCOTJHvANeCwNSd0wTZ7-9-ersYK5",  # タマホーム
    # Site: Shuuko
    "b!vArDktlKE0uGKwPHe6i71cHlFfas-b9DhL0W0_9h3SLpFob0RyrQRrPmZvYxcvot",  # Documents
    "b!vArDktlKE0uGKwPHe6i71cHlFfas-b9DhL0W0_9h3SLNUAN_SP5FRZzzVzIygXm8",  # Search Config List
]

# 📁 Base Folder for Downloads
DOWNLOAD_DIR = os.path.join(os.getcwd(), "Ankens")

# 📄 Excel Path for Ankens
EXCEL_PATH = os.path.join(os.getcwd(), "anken_list.xlsx")

# ✏️ Column Names
ANKEN_COLUMN = "案件番号"
STATUS_COLUMN = "Download Status"

# ✨ Folder Paths
BASE_DIR = os.getcwd()
CSV_INPUT_FOLDER = os.path.join(BASE_DIR, "CSV")  # Folder where CSVs are stored
EXCEL_OUTPUT_FOLDER = os.path.join(BASE_DIR, "Excels")  # Folder where Excel files are saved
DOWNLOAD_DIR = os.path.join(BASE_DIR, "Ankens")  # Folder where PDFs are downloaded

# ✨ Other configs
REGION = "JPN"  # For Microsoft Graph Search API (Japan site)

# Download Config
BATCH_SIZE = 5  # Only 5 ankens in parallel to avoid 429
MAX_RETRIES = 3  # Retry 3 times on 429
RETRY_SLEEP = 5  # 5 seconds sleep on 429
GRAPH_SCOPE = ["https://graph.microsoft.com/.default"]
