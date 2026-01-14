import os
from pathlib import Path

from dotenv import load_dotenv

# === config.py ===

# 📚 Microsoft Graph API Credentials
current_dir = Path(__file__).resolve().parent
env_path = current_dir.parent.parent.parent / ".env"
load_dotenv(dotenv_path=env_path)
CLIENT_ID = os.getenv("API_SHAREPOINT_CLIENT_ID")
CLIENT_SECRET = os.getenv("API_SHAREPOINT_CLIENT_SECRET")
TENANT_ID = os.getenv("API_SHAREPOINT_TENANT_ID")

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
EXCEL_PATH = os.path.join(os.getcwd(), "Data.xlsx")

# ✏️ Column Names
ANKEN_COLUMN = "案件番号"
STATUS_COLUMN = "Download Status"

# ✨ Folder Paths
BASE_DIR = os.getcwd()
CSV_INPUT_FOLDER = os.path.join(BASE_DIR, "CSV")          # Folder where CSVs are stored
EXCEL_OUTPUT_FOLDER = os.path.join(BASE_DIR, "Excels")     # Folder where Excel files are saved
DOWNLOAD_DIR = os.path.join(BASE_DIR, "Ankens")            # Folder where PDFs are downloaded

# ✨ Other configs
REGION = "JPN"  # For Microsoft Graph Search API (Japan site)

# Download Config
BATCH_SIZE = 5  # Only 5 ankens in parallel to avoid 429
MAX_RETRIES = 3  # Retry 3 times on 429
RETRY_SLEEP = 5  # 5 seconds sleep on 429
GRAPH_SCOPE = ["https://graph.microsoft.com/.default"]