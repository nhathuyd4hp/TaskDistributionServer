import logging
import os
import shutil
import sys
import time
from pathlib import Path

import requests
from dotenv import load_dotenv
from msal import ConfidentialClientApplication

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s  - %(levelname)s - %(message)s",
    handlers=[logging.FileHandler(f"{'Access_token_log'}"), logging.StreamHandler()],
)
logging.info("Access_token_log file created")

# --- Microsoft Graph API Credentials ---
current_dir = Path(__file__).resolve().parent
env_path = current_dir.parent.parent.parent / ".env"
load_dotenv(dotenv_path=env_path)
CLIENT_ID = os.getenv("API_SHAREPOINT_CLIENT_ID")
CLIENT_SECRET = os.getenv("API_SHAREPOINT_CLIENT_SECRET")
TENANT_ID = os.getenv("API_SHAREPOINT_TENANT_ID")

BASE_URL = "https://graph.microsoft.com/v1.0"
GRAPH_SCOPE = ["https://graph.microsoft.com/.default"]

_token_cache = {"access_token": None, "expires_at": 0}


def get_access_token():
    now = time.time()
    if _token_cache["access_token"] and now < _token_cache["expires_at"] - 60:
        return _token_cache["access_token"]

    authority = f"https://login.microsoftonline.com/{TENANT_ID}"
    app = ConfidentialClientApplication(client_id=CLIENT_ID, client_credential=CLIENT_SECRET, authority=authority)
    result = app.acquire_token_for_client(scopes=GRAPH_SCOPE)

    if "access_token" not in result:
        raise Exception(f"Failed to get token: {result.get('error_description')}")

    _token_cache["access_token"] = result["access_token"]
    _token_cache["expires_at"] = now + result["expires_in"]
    return _token_cache["access_token"]


def get_site_id():
    # Adjust the site URL to your SharePoint site domain and name
    url = "https://graph.microsoft.com/v1.0/sites/nskkogyo.sharepoint.com:/sites/DQ"
    headers = {"Authorization": f"Bearer {get_access_token()}"}
    resp = requests.get(url, headers=headers)
    if resp.status_code != 200:
        raise Exception(f"Failed to get site ID: {resp.text}")
    return resp.json()["id"]


def get_drive_id(site_id):
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
    headers = {"Authorization": f"Bearer {get_access_token()}"}
    resp = requests.get(url, headers=headers)
    if resp.status_code != 200:
        raise Exception(f"Failed to get drive ID: {resp.text}")

    for drive in resp.json()["value"]:
        if drive["name"] == "Shared Documents" or drive["name"] == "ドキュメント":
            return drive["id"]

    raise Exception("Shared Documents drive not found.")


def get_folder_contents(site_id, drive_id, folder_path):
    # URL encode the folder path for Microsoft Graph API
    from requests.utils import quote

    encoded_path = quote(folder_path, safe="/")

    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{encoded_path}:/children"
    headers = {"Authorization": f"Bearer {get_access_token()}", "Content-Type": "application/json"}
    resp = requests.get(url, headers=headers)
    if resp.status_code != 200:
        raise Exception(f"Failed to get folder contents: {resp.text}")
    return resp.json().get("value", [])


# Example usage:

site_id = get_site_id()
logging.info(f"site_id is: {site_id}")
drive_id = get_drive_id(site_id)
logging.info(f"get_drive_id is: {drive_id}")

# The target folder path exactly as in your SharePoint URL (decoded):
target_folder_path = "インド事務所関係/Access_token"


def download_file(drive_id, item_id, save_path):
    url = f"{BASE_URL}/drives/{drive_id}/items/{item_id}/content"
    headers = {"Authorization": f"Bearer {get_access_token()}"}
    resp = requests.get(url, headers=headers, stream=True)
    if resp.status_code != 200:
        raise Exception(f"Failed to download file: {resp.text}")

    os.makedirs(os.path.dirname(save_path), exist_ok=True)
    with open(save_path, "wb") as f:
        for chunk in resp.iter_content(chunk_size=8192):
            if chunk:
                f.write(chunk)
    logging.info(f"Downloaded file to: {save_path}")


items = get_folder_contents(site_id, drive_id, target_folder_path)
logging.info(f"Contents of target folder: {items}")
for item in items:
    kind = "Folder" if "folder" in item else "File"
    logging.info(f"{kind}: {item['name']} - URL: {item['webUrl']}")

if not os.path.exists("Access_token"):
    os.makedirs("Access_token")
    logging.info("Access_token Folder Created")
else:
    logging.info("Access_token Folder Already Exists")
    try:
        for f in os.listdir("Access_token"):
            fpath = os.path.join("Access_token", f)
            if os.path.isfile(fpath) or os.path.islink(fpath):
                os.remove(fpath)
            elif os.path.isdir(fpath):
                shutil.rmtree(fpath)
            logging.info(f"Removed: {fpath}")
    except Exception as e:
        logging.info(f"Error in Exception(Access_token folder):\n{e}")
        sys.exit()

# Find the file named "Template_File.xlsm"
token_file = next((item for item in items if item["name"] == "Access_token.txt"), None)

if token_file:
    save_path = os.path.join("Access_token", token_file["name"])
    down_file = download_file(drive_id, token_file["id"], save_path)
    logging.info("token file downloaded successfully")
else:
    logging.info("token file not found in the folder.")
    sys.exit()
