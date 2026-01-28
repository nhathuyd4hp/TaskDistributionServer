import datetime
import logging
import os
import time
from pathlib import Path

import pandas as pd
import requests
from dotenv import load_dotenv
from msal import ConfidentialClientApplication

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

# ✨ Other configs
REGION = "JPN"  # For Microsoft Graph Search API (Japan site)

# Download Config
BATCH_SIZE = 5  # Only 5 ankens in parallel to avoid 429
MAX_RETRIES = 3  # Retry 3 times on 429
RETRY_SLEEP = 5  # 5 seconds sleep on 429
GRAPH_SCOPE = ["https://graph.microsoft.com/.default"]


################# BELOW IS TOKEN MANAGAGER PART ################################


_token_cache = {"access_token": None, "expires_at": 0}


def get_access_token():
    now = time.time()  # gives current time in seconds
    if _token_cache["access_token"] and now < _token_cache["expires_at"] - 60:
        return _token_cache["access_token"]

    authority = f"https://login.microsoftonline.com/{TENANT_ID}"
    app = ConfidentialClientApplication(client_id=CLIENT_ID, client_credential=CLIENT_SECRET, authority=authority)
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])

    if "access_token" not in result:
        raise Exception(f"Failed to get token: {result.get('error_description')}")

    _token_cache["access_token"] = result["access_token"]
    _token_cache["expires_at"] = now + result["expires_in"]
    return _token_cache["access_token"]


######################## BELOW is template download part ########################
def template_down(folder_365, download_dir):
    logging.info("Opening sharepoint to download template")

    file_info = search_anken_folder(folder_365)
    if not file_info:
        logging.warning(f"No file found for {folder_365}")
        return False, False
    logging.info(f"file info is: {file_info}")

    drive_id = file_info["parentReference"]["driveId"]
    file_id = file_info["id"]
    file_name = file_info["name"]
    headers = {"Authorization": f"Bearer {get_access_token()}"}

    # 1. Get full metadata including download URL
    url = f"{BASE_URL}/drives/{drive_id}/items/{file_id}"
    resp = requests.get(url, headers=headers)
    if resp.status_code != 200:
        logging.error(f"Failed to get file metadata: {resp.text}")
        return False, False

    file_metadata = resp.json()
    download_url = file_metadata.get("@microsoft.graph.downloadUrl")
    if not download_url:
        logging.error("Download URL not found in file metadata")
        return False, False

    # 2. Download the file
    # save_dir = "downloads"  # Or any other path you want
    os.makedirs(download_dir, exist_ok=True)
    save_path = os.path.join(download_dir, file_name)

    r = requests.get(download_url)
    with open(save_path, "wb") as f:
        f.write(r.content)

    logging.info(f"Downloaded {file_name} to {save_path}")
    return file_name, True


################# BELOW IS Graph_searcher PART ################################


# from config import get_access_token
# from config import BASE_URL


def search_anken_folder(anken_number, target_names=None):
    """
    Correct logic:
    1) First find the MAIN 案件 folder by に in folder name.
    2) Only THEN find target subfolder inside it (children API).
    """
    query = str(anken_number).strip()

    # --- Step 1: SEARCH for the parent folder only ---
    search_url = f"{BASE_URL}/search/query"
    payload = {"requests": [{"entityTypes": ["driveItem"], "query": {"queryString": query}, "region": "JPN"}]}
    headers = {"Authorization": f"Bearer {get_access_token()}", "Content-Type": "application/json"}

    resp = requests.post(search_url, headers=headers, json=payload)
    if resp.status_code != 200:
        raise Exception(f"Search failed: {resp.text}")

    results = resp.json()
    items = results.get("value", [])[0].get("hitsContainers", [])[0].get("hits", [])

    if not items:
        logging.warning(f"No SP results for {anken_number}")
        return None

    logging.info(f"Items found: {len(items)}")

    # --- Step 1A: find the correct parent folder ---
    parent_folder = None
    for hit in items:
        resource = hit.get("resource", {})
        name = resource.get("name", "")
        if normalize(query) in normalize(name):
            parent_folder = resource
            logging.info(f"Matched parent ANKEN folder: {name}")
            break

    if not parent_folder:
        logging.warning(f"No parent folder matched for {anken_number}")
        return None

    # If no target subfolder needed, return parent folder
    if not target_names:
        return parent_folder

    # --- Step 2: search INSIDE the parent folder for target subfolder ---
    parent_id = parent_folder["id"]
    parent_drive = parent_folder["parentReference"]["driveId"]

    children = list_children(parent_drive, parent_id)

    # normalize target names
    if isinstance(target_names, str):
        target_names = [target_names]

    for child in children:
        cname = child.get("name", "")
        if cname in target_names:
            logging.info(f"Matched target subfolder: {cname}")
            return {"name": cname, "id": child.get("id"), "parentReference": {"driveId": parent_drive, "id": parent_id}}

    logging.warning(f"No matching subfolder in ANKEN {anken_number}")
    return None


def list_children(drive_id, folder_id):
    url = f"{BASE_URL}/drives/{drive_id}/items/{folder_id}/children"
    headers = {"Authorization": f"Bearer {get_access_token()}"}
    resp = requests.get(url, headers=headers)
    if resp.status_code != 200:
        raise Exception(f"List children failed: {resp.text}")
    return resp.json().get("value", [])


####################### BELOW Part is for Downloading #######################


def normalize(text):
    if not text:
        return ""
    return (
        str(text)
        .replace(" ", "")  # remove half-width spaces
        .replace("　", "")  # remove full-width spaces
        .replace("(", "（")  # normalize parentheses
        .replace(")", "）")
        .replace("[", "")
        .replace("]", "")
        .replace("【", "")
        .replace("】", "")
        .strip()
    )


def Download_folder(anken_number, download_dir):
    try:
        target_names = ["割付図・エクセル", "割付図。エクセル", "割付図・ エクセル"]
        folder_info = search_anken_folder(anken_number, target_names)
        logging.info(f"Folder info is: {folder_info}")
        # input('a')
        if not folder_info:
            logging.warning(f"No folder found for {anken_number}")
            return False
        drive_id = folder_info["parentReference"]["driveId"]
        logging.info(f"Drive ID is: {drive_id}")
        folder_id = folder_info["id"]
        logging.info(f"Folder ID is: {folder_id}")
        children = list_children(drive_id, folder_id)
        # logging.info(f"Children in folder: {children}")

        save_dir = os.path.join(download_dir)
        os.makedirs(save_dir, exist_ok=True)

        for file in children:
            if file["name"].lower().endswith(".pdf"):
                pdf_name = file["name"]
                file_url = file["@microsoft.graph.downloadUrl"]
                save_path = os.path.join(save_dir, pdf_name)
                r = requests.get(file_url)
                with open(save_path, "wb") as f:
                    f.write(r.content)
                logging.info(f"Downloaded {pdf_name}")
                return save_path

        logging.warning(f"No PDF found in 割付図・エクセル for {anken_number}")
        return False

    except Exception as e:
        logging.error(f"Download error for {anken_number}: {e}")
        return False


####################### BELOW Part is for koushin, check and Upload #######################
def upload_files_and_folder(anken_number, local_path, excellinenumber):
    try:
        target_names = "見積書"
        folder_info = search_anken_folder(anken_number, target_names)
        print(f"Folder info is: {folder_info}")

        # if not folder_info:
        #     logging.warning(f"No folder found for {anken_number}")
        #     return "案件フォルダ無し", excellinenumber, False

        # Get list of local files
        if os.path.isdir(local_path):
            local_files = [
                os.path.join(local_path, f)
                for f in os.listdir(local_path)
                if os.path.isfile(os.path.join(local_path, f))
            ]
        elif os.path.isfile(local_path):
            local_files = [local_path]
        else:
            logging.error(f"Invalid path: {local_path}")
            return "見積Pathエラー", excellinenumber, False

        # If 見積書 folder is found, upload into it
        if folder_info:
            見積書_f = "有"
            drive_id = folder_info["parentReference"]["driveId"]
            folder_id = folder_info["id"]
            logging.info(f"Uploading into existing 見積書 folder (ID: {folder_id})")

        else:
            見積書_f = "無"
            # 見積書 folder not found — fallback to案件 folder
            logging.info("見積書 folder not found. Searching for 案件 folder to create 見積書 inside...")

            anken_folder_info = search_anken_folder(anken_number)  # No target_names
            if not anken_folder_info:
                logging.warning(f"No 案件フォルダ found for {anken_number}")
                return "案件フォルダ無し", excellinenumber, False

            drive_id = anken_folder_info["parentReference"]["driveId"]
            folder_id = anken_folder_info["id"]

            # Create 見積書 folder under案件 folder
            create_folder_url = f"{BASE_URL}/drives/{drive_id}/items/{folder_id}/children"
            headers = {"Authorization": f"Bearer {get_access_token()}", "Content-Type": "application/json"}
            data = {"name": "見積書", "folder": {}, "@microsoft.graph.conflictBehavior": "rename"}
            resp = requests.post(create_folder_url, headers=headers, json=data)

            if resp.status_code not in (200, 201):
                logging.error(f"Failed to create 見積書 folder: {resp.status_code} {resp.text}")
                return "見積書フォルダー作成エラー", excellinenumber, "無"

            folder_id = resp.json()["id"]
            logging.info(f"Created 見積書 folder with ID: {folder_id}")

        # ✅ Upload all files into the determined 見積書 folder
        for file_path in local_files:
            file_name_original = os.path.basename(file_path)
            name, ext = os.path.splitext(file_name_original)
            if 見積書_f == "有":
                file_name = name + "_新" + ext  # Add "_新" suffix
            else:
                file_name = name + ext

            upload_url = f"{BASE_URL}/drives/{drive_id}/items/{folder_id}:/{file_name}:/content"
            headers = {"Authorization": f"Bearer {get_access_token()}", "Content-Type": "application/octet-stream"}

            with open(file_path, "rb") as file_data:
                upload_resp = requests.put(upload_url, headers=headers, data=file_data)

            if upload_resp.status_code in (200, 201):
                logging.info(f"Uploaded {file_name} successfully.")
            else:
                logging.error(f"Upload failed for {file_name}: {upload_resp.status_code} {upload_resp.text}")
                return "見積ファイルUPエラー", excellinenumber, "無"

        return True, excellinenumber, "有" if folder_info else "無"

    except Exception as e:
        logging.error(f"Upload error for {anken_number}: {e}")
        return "見積書UPエラー", excellinenumber, False


# def upload_files_and_folder(anken_number, local_path, excellinenumber):
#     try:
#         target_names = "見積書"
#         folder_info = search_anken_folder(anken_number, target_names)
#         print(f"Folder info is: {folder_info}")
#         if not folder_info:
#             logging.warning(f"No folder found for {anken_number}")
#             return "案件フォルダ無し", excellinenumber, False

#         drive_id = folder_info["parentReference"]["driveId"]
#         folder_id = folder_info["id"]
#         children = list_children(drive_id, folder_id)
#         print(f"Children in folder: {children}")
#         # input('a')

#         # Get list of local files
#         if os.path.isdir(local_path):
#             local_files = [
#                 os.path.join(local_path, f)
#                 for f in os.listdir(local_path)
#                 if os.path.isfile(os.path.join(local_path, f))
#             ]
#         elif os.path.isfile(local_path):
#             local_files = [local_path]
#         else:
#             logging.error(f"Invalid path: {local_path}")
#             return "見積Pathエラー", excellinenumber, False

#         for child in children:
#             if "見積" in child["name"]:
#                 logging.info(f"見積書 folder found, uploading files inside 見積書...")
#                 target_folder_id = child["id"]

#                 for file_path in local_files:
#                     file_name_original = os.path.basename(file_path)
#                     name, ext = os.path.splitext(file_name_original)
#                     file_name = name + "_新" + ext  # Add _新 at end
#                     logging.info(f" file being uploaded is: {file_name}")

#                     # Upload file into existing 見積書 folder
#                     upload_url = f"{BASE_URL}/drives/{drive_id}/items/{target_folder_id}:/{file_name}:/content"
#                     headers = {
#                         "Authorization": f"Bearer {get_access_token()}",
#                         "Content-Type": "application/octet-stream"
#                     }

#                     with open(file_path, "rb") as file_data:
#                         resp = requests.put(upload_url, headers=headers, data=file_data)

#                     if resp.status_code in (200, 201):
#                         logging.info(f"Uploaded {file_name} successfully.")
#                     else:
#                         logging.error(f"Upload failed for {file_name}: {resp.status_code} {resp.text}")
#                         return "見積ファイルUPエラー1", excellinenumber

#                 return True, excellinenumber, "有"

#         # 見積書 folder not found, create and upload
#         logging.info(f"見積書 folder not found, creating 見積書 folder and uploading files...")

#         # Create 見積書 folder under the案件 folder
#         create_folder_url = f"{BASE_URL}/drives/{drive_id}/items/{folder_id}/children"
#         headers = {
#             "Authorization": f"Bearer {get_access_token()}",
#             "Content-Type": "application/json"
#             }
#         data = {
#             "name": "見積書", "folder": {},
#             "@microsoft.graph.conflictBehavior": "rename"
#         }
#         resp = requests.post(create_folder_url, headers=headers, json=data)

#         if resp.status_code in (200, 201):
#             new_folder_id = resp.json()["id"]
#             logging.info(f"Created 見積書 folder successfully with id: {new_folder_id}")

#             # Upload each file inside local 見積書 folder
#             for file_path in local_files:
#                 file_name = os.path.basename(file_path)
#                 upload_url = f"{BASE_URL}/drives/{drive_id}/items/{new_folder_id}:/{file_name}:/content"
#                 with open(file_path, "rb") as file_data:
#                     upload_resp = requests.put(upload_url, headers=headers, data=file_data)

#                 if upload_resp.status_code in (200, 201):
#                     logging.info(f"Uploaded {file_name} successfully to new 見積書 folder.")
#                 else:
#                     logging.error(f"Failed to upload {file_name}: {upload_resp.status_code} {upload_resp.text}")
#                     return "見積ファイルUPエラー2", excellinenumber, '無'
#             return True, excellinenumber, '無'
#         else:
#             logging.error(f"Failed to create 見積書 folder: {resp.status_code} {resp.text}")
#             return "見積書フォルダー作成エラー", excellinenumber, '無'

#     except Exception as e:
#         logging.error(f"Upload error for {anken_number}: {e}")
#         return "見積書UPエラー", excellinenumber, False


################# BELOW IS Utilis PART ################################


_progress_data = []


def clear_report():
    global _progress_data
    _progress_data = []


def add_result(anken_number, factory, status, file_name=None):
    _progress_data.append(
        {"案件番号": anken_number, "工場": factory, "ダウンロード状況": status, "ファイル名": file_name or ""}
    )


def save_report():
    if not _progress_data:
        return
    df = pd.DataFrame(_progress_data)
    downloaded = df[df["ダウンロード状況"] == "Success"]
    failed = df[df["ダウンロード状況"] == "Failed"]
    reports_folder = os.path.join(os.getcwd(), "ProgressReports")
    os.makedirs(reports_folder, exist_ok=True)
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M")
    file_path = os.path.join(reports_folder, f"Progress_Report_{timestamp}.xlsx")
    with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
        downloaded.to_excel(writer, sheet_name="Downloaded", index=False)
        failed.to_excel(writer, sheet_name="Failed", index=False)
    logging.info(f"Progress Report Saved at {file_path}")


def get_report():
    return _progress_data
