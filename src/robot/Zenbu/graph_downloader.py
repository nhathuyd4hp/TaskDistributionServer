# --- graph_downloader.py ---

import logging
import os

import requests
from config import BASE_URL
from graph_searcher import list_children, search_anken_folder
from Token_Manager import get_access_token


def download_pdf(download_url, save_path):
    """
    Download a single file from its download URL.
    """
    try:
        resp = requests.get(download_url, stream=True)
        resp.raise_for_status()
        os.makedirs(os.path.dirname(save_path), exist_ok=True)
        with open(save_path, "wb") as f:
            for chunk in resp.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)
        logging.info(f"✅ Saved {os.path.basename(save_path)} successfully.")
    except Exception as e:
        logging.error(f"❌ Failed to download {os.path.basename(save_path)}: {e}")


def download_files_inside_folder(drive_id, folder_id, local_folder_path):
    """
    Download only PDF and Excel files DIRECTLY inside 割付図 (no recursion into subfolders).
    """
    os.makedirs(local_folder_path, exist_ok=True)
    allowed_extensions = (".pdf", ".xls")  # Extend if needed
    downloaded_count = 0  # ✅ Track count

    url = f"{BASE_URL}/drives/{drive_id}/items/{folder_id}/children"
    headers = {"Authorization": f"Bearer {get_access_token()}"}

    while url:
        resp = requests.get(url, headers=headers)
        resp.raise_for_status()
        data = resp.json()

        for item in data.get("value", []):
            # 📌 Only files (not folders)
            if "folder" not in item:
                file_name = item["name"]
                if file_name.lower().endswith(allowed_extensions):
                    download_url = item["@microsoft.graph.downloadUrl"]
                    save_path = os.path.join(local_folder_path, file_name)
                    download_pdf(download_url, save_path)
                    downloaded_count += 1  # ✅ Increment
                else:
                    logging.info(f"⏩ Skipped non-target file: {file_name}")
            else:
                logging.info(f"📂 Skipped subfolder: {item['name']}")

        url = data.get("@odata.nextLink", None)

    if downloaded_count == 0:
        raise Exception("❌ 割付図 folder found but no files downloaded!")


def download_folder_by_anken(anken_number, local_folder_path):
    """
    Handles outliers where the search result itself is 割付図 folder.
    """
    anken_info = search_anken_folder(anken_number)
    if not anken_info:
        raise Exception(f"❌ No Anken folder found for: {anken_number}")

    drive_id = anken_info["parentReference"]["driveId"]
    folder_id = anken_info["id"]
    folder_name = anken_info["name"]

    # Case A: Anomaly — the search result IS 割付図 folder
    if "割付図" in folder_name or "割付図・エクセル" in folder_name:
        logging.warning(f"⚠️ Anomaly: Got 割付図 folder directly from search result: {folder_name}")
        download_files_inside_folder(drive_id, folder_id, local_folder_path)
        return

    # Case B: Normal — search inside children for 割付図
    children = list_children(drive_id, folder_id)
    target_folder_id = None

    for item in children:
        child_name = item.get("name", "")
        if "割付図" in child_name or "割付図・エクセル" in child_name:
            target_folder_id = item["id"]
            logging.info(f"📂 Found 割付図 in subfolders: {child_name}")
            break

    if not target_folder_id:
        logging.warning(f"🚫 No folder containing '割付図' found under {anken_number}")
        logging.warning(f"📝 Subfolders found: {[item.get('name') for item in children]}")
        raise Exception(f"❌ No folder containing '割付図' found under {anken_number}")

    download_files_inside_folder(drive_id, target_folder_id, local_folder_path)
    logging.info(f"✅ Downloaded 割付図 successfully for {anken_number}")
