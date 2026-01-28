# === graph_downloader.py ===

import logging
import os

import requests
from token_manager import get_access_token

# === Constants ===
BASE_URL = "https://graph.microsoft.com/v1.0"

# === Functions ===


def search_anken_folder(anken_number):
    """Search for folder by anken_number (案件番号) globally."""
    search_url = f"{BASE_URL}/search/query"
    payload = {"requests": [{"entityTypes": ["driveItem"], "query": {"queryString": anken_number}, "region": "JPN"}]}
    headers = {"Authorization": f"Bearer {get_access_token()}", "Content-Type": "application/json"}

    resp = requests.post(search_url, headers=headers, json=payload)
    if resp.status_code != 200:
        raise Exception(f"Search failed: {resp.text}")

    results = resp.json()
    items = results.get("value", [])[0].get("hitsContainers", [])[0].get("hits", [])

    if not items:
        return None

    resource = items[0].get("resource", {})
    return {
        "id": resource.get("id"),
        "parentReference": resource.get("parentReference", {}),
        "name": resource.get("name"),
    }


def list_children(drive_id, folder_id):
    """List children under a folder."""
    url = f"{BASE_URL}/drives/{drive_id}/items/{folder_id}/children"
    headers = {"Authorization": f"Bearer {get_access_token()}"}
    resp = requests.get(url, headers=headers)
    if resp.status_code != 200:
        raise Exception(f"List children failed: {resp.text}")
    return resp.json().get("value", [])


def download_pdf(file_url, save_path):
    """Download a PDF file."""
    resp = requests.get(file_url)
    if resp.status_code == 200:
        with open(save_path, "wb") as f:
            f.write(resp.content)
        logging.info(f"✅ Downloaded: {os.path.basename(save_path)}")


def graph_download_and_save_files(anken_number, factory_folder, builder_name, anken_name, 納期):
    """Main function to download PDFs."""
    try:
        folder_info = search_anken_folder(anken_number)
        if not folder_info:
            logging.warning(f"⚠️ No folder found for ankenbango: {anken_number}")
            return False

        drive_id = folder_info["parentReference"]["driveId"]
        folder_id = folder_info["id"]

        children = list_children(drive_id, folder_id)

        # 📂 Save path setup
        safe_builder = builder_name.replace("/", "_").replace("\\", "_").strip()
        safe_anken = anken_name.replace("/", "_").replace("\\", "_").strip()

        local_folder = os.path.join(factory_folder, safe_anken)
        os.makedirs(local_folder, exist_ok=True)

        for child in children:
            if "割付図" in child["name"] or "割付図・エクセル" in child["name"] or "見積" in child["name"]:
                subfolder_id = child["id"]
                subfolder_name = child["name"]
                subfolder_files = list_children(drive_id, subfolder_id)

                for file in subfolder_files:
                    if file["name"].lower().endswith(".pdf"):
                        download_url = file["@microsoft.graph.downloadUrl"]

                        # Different filename depending on the folder
                        if "割付図" in subfolder_name:
                            final_filename = f"★軽天割付図面-{safe_builder}-{safe_anken}.pdf"
                        elif "見積" in subfolder_name:
                            final_filename = f"★御見積書-{safe_builder}-{safe_anken}.pdf"
                        else:
                            final_filename = file["name"]  # fallback

                        final_path = os.path.join(local_folder, final_filename)
                        download_pdf(download_url, final_path)
                        found_pdf = True

        if found_pdf:
            return True
        else:
            logging.warning(f"⚠️ No PDF files found under {anken_number}")
            return False

    except Exception as e:
        logging.error(f"❌ Graph API download failed: {e}")
        return False
