import logging
import os

import requests
from config import BASE_URL
from Token_Manager import get_access_token


class Builder_SharePoint_GraphAPI:
    def __init__(self, builder_file_name, local_folder_path):
        self.builder_file_name = builder_file_name
        self.local_folder_path = local_folder_path
        self.download_and_save_file()

    def search_builder_file(self):
        """
        Search the builder file by name and return driveId and id.
        """
        headers = {"Authorization": f"Bearer {get_access_token()}", "Content-Type": "application/json"}

        search_url = f"{BASE_URL}/search/query"
        payload = {
            "requests": [
                {"entityTypes": ["driveItem"], "query": {"queryString": self.builder_file_name}, "region": "JPN"}
            ]
        }

        resp = requests.post(search_url, headers=headers, json=payload)
        if resp.status_code != 200:
            raise Exception(f"Graph API: Search failed: {resp.text}")

        results = resp.json()
        items = results.get("value", [])[0].get("hitsContainers", [])[0].get("hits", [])

        if not items:
            raise Exception(f"❌ No matching file found for {self.builder_file_name}")

        file_info = items[0].get("resource", {})
        drive_id = file_info["parentReference"]["driveId"]
        file_id = file_info["id"]

        return drive_id, file_id

    def get_download_url(self, drive_id, file_id):
        """
        After search, fetch full metadata to get download URL.
        """
        headers = {"Authorization": f"Bearer {get_access_token()}"}
        file_url = f"{BASE_URL}/drives/{drive_id}/items/{file_id}"
        resp = requests.get(file_url, headers=headers)

        if resp.status_code != 200:
            raise Exception(f"Graph API: Failed to get file metadata: {resp.text}")

        file_metadata = resp.json()
        return file_metadata["@microsoft.graph.downloadUrl"]

    def download_file(self, download_url):
        resp = requests.get(download_url, stream=True)
        if resp.status_code != 200:
            raise Exception(f"Failed to download file content: {resp.text}")
        return resp.content

    def save_file(self, file_bytes):
        os.makedirs(self.local_folder_path, exist_ok=True)
        save_path = os.path.join(self.local_folder_path, self.builder_file_name)
        with open(save_path, "wb") as f:
            f.write(file_bytes)
        logging.info(f"✅ Builder file saved: {save_path}")

    def download_and_save_file(self):
        """
        Complete flow: search → fetch metadata → download → save
        """
        try:
            drive_id, file_id = self.search_builder_file()
            download_url = self.get_download_url(drive_id, file_id)
            file_bytes = self.download_file(download_url)
            self.save_file(file_bytes)
        except Exception as e:
            logging.error(f"❌ Error downloading builder file {self.builder_file_name}: {e}")
            raise
