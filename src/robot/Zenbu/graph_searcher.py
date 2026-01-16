# === graph_searcher.py ===

import time

import requests
from config import BASE_URL
from Token_Manager import get_access_token


def search_anken_folder(anken_number, retries=3, delay=3):
    """
    Search for the folder corresponding to the anken number.
    """
    search_url = f"{BASE_URL}/search/query"
    payload = {"requests": [{"entityTypes": ["driveItem"], "query": {"queryString": anken_number}, "region": "JPN"}]}
    headers = {"Authorization": f"Bearer {get_access_token()}", "Content-Type": "application/json"}

    for attempt in range(retries):
        try:
            resp = requests.post(search_url, headers=headers, json=payload)
            resp.raise_for_status()

            results = resp.json()
            items = results.get("value", [])[0].get("hitsContainers", [])[0].get("hits", [])

            if not items:
                return None

            first = items[0]
            item = first.get("resource", {})
            return {"name": item.get("name"), "id": item.get("id"), "parentReference": item.get("parentReference", {})}

        except Exception as e:
            print(f"Search attempt {attempt+1} failed: {e}")
            if attempt < retries - 1:
                time.sleep(delay)
            else:
                raise Exception(f"Search failed after {retries} attempts: {e}") from None


def list_children(drive_id, folder_id):
    """
    List all children under a folder, handling pagination if necessary.
    """
    url = f"{BASE_URL}/drives/{drive_id}/items/{folder_id}/children"
    headers = {"Authorization": f"Bearer {get_access_token()}"}

    all_items = []

    while url:
        resp = requests.get(url, headers=headers)
        if resp.status_code != 200:
            raise Exception(f"List children failed: {resp.text}")

        data = resp.json()
        items = data.get("value", [])
        all_items.extend(items)

        url = data.get("@odata.nextLink", None)  # handle pagination

    return all_items
