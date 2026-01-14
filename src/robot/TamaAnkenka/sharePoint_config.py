import datetime
import logging
import os
import time
import requests
from msal import ConfidentialClientApplication
from dotenv import load_dotenv
from pathlib import Path
# === Microsoft Graph API Credentials ===
current_dir = Path(__file__).resolve().parent
env_path = current_dir.parent.parent.parent / ".env"
load_dotenv(dotenv_path=env_path)
CLIENT_ID = os.getenv("API_SHAREPOINT_CLIENT_ID")
CLIENT_SECRET = os.getenv("API_SHAREPOINT_CLIENT_SECRET")
TENANT_ID = os.getenv("API_SHAREPOINT_TENANT_ID")

BASE_URL = "https://graph.microsoft.com/v1.0"
GRAPH_SCOPE = ["https://graph.microsoft.com/.default"]

# === Token cache ===
_token_cache = {"access_token": None, "expires_at": 0}

def get_access_token():
    now = time.time()
    if _token_cache["access_token"] and now < _token_cache["expires_at"] - 60:
        return _token_cache["access_token"]
    
    authority = f"https://login.microsoftonline.com/{TENANT_ID}"
    app = ConfidentialClientApplication(
        client_id=CLIENT_ID,
        client_credential=CLIENT_SECRET,
        authority=authority
    )
    result = app.acquire_token_for_client(scopes=GRAPH_SCOPE)
    
    if "access_token" not in result:
        raise Exception(f"Failed to get token: {result.get('error_description')}")
    
    _token_cache["access_token"] = result["access_token"]
    _token_cache["expires_at"] = now + result["expires_in"]
    return _token_cache["access_token"]

# === Site ID ===
def get_site_id():
    url = "https://graph.microsoft.com/v1.0/sites/nskkogyo.sharepoint.com:/sites/2019"
    headers = {
        "Authorization": f"Bearer {get_access_token()}",
    }
    resp = requests.get(url, headers=headers)
    if resp.status_code != 200:
        raise Exception(f"Failed to get site ID: {resp.text}")
    return resp.json()["id"]


# === List All Drives ===
def list_all_drives(site_id):
    url = f"{BASE_URL}/sites/{site_id}/drives"
    headers = {"Authorization": f"Bearer {get_access_token()}"}
    resp = requests.get(url, headers=headers)
    if resp.status_code != 200:
        raise Exception(f"Failed to list drives: {resp.text}")
    
    drives = resp.json()["value"]
    logging.info("\n🔍 Available Drives:")
    for i, drive in enumerate(drives):
        logging.info(f"  {i + 1}. {drive['name']} (ID: {drive['id']})")
    return drives

# === Choose Correct Drive ===
def choose_drive_by_name(drives, name_hint):
    for drive in drives:
        if name_hint in drive["name"]:
            return drive["id"]
    raise Exception(f"No drive found with name containing: {name_hint}")

# === Determine index folder from builder name ===
def get_index_folder(builder_name):
    hira_index_map = {
        'あ': 'あ行', 'い': 'あ行', 'う': 'あ行', 'え': 'あ行', 'お': 'あ行',
        'か': 'か行', 'き': 'か行', 'く': 'か行', 'け': 'か行', 'こ': 'か行',
        'さ': 'さ行', 'し': 'さ行', 'す': 'さ行', 'せ': 'さ行', 'そ': 'さ行',
        'た': 'た行', 'ち': 'た行', 'つ': 'た行', 'て': 'た行', 'と': 'た行',
        'な': 'な行', 'に': 'な行', 'ぬ': 'な行', 'ね': 'な行', 'の': 'な行',
        'は': 'は行', 'ひ': 'は行', 'ふ': 'は行', 'へ': 'は行', 'ほ': 'は行',
        'ま': 'ま行', 'み': 'ま行', 'む': 'ま行', 'め': 'ま行', 'も': 'ま行',
    }
    first_char = builder_name[0]
    return hira_index_map.get(first_char, None)


def search_folder_in_folder(drive_id, parent_id, target_folder_name):
    """
    指定したフォルダ(parent_id)配下の子フォルダの中から、名前が一致するフォルダを探す
    """
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{parent_id}/children?$top=999"
    headers = {
        "Authorization": f"Bearer {get_access_token()}"
    }

    response = requests.get(url, headers=headers)
    if response.status_code != 200:
        logging.info(f"子フォルダ取得失敗: {response.status_code} - {response.text}")
        return False

    items = response.json().get("value", [])

    logging.info("子フォルダ一覧:")
    for item in items:
        if "folder" in item:
            logging.info(f"・ {item['name']}")

    # フォルダ名の完全一致を探す（前後空白除去）
    for item in items:
        if "folder" in item:
            folder_name = item["name"].strip()
            if folder_name == target_folder_name.strip():
                return item

    return False


    
    
def create_folder(drive_id, parent_folder_id, new_folder_name):
    url = f"{BASE_URL}/drives/{drive_id}/items/{parent_folder_id}/children"

    headers = {
        "Authorization": f"Bearer {get_access_token()}",
        "Content-Type": "application/json"
    }

    data = {
        "name": new_folder_name,
        "folder": {},  # Specifies it's a folder
        "@microsoft.graph.conflictBehavior": "replace"
    }

    resp = requests.post(url, headers=headers, json=data)
    if resp.status_code not in (200, 201):
        raise Exception(f"❌ Failed to create folder '{new_folder_name}': {resp.text}")
    
    logging.info(f"Folder '{new_folder_name}' created successfully.")
    return resp.json()

def upload_file(drive_id, parent_folder_id, file_path, file_name):
    with open(file_path, 'rb') as f:
        content = f.read()

    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{parent_folder_id}:/{file_name}:/content"
    headers = {
        "Authorization": f"Bearer {get_access_token()}",
        "Content-Type": "application/octet-stream"
    }

    response = requests.put(url, headers=headers, data=content)

    if response.status_code in [200, 201]:
        logging.info(f"Uploaded: {file_name}")
    else:
        logging.info(f"Failed: {file_name} - {response.status_code} - {response.text}")


def upload_folder(drive_id, parent_folder_id, local_folder_path):
    folder_name = os.path.basename(local_folder_path)
    sharepoint_folder = create_folder(drive_id, parent_folder_id, folder_name)
    sharepoint_folder_id = sharepoint_folder["id"]

    for root, _, files in os.walk(local_folder_path):
        for file in files:
            local_file_path = os.path.join(root, file)
            relative_path = os.path.relpath(local_file_path, local_folder_path)
            upload_file(drive_id, sharepoint_folder_id, local_file_path, relative_path)

    logging.info(f"Uploaded folder '{folder_name}'")

def search_folder_in_drive_root(drive_id, target_folder_name):
    try:
        headers = {
            "Authorization": f"Bearer {get_access_token()}",
            "Content-Type": "application/json"
        }

        # Get children of the drive root
        url = f"{BASE_URL}/drives/{drive_id}/root/children"
        resp = requests.get(url, headers=headers)

        if resp.status_code != 200:
            logging.error(f" Failed to list root folders: {resp.text}")
            return None

        items = resp.json().get("value", [])
        target_folder = next((item for item in items if item["name"] == target_folder_name), None)

        return target_folder

    except Exception as e:
        logging.error(f"Error while searching in drive root: {e}")
        return None


# === MAIN ===
# if __name__ == "__main__":
logging.basicConfig(level=logging.INFO)
def builder_sharepoint(builder_name, 案件番号, 案件名):
    try:
        logging.info(f"Builder: {builder_name}")

        # ① 固定の 2019 サイトから Site ID を取得
        site_id = get_site_id()  # get_site_id 内のURLも修正必要
        drives = list_all_drives(site_id)

        # ② 常に DocLib という名前のドライブを使用
        drive_name = "タマホーム"
        matching_drive = next((d for d in drives if d["name"] == drive_name), None)
        if not matching_drive:
            logging.info(f"Drive '{drive_name}' not found.")
            return False

        drive_id = matching_drive["id"]
        logging.info(f"Using Drive: {matching_drive['name']} (ID: {drive_id})")

        # ③ ルートにBuilder名のフォルダがあるか確認
        result = search_folder_in_drive_root(drive_id, target_folder_name=builder_name)
        if result:
            logging.info(f"Found '{builder_name}' in drive '{drive_name}'")
            logging.info(f"URL: {result['webUrl']}")
            parent_id = result["id"]
        else:
            # フォルダがない場合は作成
            result = create_folder(drive_id, "root", builder_name)
            parent_id = result["id"]
            logging.info(f"Created builder folder: {builder_name}")

        # ④ 案件フォルダ作成
        main_folder_name = f"{案件番号} {案件名}"
        main_folder = create_folder(drive_id, parent_id, main_folder_name)
        main_folder_id = main_folder["id"]

        # ⑤ サブフォルダ作成
        
        create_folder(drive_id, main_folder_id, "資料")
        logging.info(f"Created '{main_folder_name}/資料'")

        logging.info(f"✅ アップロード完了！フォルダリンク: {main_folder['webUrl']}")
        # return main_folder["webUrl"]


        # ⑥ ローカルからアップロード
        # date = datetime.datetime.now().strftime('%d_%m_%y')
        local_base_path = os.path.join(os.getcwd(), "新規案件")
        local_main_folder = None

        for name in os.listdir(local_base_path):
            if 案件番号 in name:
                local_main_folder = os.path.join(local_base_path, name)
                break

        if not local_main_folder:
            logging.info(f"'資料' フォルダに '{案件番号}' を含むものが見つかりません。")
            return False
        else:
            for subfolder_name in ["資料"]:
                subfolder_path = os.path.join(local_main_folder, subfolder_name)
                if os.path.exists(subfolder_path):
                    upload_folder(drive_id, main_folder_id, subfolder_path)
                else:
                    logging.info(f"ローカルフォルダが存在しません: {subfolder_path}")
                    return False
            return True

    except Exception as e:
        logging.info(f"❌ Error: {e}")
        return False

    
# builder = "□案件番号500000～□"
# builder_sharepoint("□案件番号500000～□", "12345", "asdfgh")


