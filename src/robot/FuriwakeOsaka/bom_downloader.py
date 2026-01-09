# -*- coding: utf-8 -*-
"""
bom_downloader.py — Download factory BOM files from SharePoint via Graph

Supports multiple factories via SharePoint /shares:

- Each factory has a base URL template with a `{date}` placeholder.
- The `{date}` part is typically the Japanese-style folder name like "11月15日"
  or "11月15日配送分", passed in from Main.py.

- All xls/xlsx/xlsm/pdf/csv files are downloaded into:
    <base_dir>/BOM/<date>/

- Designed to be called from Main.py:
    download_factory_bom_for_date("大阪", "11月15日", Path.cwd())

For backward compatibility, the old helper:
    download_osaka_bom_for_date("11月15日", Path.cwd())
is still provided and internally calls the generic function.
"""

import logging
import base64
from pathlib import Path

import requests
from token_manager import get_access_token
from config import BASE_URL  # same BASE_URL used in graph_downloader.py
import os
from pathlib import Path, PurePosixPath


GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"


# ---------------------------------------------------------------------------
# Factory → SharePoint URL templates
# ---------------------------------------------------------------------------
# NOTE:
# - `{date}` will be replaced with jp_folder_name passed from Main.py.
# - Make sure jp_folder_name matches the actual folder naming, e.g.:
#     "11月15日"            → .../11月15日
#     "11月15日配送分"      → .../11月15日配送分
# ---------------------------------------------------------------------------
FACTORY_SHARE_URLS: dict[str, str] = {
    # 大阪
    "大阪": (
        "https://nskkogyo.sharepoint.com/sites/yanase/"
        "Shared Documents/大阪工場　製造データ/{date}/🔹関西工場確定データ🔹"
    ),

    # 栃木（真岡工場）
    "栃木": (
        "https://nskkogyo.sharepoint.com/sites/mouka/"
        "Shared Documents/真岡工場　製造データ/{date}/栃木工場確定データ"
    ),

    # 千葉
    "千葉": (
        "https://nskkogyo.sharepoint.com/sites/nskhome/"
        "Shared Documents/千葉工場 製造データ/{date}"
    ),

    # 豊橋
    "豊橋": (
        "https://nskkogyo.sharepoint.com/sites/toyohashi/"
        "Shared Documents/豊橋工場製造データ/{date}"
    ),

    # 九州
    "九州": (
        "https://nskkogyo.sharepoint.com/sites/kyuusyuukouzyou/"
        "Shared Documents/九州工場 製造データー/{date}/製造"
    ),

    # 滋賀
    "滋賀": (
        "https://nskkogyo.sharepoint.com/sites/shiga/"
        "Shared Documents/滋賀工場 製造データ/{date}/製造　手配済み(DL済み)"
    ),
}


def _encode_share_url(url: str) -> str:
    """
    Graph /shares API expects a base64url-encoded share URL with 'u!' prefix.
    """
    b64 = base64.b64encode(url.encode("utf-8")).decode("ascii")
    b64 = b64.rstrip("=")          # remove padding
    b64 = b64.replace("+", "-").replace("/", "_")
    return f"u!{b64}"


def download_factory_bom_for_date(
    factory_label: str,
    jp_folder_name: str,
    base_dir: Path,
) -> Path | None:
    """
    Generic factory BOM downloader.

    factory_label:
        Factory key as used in FACTORY_SHARE_URLS, e.g. "大阪", "栃木", "千葉", "豊橋", "九州", "滋賀".

    jp_folder_name:
        e.g. "11月15日" or "11月15日配送分"
        (must match the actual folder naming on SharePoint).

    base_dir:
        Usually Path.cwd() from Main.py.

    Returns:
        Path to <base_dir>/BOM/<jp_folder_name> if files were downloaded,
        or None if folder not found / no files / factory not configured.
    """
    factory_key = factory_label.strip()

    if factory_key not in FACTORY_SHARE_URLS:
        logging.info(
            f"[Graph] No SharePoint BOM path configured for factory: {factory_key}"
        )
        return None

    url_template = FACTORY_SHARE_URLS[factory_key]
    target_url = url_template.format(date=jp_folder_name)

    logging.info(
        f"[Graph] BOM target SharePoint URL for factory '{factory_key}': {target_url}"
    )

    share_id = _encode_share_url(target_url)
    list_url = f"{BASE_URL}/shares/{share_id}/driveItem/children"
    headers = {"Authorization": f"Bearer {get_access_token()}"}

    try:
        resp = requests.get(list_url, headers=headers)
    except Exception as e:
        logging.error(
            f"[Graph] BOM list request failed for factory '{factory_key}': {e}"
        )
        return None

    if resp.status_code == 404:
        logging.warning(
            f"[Graph] BOM folder not found for factory '{factory_key}', "
            f"date '{jp_folder_name}'"
        )
        return None

    try:
        resp.raise_for_status()
    except Exception as e:
        logging.error(
            f"[Graph] BOM list error for factory '{factory_key}': {e} | "
            f"body={resp.text[:500]}"
        )
        return None

    items = resp.json().get("value", [])
    if not items:
        logging.warning(
            f"[Graph] BOM folder is empty for factory '{factory_key}', "
            f"date '{jp_folder_name}'"
        )
        return None

    dest_root = Path(base_dir) / "BOM" / jp_folder_name
    dest_root.mkdir(parents=True, exist_ok=True)

    count = 0
    for it in items:
        # skip sub-folders
        if "file" not in it:
            continue

        name = it.get("name", "")
        # only BOM-related formats
        if not any(
            name.lower().endswith(ext)
            for ext in (".xlsx", ".xlsm", ".xls", ".pdf", ".csv")
        ):
            continue

        drive_id = it["parentReference"]["driveId"]
        file_id = it["id"]
        dl_url = f"{BASE_URL}/drives/{drive_id}/items/{file_id}/content"

        logging.info(
            f"[Graph] Downloading BOM file for '{factory_key}': {name}"
        )
        try:
            r = requests.get(dl_url, headers=headers, stream=True)
            r.raise_for_status()
        except Exception as e:
            logging.error(
                f"[Graph] Download failed for {name} (factory '{factory_key}'): "
                f"{e} | body={getattr(r, 'text', '')[:200]}"
            )
            continue

        out_path = dest_root / name
        with open(out_path, "wb") as f:
            for chunk in r.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)

        count += 1

    logging.info(
        f"[Graph] BOM download finished for factory '{factory_key}' — "
        f"{count} files saved under: {dest_root}"
    )
    return dest_root if count > 0 else None

# ---------------------------------------------------------------------------
# 配車表 (大阪) downloader — same Graph /shares style, different folder
# ---------------------------------------------------------------------------

# ---------------------------------------------------------------------------
# 配車表 (大阪) downloader — now writes into ./配車表
# ---------------------------------------------------------------------------

OSAKA_HAISHA_SHARE_URL_TEMPLATE = (
    "https://nskkogyo.sharepoint.com/sites/yanase/"
    "Shared Documents/大阪工場　製造データ/{date}"
)

def download_osaka_haisha_for_date(jp_folder_name: str, base_dir: Path) -> Path | None:
    """
    Download 配車表 Excel(s) for 大阪 from:

        https://nskkogyo.sharepoint.com/sites/yanase/
        Shared Documents/大阪工場　製造データ/{date}

    and save them under:

        <base_dir>/配車表/

    Only files whose name contains '配車' and ends with .xls / .xlsx are downloaded.
    """
    target_url = OSAKA_HAISHA_SHARE_URL_TEMPLATE.format(date=jp_folder_name)

    logging.info(f"[Graph] 配車表 target SharePoint URL (大阪): {target_url}")

    share_id = _encode_share_url(target_url)
    list_url = f"{BASE_URL}/shares/{share_id}/driveItem/children"
    headers = {"Authorization": f"Bearer {get_access_token()}"}

    try:
        resp = requests.get(list_url, headers=headers)
    except Exception as e:
        logging.error(f"[Graph] 配車表 list request failed (大阪): {e}")
        return None

    if resp.status_code == 404:
        logging.warning(f"[Graph] 配車表 folder not found for 大阪, date '{jp_folder_name}'")
        return None

    try:
        resp.raise_for_status()
    except Exception as e:
        logging.error(f"[Graph] 配車表 list error (大阪): {e} | body={resp.text[:500]}")
        return None

    items = resp.json().get("value", [])
    if not items:
        logging.warning(f"[Graph] 配車表 folder is empty for 大阪, date '{jp_folder_name}'")
        return None

    dest_root = Path(base_dir) / "配車表"
    dest_root.mkdir(parents=True, exist_ok=True)

    count = 0
    for it in items:
        if "file" not in it:
            continue

        name = it.get("name", "")
        lower = name.lower()

        if "配車" not in name:
            continue
        if not (lower.endswith(".xls") or lower.endswith(".xlsx")):
            continue

        drive_id = it["parentReference"]["driveId"]
        file_id  = it["id"]
        dl_url   = f"{BASE_URL}/drives/{drive_id}/items/{file_id}/content"

        logging.info(f"[Graph] Downloading 配車表 (大阪): {name}")
        try:
            r = requests.get(dl_url, headers=headers, stream=True)
            r.raise_for_status()
        except Exception as e:
            body = getattr(r, "text", "")[:200] if "r" in locals() else ""
            logging.error(
                f"[Graph] 配車表 download failed for {name} (大阪): {e} | body={body}"
            )
            continue

        out_path = dest_root / name
        with open(out_path, "wb") as f:
            for chunk in r.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)

        count += 1

    if count == 0:
        logging.warning(
            f"[Graph] 配車表 not found (no matching Excel with '配車') "
            f"for 大阪, date '{jp_folder_name}'"
        )
        return None

    logging.info(
        f"[Graph] 配車表 download finished for 大阪 — {count} file(s) saved under: {dest_root}"
    )
    return dest_root

TOCHIGI_HAISHA_SHARE_URL_TEMPLATE = (
    "https://nskkogyo.sharepoint.com/sites/mouka/"
    "Shared Documents/真岡工場　製造データ/{date}"
)

def download_tochigi_haisha_for_date(jp_folder_name: str, base_dir: Path) -> Path | None:
    """
    配車表 downloader for 栃木.
    The 配車表 is directly under the DATE ROOT.
    """
    target_url = TOCHIGI_HAISHA_SHARE_URL_TEMPLATE.format(date=jp_folder_name)

    logging.info(f"[Graph] 配車表 target SharePoint URL (栃木): {target_url}")

    share_id = _encode_share_url(target_url)
    list_url = f"{BASE_URL}/shares/{share_id}/driveItem/children"
    headers = {"Authorization": f"Bearer {get_access_token()}"}

    try:
        resp = requests.get(list_url, headers=headers)
        resp.raise_for_status()
    except Exception as e:
        logging.error(f"[Graph] 配車表 list request failed (栃木): {e}")
        return None

    items = resp.json().get("value", [])
    if not items:
        logging.warning(f"[Graph] 配車表 folder empty for 栃木, date '{jp_folder_name}'")
        return None

    dest_root = Path(base_dir) / "配車表"
    dest_root.mkdir(parents=True, exist_ok=True)

    count = 0
    for it in items:
        if "file" not in it:
            continue

        name = it.get("name", "")
        lower = name.lower()

        # SAME RULE AS OSAKA — 必ず "配車" + excel 
        if "配車" not in name:
            continue
        if not (lower.endswith(".xls") or lower.endswith(".xlsx")):
            continue

        drive_id = it["parentReference"]["driveId"]
        file_id = it["id"]
        dl_url = f"{BASE_URL}/drives/{drive_id}/items/{file_id}/content"

        logging.info(f"[Graph] Downloading 配車表 (栃木): {name}")
        try:
            r = requests.get(dl_url, headers=headers, stream=True)
            r.raise_for_status()
        except Exception as e:
            logging.error(f"[Graph] 配車表 download failed for {name} (栃木): {e}")
            continue

        out_path = dest_root / name
        with open(out_path, "wb") as f:
            for chunk in r.iter_content(chunk_size=8192):
                f.write(chunk)

        count += 1

    if count == 0:
        logging.warning(f"[Graph] No 配車 Excel found (栃木) for date '{jp_folder_name}'")
        return None

    logging.info(f"[Graph] 配車表 download finished for 栃木 → {count} files")
    return dest_root


# ---------------------------------------------------------------------------
# Backward-compatible wrapper (Osaka only)
# ---------------------------------------------------------------------------
def download_osaka_bom_for_date(jp_folder_name: str, base_dir: Path) -> Path | None:
    """
    Legacy helper kept for backward compatibility.

    Internally calls:
        download_factory_bom_for_date("大阪", jp_folder_name, base_dir)
    """
    return download_factory_bom_for_date("大阪", jp_folder_name, base_dir)

def upload_usb_to_tochigi_date(jp_date: str, usb_folder: Path) -> int:
    """
    Uploads the local ▽USB folder contents into:

        真岡工場 製造データ/{date}/A
        真岡工場 製造データ/{date}/B
        真岡工場 製造データ/{date}/C
        ...

    (No USB folder exists in Tochigi — truck folders go directly
     into the DATE ROOT.)

    Returns number of uploaded files.
    """

    usb_folder = Path(usb_folder)
    if not usb_folder.exists():
        logging.info(f"[Tochigi Upload] USB folder not found: {usb_folder}")
        return 0

    # --- TARGET ROOT ---
    target_url = (
        "https://nskkogyo.sharepoint.com/sites/mouka/"
        "Shared Documents/真岡工場　製造データ/{date}"
    ).format(date=jp_date)

    logging.info(f"[Graph] Tochigi upload target = {target_url}")

    share_id = _encode_share_url(target_url)
    token = get_access_token()
    headers = {"Authorization": f"Bearer {token}"}

    # Resolve DATE folder driveItem
    resp = requests.get(
        f"{GRAPH_BASE_URL}/shares/{share_id}/driveItem",
        headers=headers
    )
    if not resp.ok:
        logging.error(
            f"[Tochigi Upload] Failed to resolve date folder: "
            f"{resp.status_code} {resp.text}"
        )
        return 0

    info = resp.json()
    drive_id = info["parentReference"]["driveId"]
    date_folder_id = info["id"]

    # --- UPLOAD ---
    uploaded = 0

    # Iterate through ▽USB/*  (truck folders)
    for root, dirs, files in os.walk(usb_folder):
        rel_root = Path(root).relative_to(usb_folder)

        for fname in files:
            local_path = Path(root) / fname

            # For Tochigi — upload directly under date folder
            # Ex: A/file.pdf → {date}/A/file.pdf
            if rel_root == Path("."):
                # A file directly inside ▽USB (rare but supported)
                remote_rel = fname
            else:
                remote_rel = str(
                    PurePosixPath(rel_root.as_posix()) / fname
                )

            put_url = (
                f"{GRAPH_BASE_URL}/drives/{drive_id}/items/"
                f"{date_folder_id}:/{remote_rel}:/content"
            )

            try:
                with open(local_path, "rb") as f:
                    put_resp = requests.put(put_url, headers=headers, data=f)

                if put_resp.status_code in (200, 201):
                    uploaded += 1
                    logging.info(
                        f"[Tochigi Upload] OK → {remote_rel}"
                    )
                else:
                    logging.error(
                        f"[Tochigi Upload] FAILED {remote_rel}: "
                        f"{put_resp.status_code} {put_resp.text}"
                    )
            except Exception as e:
                logging.error(
                    f"[Tochigi Upload] Exception for {remote_rel}: {e}"
                )

    logging.info(f"[Tochigi Upload] DONE → {uploaded} file(s)")
    return uploaded

