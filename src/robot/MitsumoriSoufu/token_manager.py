import time

from config import CLIENT_ID, CLIENT_SECRET, TENANT_ID
from msal import ConfidentialClientApplication

_token_cache = {"access_token": None, "expires_at": 0}


def get_access_token():
    now = time.time()
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
