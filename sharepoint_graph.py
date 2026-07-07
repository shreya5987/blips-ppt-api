import os
from pathlib import Path
from urllib.parse import quote

import requests
import msal
from dotenv import load_dotenv


load_dotenv()

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


def get_access_token() -> str:
    tenant_id = os.getenv("TENANT_ID")
    client_id = os.getenv("CLIENT_ID")
    client_secret = os.getenv("CLIENT_SECRET")

    if not tenant_id or not client_id or not client_secret:
        raise RuntimeError("Missing TENANT_ID, CLIENT_ID, or CLIENT_SECRET in .env")

    authority = f"https://login.microsoftonline.com/{tenant_id}"

    app = msal.ConfidentialClientApplication(
        client_id=client_id,
        authority=authority,
        client_credential=client_secret,
    )

    result = app.acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"]
    )

    if "access_token" not in result:
        raise RuntimeError(
            f"Could not acquire token: {result.get('error')} - {result.get('error_description')}"
        )

    return result["access_token"]


def get_site_id(access_token: str) -> str:
    hostname = os.getenv("SHAREPOINT_HOSTNAME")
    site_path = os.getenv("SHAREPOINT_SITE_PATH")

    if not hostname or not site_path:
        raise RuntimeError("Missing SHAREPOINT_HOSTNAME or SHAREPOINT_SITE_PATH in .env")

    url = f"{GRAPH_BASE}/sites/{hostname}:{site_path}"

    response = requests.get(
        url,
        headers={"Authorization": f"Bearer {access_token}"},
        timeout=30,
    )

    response.raise_for_status()
    return response.json()["id"]


def get_drive_id(access_token: str, site_id: str) -> str:
    drive_name = os.getenv("SHAREPOINT_DRIVE_NAME", "Documents")

    url = f"{GRAPH_BASE}/sites/{site_id}/drives"

    response = requests.get(
        url,
        headers={"Authorization": f"Bearer {access_token}"},
        timeout=30,
    )

    response.raise_for_status()

    drives = response.json()["value"]

    for drive in drives:
        if drive["name"].lower() == drive_name.lower():
            return drive["id"]

    available = [drive["name"] for drive in drives]
    raise RuntimeError(
        f"Could not find drive named '{drive_name}'. Available drives: {available}"
    )


def upload_file_to_sharepoint(local_file_path: str, target_file_name: str) -> dict:
    access_token = get_access_token()
    site_id = get_site_id(access_token)
    drive_id = get_drive_id(access_token, site_id)

    folder_path = os.getenv("SHAREPOINT_FOLDER_PATH", "").strip("/")

    if folder_path:
        target_path = f"{folder_path}/{target_file_name}"
    else:
        target_path = target_file_name

    encoded_target_path = quote(target_path)

    upload_url = (
        f"{GRAPH_BASE}/drives/{drive_id}/root:/{encoded_target_path}:/content"
    )

    file_path = Path(local_file_path)

    with open(file_path, "rb") as file:
        response = requests.put(
            upload_url,
            headers={
                "Authorization": f"Bearer {access_token}",
                "Content-Type": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            },
            data=file,
            timeout=120,
        )

    response.raise_for_status()

    uploaded = response.json()

    return {
        "name": uploaded.get("name"),
        "webUrl": uploaded.get("webUrl"),
        "id": uploaded.get("id"),
    }