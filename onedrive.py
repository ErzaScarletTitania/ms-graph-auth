"""
Browse and read files from a folder shared with you via OneDrive for Business.
Uses Microsoft Graph API with tokens managed by auth.py.
"""

import base64
import json
import os
import sys

import requests

from auth import get_access_token, SECRETS_FILE

GRAPH_BASE = "https://graph.microsoft.com/v1.0"

DEFAULT_SHARED_FOLDER_NAME = "Shared Folder"


def _headers():
    return {"Authorization": f"Bearer {get_access_token()}"}


def _request_headers(extra_headers=None):
    headers = _headers()
    if extra_headers:
        headers.update(extra_headers)
    return headers


def _load_secrets():
    if not os.path.exists(SECRETS_FILE):
        return {}
    with open(SECRETS_FILE) as f:
        return json.load(f)


def _save_secrets(secrets):
    with open(SECRETS_FILE, "w") as f:
        json.dump(secrets, f, indent=2)
    os.chmod(SECRETS_FILE, 0o600)


def _cache_shared_target(target):
    secrets = _load_secrets()
    secrets["shared_folder"] = target
    _save_secrets(secrets)


def _encode_sharing_url(sharing_url):
    encoded = base64.b64encode(sharing_url.encode("utf-8")).decode("ascii")
    return "u!" + encoded.rstrip("=").replace("/", "_").replace("+", "-")


def _graph_get_json(url, extra_headers=None):
    resp = requests.get(url, headers=_request_headers(extra_headers))
    resp.raise_for_status()
    return resp.json()


def _cache_drive_item(item, mode, share_id=None, source_url=None):
    parent = item.get("parentReference", {})
    target = {
        "mode": mode,
        "drive_id": parent.get("driveId"),
        "item_id": item["id"],
        "name": item["name"],
        "web_url": item.get("webUrl", source_url or ""),
        "is_folder": bool(item.get("folder")),
    }
    if share_id:
        target["share_id"] = share_id
    _cache_shared_target(target)
    return target


def _shared_folder_name():
    return os.environ.get("MS_GRAPH_SHARED_FOLDER_NAME", DEFAULT_SHARED_FOLDER_NAME)


def find_shared_folder(name=None):
    """Return items from /me/drive/sharedWithMe, caching the named folder."""
    url = f"{GRAPH_BASE}/me/drive/sharedWithMe"
    items = _graph_get_all(url)

    if not items:
        print("No items shared with you.")
        return []

    for item in items:
        rid = item.get("remoteItem", {})
        label = (
            f"{item['name']}  "
            f"(shared by: {rid.get('createdBy', {}).get('user', {}).get('displayName', '?')})"
        )
        print(f"  [{item.get('id')}] {label}")

    target = name or _shared_folder_name()
    match = next(
        (i for i in items if i["name"].lower() == target.lower()),
        None,
    )
    if match:
        rid = match["remoteItem"]
        _cache_shared_target({
            "mode": "drive",
            "drive_id": rid["parentReference"]["driveId"],
            "item_id": rid["id"],
            "name": match["name"],
            "web_url": rid.get("webUrl", ""),
            "is_folder": bool(rid.get("folder")),
        })
        print(f"\nCached shared folder: {match['name']}")
    else:
        print(
            f"\nShared folder '{target}' was not found in sharedWithMe. "
            "Use python onedrive.py link <sharing-url> for guest or direct-link access."
        )

    return items


def use_shared_link(sharing_url):
    """Resolve and cache a OneDrive/SharePoint sharing URL for guest-friendly access."""
    share_id = _encode_sharing_url(sharing_url)
    url = f"{GRAPH_BASE}/shares/{share_id}/driveItem"
    item = _graph_get_json(url, extra_headers={"Prefer": "redeemSharingLink"})
    target = _cache_drive_item(item, mode="share", share_id=share_id, source_url=sharing_url)
    print(f"Cached shared link: {target['name']}")
    return item


def _graph_get_all(url, extra_headers=None):
    """Handle paginated Graph API responses."""
    results = []
    while url:
        resp = requests.get(url, headers=_request_headers(extra_headers))
        resp.raise_for_status()
        data = resp.json()
        results.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return results


def _get_shared_target():
    secrets = _load_secrets()
    sf = secrets.get("shared_folder")
    if not sf:
        print("No cached shared folder. Searching sharedWithMe...")
        find_shared_folder()
        secrets = _load_secrets()
        sf = secrets.get("shared_folder")
        if not sf:
            raise RuntimeError(
                f"Shared folder '{_shared_folder_name()}' not found. "
                "Run: python onedrive.py link <sharing-url> or python onedrive.py shared"
            )
    return sf


def _drive_item_url(drive_id, item_id):
    return f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}"


def _resolve_item(target, path=""):
    drive_id = target.get("drive_id")
    item_id = target.get("item_id")
    if not drive_id or not item_id:
        raise RuntimeError("Cached shared target does not include a drive ID or item ID.")

    if not path:
        return {"drive_id": drive_id, "item_id": item_id}

    current_item_id = item_id
    for segment in [part for part in path.split("/") if part]:
        url = f"{_drive_item_url(drive_id, current_item_id)}/children"
        items = _graph_get_all(url)
        match = next((item for item in items if item.get("name") == segment), None)
        if not match:
            raise FileNotFoundError(f"Path not found: {path}")
        current_item_id = match["id"]

    return {"drive_id": drive_id, "item_id": current_item_id}


def _print_items(items):
    print(f"\n{'Name':<40} {'Type':<12} {'Size':>10} {'Modified':>20}   ID")
    print("-" * 110)
    for item in items:
        name = item["name"][:38] + ".." if len(item["name"]) > 40 else item["name"]
        ftype = "FOLDER" if item.get("folder") else "FILE"
        size = str(item.get("size", "")) if "size" in item else ""
        mod = item.get("lastModifiedDateTime", "")[:19].replace("T", " ")
        print(f"{name:<40} {ftype:<12} {size:>10} {mod:>20}   {item['id']}")


def list_folder(path=""):
    """List children of the shared folder, or a subfolder within it."""
    target = _get_shared_target()

    if not path and not target.get("is_folder"):
        item = {
            "id": target["item_id"],
            "name": target["name"],
            "size": "",
            "lastModifiedDateTime": "",
        }
        _print_items([item])
        return [item]

    resolved = _resolve_item(target, path)
    url = f"{_drive_item_url(resolved['drive_id'], resolved['item_id'])}/children"

    items = _graph_get_all(url)

    if not items:
        print(f"No items found in '{path or 'root'}'.")
        return []

    _print_items(items)

    return items


def download_file(item_id, dest_dir="."):
    """Download a file by its item ID. Destination directory is created if needed."""
    target = _get_shared_target()
    drive_id = target.get("drive_id")
    if not drive_id:
        raise RuntimeError("Cached shared target does not include a drive ID.")

    meta_url = f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}"
    meta = _graph_get_json(meta_url)
    filename = meta.get("name", item_id)

    content_url = f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/content"
    resp = requests.get(content_url, headers=_headers())
    resp.raise_for_status()

    os.makedirs(dest_dir, exist_ok=True)
    dest_path = os.path.join(dest_dir, filename)
    with open(dest_path, "wb") as f:
        f.write(resp.content)

    print(f"Downloaded: {dest_path}  ({len(resp.content):,} bytes)")
    return dest_path


def get_content(item_id):
    """Return file bytes — useful for in-memory inspection by agents."""
    target = _get_shared_target()
    drive_id = target.get("drive_id")
    if not drive_id:
        raise RuntimeError("Cached shared target does not include a drive ID.")
    url = f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/content"
    resp = requests.get(url, headers=_headers())
    resp.raise_for_status()
    return resp.content


def search_shared(query):
    """Search within the shared folder by keyword."""
    target = _get_shared_target()
    drive_id = target.get("drive_id")
    item_id = target.get("item_id")
    if not drive_id or not item_id:
        raise RuntimeError("Cached shared target does not include a drive ID or item ID.")
    escaped_query = query.replace("'", "''")
    url = f"{_drive_item_url(drive_id, item_id)}/search(q='{escaped_query}')"
    items = _graph_get_all(url)

    if not items:
        print(f"No results for '{query}'.")
        return []

    print(f"\nResults for '{query}':")
    for item in items:
        rid = item.get("parentReference", {})
        path = rid.get("path", "") + "/" + item["name"]
        print(f"  {path}")
        print(f"    ID: {item['id']}  |  Modified: {item.get('lastModifiedDateTime', '?')[:19]}")

    return items


# ── CLI ──────────────────────────────────────────────────────────────────────

def _cli():
    if len(sys.argv) < 2:
        print("Usage:")
        print("  python onedrive.py shared              — list items shared with you")
        print("  python onedrive.py link <sharing-url>  — cache a shared link directly")
        print("  python onedrive.py list [subfolder]    — list folder contents")
        print("  python onedrive.py download <id> [dir] — download a file")
        print("  python onedrive.py search <query>      — search for files")
        sys.exit(1)

    cmd = sys.argv[1]

    try:
        _load_secrets()  # trigger early FileNotFoundError for nice message
    except (FileNotFoundError, json.JSONDecodeError):
        pass

    if cmd == "shared":
        find_shared_folder()
    elif cmd == "link":
        if len(sys.argv) < 3:
            print("Usage: python onedrive.py link <sharing-url>")
            sys.exit(1)
        use_shared_link(sys.argv[2])
    elif cmd == "list":
        path = sys.argv[2] if len(sys.argv) > 2 else ""
        list_folder(path)
    elif cmd == "download":
        if len(sys.argv) < 3:
            print("Usage: python onedrive.py download <item-id> [dest-dir]")
            sys.exit(1)
        dest = sys.argv[3] if len(sys.argv) > 3 else "."
        download_file(sys.argv[2], dest)
    elif cmd == "search":
        if len(sys.argv) < 3:
            print("Usage: python onedrive.py search <query>")
            sys.exit(1)
        search_shared(sys.argv[2])
    else:
        print(f"Unknown command: {cmd}")
        sys.exit(1)


if __name__ == "__main__":
    _cli()
