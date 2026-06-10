"""
Auth for Microsoft Graph — interactive browser flow (PKCE + localhost).
Tries multiple well-known Microsoft client IDs. No app registration needed.
Saves tokens to secrets.json after one-time browser sign-in.
"""

import json
import os
import time

SCOPES = ["Files.ReadWrite", "Sites.Read.All", "User.Read"]
SECRETS_FILE = os.path.join(os.path.dirname(__file__), "secrets.json")

CLIENT_IDS = [
    "14d82eec-204b-4c2f-b7e8-296a70dab67e",  # Microsoft Graph CLI
    "de8bc8b5-d9f9-48b1-a8ad-b748da725064",  # Graph Explorer
    "1950a258-227b-4e31-a9cf-717495945fc2",  # Azure PowerShell
]

AUTHORITY = "https://login.microsoftonline.com/common"


def _save_secrets(access_token, refresh_token, expires_in, client_id):
    secrets = {
        "access_token": access_token,
        "refresh_token": refresh_token,
        "expires_at": time.time() + expires_in,
        "client_id": client_id,
    }
    with open(SECRETS_FILE, "w") as f:
        json.dump(secrets, f, indent=2)
    os.chmod(SECRETS_FILE, 0o600)


def do_auth_flow():
    """Interactive browser auth with PKCE. Opens browser, you sign in,
    redirect lands on localhost:8399 where MSAL captures the auth code.
    Long timeout in case your org requires an access justification form."""
    import msal

    for client_id in CLIENT_IDS:
        print(f"\nClient {client_id[:8]}... opening browser")
        app = msal.PublicClientApplication(client_id, authority=AUTHORITY)

        for port in [8399, 8400, 8401]:
            try:
                result = app.acquire_token_interactive(
                    scopes=SCOPES,
                    port=port,
                    prompt="select_account",
                    timeout=600,  # 10 minutes — plenty of time for access justification
                )
            except Exception as e:
                print(f"  Port {port}: {e}")
                continue

            if "access_token" in result:
                _save_secrets(
                    result["access_token"],
                    result["refresh_token"],
                    result["expires_in"],
                    client_id,
                )
                print(f"  Signed in as: {result.get('id_token_claims', {}).get('upn', '?')}")
                print("  Tokens saved to secrets.json.")
                return

            error = result.get("error_description", result.get("error", ""))
            if "AADSTS" in error:
                print(f"  Auth error: {error[:150]}")
                break  # Don't retry ports for fatal errors
            else:
                print(f"  {error[:100]}")

    raise RuntimeError(
        "All client IDs failed.\n"
        "Workaround: register your own app at https://portal.azure.com\n"
        "  - 'Accounts in any organizational directory'\n"
        "  - Redirect URI: 'http://localhost:8399'\n"
        "  - Copy client ID into auth.py and retry."
    )


def get_access_token():
    """Return a valid access token, refreshing if needed."""
    import msal

    if not os.path.exists(SECRETS_FILE):
        raise FileNotFoundError("No secrets.json found. Run: python auth.py")

    with open(SECRETS_FILE) as f:
        secrets = json.load(f)

    if secrets.get("expires_at", 0) > time.time() + 60:
        return secrets["access_token"]

    print("Refreshing access token...")
    client_id = secrets.get("client_id", CLIENT_IDS[0])
    app = msal.PublicClientApplication(client_id, authority=AUTHORITY)

    result = app.acquire_token_by_refresh_token(
        secrets["refresh_token"], scopes=SCOPES
    )

    if "access_token" not in result:
        raise RuntimeError(
            f"Token refresh failed. Re-run auth.py. "
            f"Error: {result.get('error_description', result)}"
        )

    _save_secrets(
        result["access_token"], result["refresh_token"], result["expires_in"], client_id
    )
    return result["access_token"]


if __name__ == "__main__":
    do_auth_flow()
