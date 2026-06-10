# ms-graph-auth

**A Microsoft Graph automation tool built to turn browser-only access to shared OneDrive / SharePoint content into programmable access when tenant constraints block the normal API path.**

This project is a small Python CLI for authenticating with Microsoft Graph and browsing shared OneDrive / SharePoint content. I created it because I had user-level access to shared content in the browser, but no practical way to turn that access into an API/token workflow I could automate: I did not control the tenant, could not rely on creating my own app registration, and could not directly connect an agent to the shared OneDrive content.

That is the value of this repo as a showcase: **taking a blocked enterprise access problem, using AI to accelerate investigation and iteration, and turning browser-only access into a working developer tool and automation path.**

## Why this stands out

- **Real Microsoft Graph integration** with delegated auth, token caching, and refresh.
- **Enterprise debugging mindset** for guest access, tenant constraints, and API onboarding gaps.
- **AI-assisted orchestration** to speed up investigation, validate options, and automate the final workflow.
- **Usable CLI tooling** instead of a one-off script.
- **Practical access bridge**: when normal API onboarding is blocked, use delegated auth plus shared-link resolution to create programmable access.

## What it does

- Opens an interactive browser sign-in flow with MSAL.
- Stores and refreshes Graph tokens locally.
- Caches either:
  - a shared folder discovered through `sharedWithMe`, or
  - a direct OneDrive / SharePoint sharing URL.
- Lists folder contents, including nested folders.
- Searches within the cached shared target.
- Downloads files by item ID.

## Architecture in one minute

1. `auth.py` signs the user in through Microsoft identity and stores tokens in `secrets.json`.
2. `onedrive.py` resolves either a `sharedWithMe` item or a sharing URL into a cached drive/item target.
3. The CLI then reuses the cached drive/item IDs for listing, searching, and downloading.

The key engineering decision is the shared-link fallback. In some enterprise setups, a user can access shared content in the browser, but there is no straightforward way to generate a usable token flow or connect automation directly from the inviting side. Resolving the actual sharing URL through Graph `/shares/.../driveItem`, combined with delegated user authentication, turns that browser-visible access into something programmable and reusable.

## Setup

```powershell
pip install -r requirements.txt
```

## Authenticate

```powershell
python auth.py
```

This opens the browser, completes sign-in, and stores tokens in `secrets.json`.

## Usage

```powershell
# Optional: try the standard Graph sharedWithMe route
python onedrive.py shared

# Preferred for guest/external access: cache a direct sharing link
python onedrive.py link "https://tenant.sharepoint.com/:f:/g/..."

# List the cached root folder
python onedrive.py list

# List a nested folder
python onedrive.py list "Subfolder/Name"

# Search inside the cached folder
python onedrive.py search "minutes"

# Download a file by item ID
python onedrive.py download <item-id>
python onedrive.py download <item-id> .\downloads
```

If the standard `sharedWithMe` flow fails for a guest/external account, use the browser-visible sharing link:

```powershell
python onedrive.py link "<sharing-url>"
```

After that, `list`, `search`, and `download` reuse the cached target.

To change the default name used by `python onedrive.py shared`, set:

```powershell
$env:MS_GRAPH_SHARED_FOLDER_NAME = "Shared Folder"
```

## Example output

```text
Cached shared link: Shared Engineering Folder

Name                                     Type               Size             Modified   ID
--------------------------------------------------------------------------------------------------------------
Architecture                             FOLDER                0  2026-05-30 09:14:22   01ABC...
API notes.md                             FILE               8421  2026-06-01 15:41:10   01DEF...
Integration test report.pdf              FILE            1452284  2026-06-03 11:08:57   01GHI...
```

## Recruiter demo

If you want to demo this repo live, use this story:

> I had browser access to shared Microsoft 365 content, but no practical way to turn that into programmable access because I did not control the tenant, app registration, or token flow. I used AI to accelerate investigation, found a workable Graph path through delegated auth plus shared-link resolution, and turned it into a reusable CLI.

### 60-second demo flow

```powershell
# 1. Authenticate once
python auth.py

# 2. Cache the shared link when normal API onboarding is blocked
python onedrive.py link "https://tenant.sharepoint.com/:f:/g/..."

# 3. Show that the shared folder is now programmable
python onedrive.py list

# 4. Show search
python onedrive.py search "meeting"

# 5. Show download capability
python onedrive.py download <item-id> .\downloads
```

### Sanitized demo transcript

```text
> python onedrive.py link "https://tenant.sharepoint.com/:f:/g/..."
Cached shared link: Shared Engineering Folder

> python onedrive.py list
Name                                     Type               Size             Modified   ID
--------------------------------------------------------------------------------------------------------------
Architecture                             FOLDER                0  2026-05-30 09:14:22   01ABC...
Meeting notes.docx                       FILE              18244  2026-06-02 11:20:04   01DEF...
Integration status.xlsx                  FILE              25412  2026-06-03 08:15:31   01GHI...

> python onedrive.py search "meeting"
Results for 'meeting':
  /Meeting notes.docx
    ID: 01DEF...  |  Modified: 2026-06-02T11:20:04

> python onedrive.py download 01DEF... .\downloads
Downloaded: .\downloads\Meeting notes.docx  (18,244 bytes)
```

### What to emphasize while demoing

1. **The problem was access orchestration, not just coding.**
2. **The solution was to bridge browser-only access into automation.**
3. **AI helped speed up investigation and iteration, but the value was shipping a reusable workflow.**

## Using it from Python

```python
from auth import get_access_token
from onedrive import use_shared_link, list_folder, download_file, get_content, search_shared

token = get_access_token()

use_shared_link("https://tenant.sharepoint.com/:f:/g/...")
files = list_folder("Subfolder/Name")
pdf_bytes = get_content("<item-id>")
download_file("<item-id>", "./downloads")
```

## Files

- `auth.py` — interactive Graph authentication and token refresh.
- `onedrive.py` — shared-drive browsing, search, and download commands.
- `requirements.txt` — runtime dependencies.
- `.gitignore` — prevents local token files from being committed.

## Security notes

- `secrets.json` is local-only and ignored by git.
- Do not publish real tenant links, file IDs, downloaded customer files, or credentials.
- If you make this repo public, keep all examples sanitized.

## File structure

```text
ms-graph-auth/
├── auth.py           Interactive Graph auth + token refresh
├── onedrive.py       Shared-drive browsing, search, download
├── secrets.json      Tokens (auto-created, gitignored)
├── requirements.txt  requests + msal
└── .gitignore
```
