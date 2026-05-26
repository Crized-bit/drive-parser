# Google Drive Parser

A dark-themed desktop app to recursively scan a Google Drive folder, browse its contents, and export to Excel.

---

## Features

- Recursive folder scan with live progress
- Real-time file count and status
- Natural sort (1, 2, 10 — not 1, 10, 2), grouped by folder
- Filter by name — plain text or regex
- Multi-row selection with column copy
- Export to styled `.xlsx`

---

## Getting Started

### Prerequisites

- Python 3.10+
- A Google OAuth **Access Token** with `https://www.googleapis.com/auth/drive.readonly` scope

### Install dependencies

```bash
pip install -r requirements.txt
```

### Run

```bash
python drive_parser.py
# or
make run
```

---

## How to Use

### 1. Get an Access Token

The easiest way is via the [OAuth 2.0 Playground](https://developers.google.com/oauthplayground):

1. Open the Playground and select **Drive API v3 → `drive.readonly`**
2. Authorize and exchange for an access token
3. Copy the `access_token` value

### 2. Get a Folder ID

Open the folder in Google Drive — the ID is the last part of the URL:

```
https://drive.google.com/drive/folders/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs
                                        ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
```

### 3. Scan

Paste the token and folder ID into the toolbar, then click **▶ Start**.
Files appear in the table as they are found. Click **■ Stop** to abort early.

---

## Filtering

The filter bar searches file names in real time.

| Mode | How to activate | Example |
|------|----------------|---------|
| Plain text | Default | `report` |
| Regex | Click the `.*` button (turns blue) | `report_\d{4}\.pdf` |

When a regex is invalid, the input border turns red — the current view is preserved until the pattern is valid again.

---

## Sorting

Click any column header to sort. Rules:

- Files are always **grouped by folder** first
- Within each folder, the clicked column is sorted using **natural order** — so `file2` comes before `file10`
- Clicking **Folder** sorts all files by their folder path naturally

---

## Copying Data

1. **Shift+click** one or more column headers to mark them for copy — marked headers show a `·` prefix
2. **Select rows** in the table — click to select one, Shift+click for a range, Cmd+click to toggle individual rows
3. **Cmd+C** to copy

Values are tab-separated across columns and newline-separated across rows, so they paste cleanly into Excel or Numbers.

---

## Export

Click **Export Excel** in the bottom-right corner. The `.xlsx` file includes:

- Styled header row
- Alternating row colours
- Auto-fitted column widths
- Frozen header row for easy scrolling

---

## Building

Builds are automated via GitHub Actions on every tag push.

```bash
make release VERSION=v1.0.0
```

Artifacts are attached to the GitHub Release automatically:

| File | Platform |
|------|----------|
| `drive_parser.exe` | Windows |
| `drive_parser-macos.zip` | macOS `.app` bundle |

To re-release the same version:

```bash
make retag VERSION=v1.0.0
```

---

## Makefile reference

| Command | Description |
|---------|-------------|
| `make run` | Run the app locally |
| `make push` | Push `main` to GitHub |
| `make release VERSION=vX.Y.Z` | Tag and push → triggers CI build and release |
| `make retag VERSION=vX.Y.Z` | Delete and recreate a tag |
