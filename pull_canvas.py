"""
Pull files from Canvas LMS module pages and organize them into local folders.

Usage (via GitHub Actions):
  Set CANVAS_TOKEN and CANVAS_COURSE_ID env vars.
  Optionally set MODULE_FILTER to download only modules whose name contains
  the filter string (case-insensitive).

Output goes to downloads/<Module Name>/<filename>.
"""

import os
import re
import sys
import requests
from pathlib import Path

CANVAS_BASE = "https://canvas.stanford.edu/api/v1"
TOKEN = os.environ.get("CANVAS_TOKEN", "")
COURSE_ID = os.environ.get("CANVAS_COURSE_ID", "")
MODULE_FILTER = os.environ.get("MODULE_FILTER", "").strip()
OUTPUT_DIR = Path("downloads")

if not TOKEN:
    print("ERROR: CANVAS_TOKEN environment variable is not set.")
    sys.exit(1)
if not COURSE_ID:
    print("ERROR: CANVAS_COURSE_ID environment variable is not set.")
    sys.exit(1)

HEADERS = {"Authorization": f"Bearer {TOKEN}"}


def paginate(url, params=None):
    """Yield all items across paginated Canvas API responses."""
    while url:
        r = requests.get(url, headers=HEADERS, params=params, timeout=30)
        r.raise_for_status()
        yield from r.json()
        next_url = None
        for part in r.headers.get("Link", "").split(","):
            if 'rel="next"' in part:
                m = re.search(r"<([^>]+)>", part)
                if m:
                    next_url = m.group(1)
        url = next_url
        params = None  # subsequent pages have params baked into the URL


def safe_name(name: str) -> str:
    """Strip characters that are illegal in file/directory names."""
    return re.sub(r'[\\/:*?"<>|]', "_", name).strip()


def download_file(file_id: int, dest_dir: Path) -> None:
    """Fetch file metadata from Canvas, then download the actual binary."""
    r = requests.get(f"{CANVAS_BASE}/files/{file_id}", headers=HEADERS, timeout=30)
    if r.status_code == 404:
        print(f"    [skip] file {file_id} — not found or access restricted")
        return
    r.raise_for_status()
    info = r.json()

    filename = safe_name(info["display_name"])
    dest = dest_dir / filename
    if dest.exists():
        print(f"    [exists] {filename}")
        return

    # Canvas returns a pre-signed S3 URL; download it without the auth header
    resp = requests.get(info["url"], stream=True, timeout=60)
    resp.raise_for_status()
    dest_dir.mkdir(parents=True, exist_ok=True)
    with open(dest, "wb") as f:
        for chunk in resp.iter_content(chunk_size=65536):
            f.write(chunk)
    size_kb = dest.stat().st_size // 1024
    print(f"    [ok] {filename}  ({size_kb} KB)")


def main() -> None:
    print(f"Fetching modules for course {COURSE_ID} …")
    modules = list(paginate(f"{CANVAS_BASE}/courses/{COURSE_ID}/modules"))

    if not modules:
        print("No modules found. Verify the course ID and that your token has "
              "the 'url:GET|/api/v1/courses/:course_id/modules' permission.")
        sys.exit(1)

    total_files = 0

    for module in modules:
        name: str = module["name"]
        if MODULE_FILTER and MODULE_FILTER.lower() not in name.lower():
            print(f"\n[skip] {name}")
            continue

        print(f"\n── {name}")
        folder = OUTPUT_DIR / safe_name(name)

        items = list(paginate(
            f"{CANVAS_BASE}/courses/{COURSE_ID}/modules/{module['id']}/items"
        ))

        file_items = [i for i in items if i["type"] == "File"]
        if not file_items:
            print("    (no file items)")
            continue

        for item in file_items:
            download_file(item["content_id"], folder)
            total_files += 1

    print(f"\nFinished — {total_files} file(s) downloaded to {OUTPUT_DIR}/")


if __name__ == "__main__":
    main()
