#!/usr/bin/env python3
"""
Fetch image metadata from a MediaWiki site and save CSV + XLSX (with clickable hyperlinks).

Requirements:
    pip install requests tqdm openpyxl

Usage examples:
    python get_image_metadata.py
    python get_image_metadata.py --csv results.csv --xlsx results.xlsx
    python get_image_metadata.py "https://commons.wikimedia.org/wiki/Category:Some_Category" --max 50 --xlsx winners.xlsx

Notes:
- This script requires a non-empty MEDIAWIKI_USERNAME in config.py (used to build a compliant User-Agent).
- By design the script will refuse to run if MEDIAWIKI_USERNAME is empty.
"""

import sys
import os
import csv
import argparse
import time
from typing import Optional, List
from datetime import datetime, timezone
from urllib.parse import urlparse, unquote, parse_qs

import requests
from tqdm import tqdm
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

import config

# ========== CONFIG ==========
# Script expects config.py to define at least:
# MEDIAWIKI_USERNAME (non-empty string), TARGET_RESOLUTION (tuple or None), MAX_IMAGES (int) and
# DEFAULT_CATEGORIES (list of category URLs)

try:
    USER_AGENT = f"Wiki-Jio/1.0 (MediaWiki user: {config.MEDIAWIKI_USERNAME})"
except Exception:
    print("[X] config.py is missing required attributes. Please set MEDIAWIKI_USERNAME, etc.")
    sys.exit(1)

# Enforce username presence as a good-practice requirement (non-empty)
if not getattr(config, "MEDIAWIKI_USERNAME", "") or not str(config.MEDIAWIKI_USERNAME).strip():
    print("[X] MEDIAWIKI_USERNAME in config.py is empty. Please set it before running the script.")
    sys.exit(1)

DEFAULT_CATEGORIES = getattr(config, "DEFAULT_CATEGORIES", [])
TARGET_RESOLUTION = getattr(config, "TARGET_RESOLUTION", None)
MAX_IMAGES = getattr(config, "MAX_IMAGES", 0)

# keywords to identify 'winning' images (case-insensitive). Add or edit in code if you need different rules.
WINNER_KEYWORDS = [
    "winner",
    "winners",
    "first prize",
    "first_prize",
    "gold",
    "second prize",
    "second_prize",
    "third prize",
    "honourable mention",
    "honorable mention",
]

# ========== CORE LOGIC ==========

def extract_file_title(image_url: str) -> Optional[tuple[str, str]]:
    parsed = urlparse(image_url)
    base_url = f"{parsed.scheme}://{parsed.netloc}"
    if "/wiki/" in parsed.path:
        maybe_title = parsed.path.split("/wiki/", 1)[1]
        file_title = unquote(maybe_title)
    elif "title=" in parsed.query:
        query_params = parse_qs(parsed.query)
        file_title = unquote(query_params.get("title", [""])[0])
    else:
        parts = parsed.path.split("/")
        file_title = unquote(parts[-1]) if parts else ""
    if not file_title:
        return None
    return base_url, file_title


def extract_category_title(category_url: str) -> Optional[tuple[str, str]]:
    parsed = urlparse(category_url)
    base_url = f"{parsed.scheme}://{parsed.netloc}"
    if "/wiki/" in parsed.path:
        maybe_title = parsed.path.split("/wiki/", 1)[1]
        cat_title = unquote(maybe_title)
    elif "title=" in parsed.query:
        query_params = parse_qs(parsed.query)
        cat_title = unquote(query_params.get("title", [""])[0])
    else:
        cat_title = ""
    if not cat_title or not cat_title.startswith("Category:"):
        return None
    return base_url, cat_title


def api_request(base_url: str, params: dict) -> Optional[dict]:
    api_url = f"{base_url}/w/api.php"
    headers = {"User-Agent": USER_AGENT, "Accept": "application/json"}
    try:
        resp = requests.get(api_url, params=params, headers=headers, timeout=15)
        resp.raise_for_status()
        return resp.json()
    except Exception as error:
        print(f"[X] API request failed: {error}")
        return None


def _fetch_image_name_and_url(base_url: str, title: str) -> Optional[dict]:
    params = {
        "action": "query",
        "titles": title,
        "prop": "imageinfo",
        "iiprop": "url",
        "format": "json",
        "formatversion": "2",
        "redirects": "true",
        "converttitles": "true",
    }
    data = api_request(base_url, params)
    if not data:
        return None
    pages = data.get("query", {}).get("pages", [])
    if not pages:
        return None
    page = pages[0]
    if page.get("missing"):
        return None
    imageinfo = page.get("imageinfo", [])
    if not imageinfo:
        return None
    info = imageinfo[0]
    return {"title": page.get("title", title), "url": info.get("url")}


def _get_image_dimensions(base_url: str, title: str) -> Optional[tuple[int, int]]:
    params = {
        "action": "query",
        "titles": title,
        "prop": "imageinfo",
        "iiprop": "dimensions",
        "format": "json",
        "formatversion": "2",
        "redirects": "true",
        "converttitles": "true",
    }
    data = api_request(base_url, params)
    if not data:
        return None
    pages = data.get("query", {}).get("pages", [])
    if not pages:
        return None
    page = pages[0]
    if page.get("missing"):
        return None
    imageinfo = page.get("imageinfo", [])
    if not imageinfo:
        return None
    info = imageinfo[0]
    width = info.get("width"); height = info.get("height")
    if width and height:
        return (width, height)
    return None


def get_content_length(url: str, session: requests.Session) -> Optional[int]:
    try:
        resp = session.head(url, allow_redirects=True, timeout=10, headers={"User-Agent": USER_AGENT})
        if resp.status_code == 200:
            cl = resp.headers.get("content-length")
            if cl is not None:
                return int(cl)
        return None
    except Exception:
        return None


def select_best_thumbnail(base_url: str, title: str) -> Optional[dict]:
    candidate_widths = [2560, 1920, 1280, 1024, 800, 640, 512, 320, 256]
    session = requests.Session()
    best = None
    for w in candidate_widths:
        params = {
            "action": "query",
            "titles": title,
            "prop": "imageinfo",
            "iiprop": "url|size",
            "iiurlwidth": str(w),
            "format": "json",
            "formatversion": "2",
            "redirects": "true",
            "converttitles": "true",
        }
        data = api_request(base_url, params)
        if not data:
            continue
        pages = data.get("query", {}).get("pages", [])
        if not pages:
            continue
        page = pages[0]
        if page.get("missing"):
            continue
        imageinfo = page.get("imageinfo", [])
        if not imageinfo:
            continue
        info = imageinfo[0]
        thumb_url = info.get("thumburl") or info.get("url")
        thumb_w = info.get("thumbwidth")
        thumb_h = info.get("thumbheight")
        if not thumb_url:
            continue
        size = get_content_length(thumb_url, session)
        if size is None:
            try:
                r = session.get(thumb_url, stream=True, timeout=15, headers={"User-Agent": USER_AGENT})
                total = 0
                for chunk in r.iter_content(chunk_size=8192):
                    if not chunk:
                        break
                    total += len(chunk)
                    if total > 1_200_000:
                        break
                size = total
                r.close()
            except Exception:
                size = None
        if size is not None and size <= 1_000_000:
            best = {"url": thumb_url, "size": size, "width": thumb_w, "height": thumb_h}
            break
        time.sleep(0.05)
    session.close()
    return best


def is_winner(metadata: dict) -> bool:
    # Determine whether an image is a contest "winning" image by searching category names and title
    cats = metadata.get("categories") or []
    title = (metadata.get("title") or "").lower()
    # check categories
    for c in cats:
        if not isinstance(c, str):
            continue
        cl = c.lower()
        for kw in WINNER_KEYWORDS:
            if kw in cl:
                return True
    # check title
    for kw in WINNER_KEYWORDS:
        if kw in title:
            return True
    return False


def _fetch_image_metadata(base_url: str, title: str) -> Optional[dict]:
    params = {
        "action": "query", "titles": title, "prop": "imageinfo|categories",
        "iiprop": "timestamp|user|size|url|metadata|extmetadata|dimensions",
        "format": "json", "formatversion": "2", "redirects": "true", "converttitles": "true", "cllimit": "50",
    }
    data = api_request(base_url, params)
    if not data:
        return None
    pages = data.get("query", {}).get("pages", [])
    if not pages:
        return None
    page = pages[0]
    if page.get("missing"):
        return None
    imageinfo = page.get("imageinfo", [])
    if not imageinfo:
        return None
    info = imageinfo[0]
    ext = info.get("extmetadata", {}) or {}
    def ext_value(key: str) -> Optional[str]:
        val = ext.get(key)
        if isinstance(val, dict):
            return val.get("value")
        return val
    title = page.get("title", title)
    page_url = f"{base_url}/wiki/{title}"
    file_url = info.get("url")
    author = ext_value("Artist") or info.get("user")
    size_bytes = info.get("size")
    license_type = ext_value("LicenseShortName") or ext_value("License")
    description = ext_value("ImageDescription") or ext_value("ObjectName")
    creation_date = ext_value("DateTimeOriginal") or ext_value("DateTime")
    width = info.get("width"); height = info.get("height")
    categories_raw = page.get("categories", []) or []
    categories = [c.get("title") for c in categories_raw if isinstance(c, dict)]

    best_thumb = select_best_thumbnail(base_url, title)

    return {
        "page_url": page_url,
        "title": title,
        "url": file_url,
        "author": author,
        "size_bytes": size_bytes,
        "resolution": f"{width}x{height}" if width and height else None,
        "license_type": license_type,
        "description": description,
        "creation_date": creation_date,
        "categories": categories,
        "width": width,
        "height": height,
        "best_thumb_url": best_thumb.get("url") if best_thumb else None,
        "best_thumb_size": best_thumb.get("size") if best_thumb else None,
        "best_thumb_width": best_thumb.get("width") if best_thumb else None,
        "best_thumb_height": best_thumb.get("height") if best_thumb else None,
    }


def get_image_metadata(image_url: str) -> Optional[dict]:
    extracted = extract_file_title(image_url)
    if not extracted:
        print("[X] Could not parse file title from URL.")
        return None
    base_url, file_title = extracted
    return _fetch_image_metadata(base_url, file_title)


def fetch_category_files_recursive(category_url: str, max_items: int | None = None, target_size: tuple[int,int] | None = None) -> list[dict]:
    extracted = extract_category_title(category_url)
    if not extracted:
        print("[X] Could not parse category title from URL.")
        return []
    base_url, category_title = extracted
    seen_cats: set[str] = set(); seen_files: set[str] = set(); items: list[dict] = []
    def fetch_files(file_titles: list[str], depth: int) -> None:
        if not file_titles:
            return
        indent = "  " * depth
        print(f"{indent}  [*] Fetching {len(file_titles)} image(s)...")
        for title in file_titles:
            if title in seen_files:
                continue
            seen_files.add(title)
            # fetch full metadata (needed to determine winners and thumbnails)
            image_meta = _fetch_image_metadata(base_url, title)
            if not image_meta:
                continue
            # only include images identified as winners
            if not is_winner(image_meta):
                continue
            if target_size:
                iw = image_meta.get("width")
                ih = image_meta.get("height")
                if iw is None or ih is None:
                    continue
                tw, th = target_size
                if int(iw) != int(tw) or int(ih) != int(th):
                    continue
            items.append(image_meta)
            if max_items is not None and len(items) >= max_items:
                return
    def walk(cat_title: str, depth: int = 0) -> None:
        if cat_title in seen_cats:
            return
        seen_cats.add(cat_title)
        indent = "  " * depth
        print(f"{indent}[*] Scanning category: {cat_title}")
        cont: dict | None = None
        local_files: list[str] = []
        while True:
            params = {
                "action": "query", "list": "categorymembers", "cmtitle": cat_title,
                "cmlimit": "100", "cmtype": "file|subcat", "format": "json", "formatversion": "2",
            }
            if cont:
                params.update(cont)
            data = api_request(base_url, params)
            if not data:
                break
            members = data.get("query", {}).get("categorymembers", []) or []
            for m in members:
                if not isinstance(m, dict):
                    continue
                mtitle = m.get("title")
                if not mtitle:
                    continue
                mtype = m.get("type")
                if mtype is None:
                    ns = m.get("ns")
                    if ns == 14:
                        mtype = "subcat"
                    elif ns == 6:
                        mtype = "file"
                if mtype == "subcat":
                    walk(mtitle, depth + 1)
                    if max_items is not None and len(items) >= max_items:
                        return
                elif mtype == "file":
                    local_files.append(mtitle)
            cont = data.get("continue")
            if not cont:
                break
            if max_items is not None and len(items) >= max_items:
                break
        if local_files:
            fetch_files(local_files, depth)
            if max_items is not None and len(items) >= max_items:
                return
    walk(category_title)
    return items

# ========== OUTPUT HELPERS ==========

def display_metadata(metadata: dict) -> None:
    print("=" * 60)
    print("WINNING IMAGE (file name and URL)")
    print("=" * 60)
    print(f"File Name:   {metadata.get('title', 'Unknown')}")
    print(f"URL:         {metadata.get('url', 'Unknown')}")


def save_items_to_csv(items: list[dict], filename: str) -> None:
    if not items:
        print("[!] No items to save to CSV.")
        return
    os.makedirs(os.path.dirname(os.path.abspath(filename)) or ".", exist_ok=True)
    try:
        with open(filename, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow([
                "image_page_url",
                "file_url",
                "line1",
                "line2",
                "best_thumb_url",
                "best_res_under_1mb",
                "title",
                "description",
                "license",
                "author",
            ])
            for it in items:
                file_url = it.get("url") or ""
                page_url = it.get("page_url") or ""
                author = it.get("author") or ""
                license_type = it.get("license_type") or ""
                desc = (it.get("description") or "")
                line1 = f"by {author}, {license_type}" if author or license_type else ""
                line2 = (desc[:64]).replace("\n", " ")
                best_thumb = it.get("best_thumb_url") or ""
                best_res = f"{it.get('best_thumb_width') or ''} × {it.get('best_thumb_height') or ''} pixels" if it.get('best_thumb_width') else ""
                writer.writerow([
                    page_url,
                    file_url,
                    line1,
                    line2,
                    best_thumb,
                    best_res,
                    it.get("title") or "",
                    desc,
                    license_type,
                    author,
                ])
        print(f"[+] Saved {len(items)} rows to CSV: {filename}")
    except Exception as e:
        print(f"[X] Failed to write CSV: {e}")


def save_items_to_xlsx_with_hyperlinks(items: list[dict], filename: str) -> None:
    if not items:
        print("[!] No items to save to XLSX.")
        return
    os.makedirs(os.path.dirname(os.path.abspath(filename)) or ".", exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "images"
    headers = [
        "image_page_url",
        "file_url",
        "line1",
        "line2",
        "best_thumb_url",
        "best_res_under_1mb",
        "title",
        "description",
        "license",
        "author",
    ]
    header_font = Font(bold=True)
    ws.append(headers)
    for col_idx, _ in enumerate(headers, start=1):
        ws.cell(row=1, column=col_idx).font = header_font
    for it in items:
        file_url = it.get("url") or ""
        page_url = it.get("page_url") or ""
        author = it.get("author") or ""
        license_type = it.get("license_type") or ""
        desc = (it.get("description") or "")
        line1 = f"by {author}, {license_type}" if author or license_type else ""
        line2 = (desc[:64]).replace("\n", " ")
        best_thumb = it.get("best_thumb_url") or ""
        best_res = f"{it.get('best_thumb_width') or ''} × {it.get('best_thumb_height') or ''} pixels" if it.get('best_thumb_width') else ""
        row_idx = ws.max_row + 1
        cell = ws.cell(row=row_idx, column=1, value=page_url)
        if page_url:
            cell.hyperlink = page_url
            cell.font = Font(color="0000FF", underline="single")
        cell = ws.cell(row=row_idx, column=2, value=file_url)
        if file_url:
            cell.hyperlink = file_url
            cell.font = Font(color="0000FF", underline="single")
        ws.cell(row=row_idx, column=3, value=line1)
        ws.cell(row=row_idx, column=4, value=line2)
        cell = ws.cell(row=row_idx, column=5, value=best_thumb)
        if best_thumb:
            cell.hyperlink = best_thumb
            cell.font = Font(color="0000FF", underline="single")
        ws.cell(row=row_idx, column=6, value=best_res)
        ws.cell(row=row_idx, column=7, value=it.get("title") or "")
        ws.cell(row=row_idx, column=8, value=desc)
        ws.cell(row=row_idx, column=9, value=license_type)
        ws.cell(row=row_idx, column=10, value=author)
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                val = str(cell.value) if cell.value is not None else ""
            except Exception:
                val = ""
            if len(val) > max_len:
                max_len = len(val)
        adjusted = (max_len + 2)
        ws.column_dimensions[col_letter].width = adjusted if adjusted < 100 else 100
    try:
        wb.properties.creator = USER_AGENT
        wb.properties.created = datetime.now(timezone.utc)
    except Exception:
        pass
    try:
        wb.save(filename)
        print(f"[+] Saved {len(items)} rows to XLSX: {filename}")
    except PermissionError:
        print(f"[X] Failed to write XLSX: Permission denied")
        print(f"[!] The file '{filename}' is likely open in Excel or another program.")
    except Exception as e:
        print(f"[X] Failed to write XLSX: {e}")

# ========== DOWNLOAD (unchanged) ==========

def download_image(url: str, dst_path: str, session: requests.Session, user_agent: str) -> bool:
    headers = {"User-Agent": user_agent}
    try:
        with session.get(url, headers=headers, stream=True, timeout=30) as r:
            r.raise_for_status()
            total = int(r.headers.get("content-length") or 0)
            tmp_path = dst_path + ".part"
            with open(tmp_path, "wb") as f, tqdm(total=total, unit="B", unit_scale=True, desc=os.path.basename(dst_path), leave=True) as pbar:
                for chunk in r.iter_content(chunk_size=8192):
                    if not chunk:
                        continue
                    f.write(chunk)
                    pbar.update(len(chunk))
            os.replace(tmp_path, dst_path)
            return True
    except Exception as ex:
        print(f"[X] Failed to download {url}: {ex}")
        try:
            if os.path.exists(dst_path + ".part"):
                os.remove(dst_path + ".part")
        except Exception:
            pass
        return False


def download_images(items: List[dict], out_dir: str, user_agent: str) -> None:
    if not items:
        print("[!] No items to download.")
        return
    os.makedirs(out_dir, exist_ok=True)
    session = requests.Session()
    for it in items:
        url = it.get("url")
        if not url:
            print("[!] Skipping item with no URL.")
            continue
        raw_title = it.get("title") or os.path.basename(urlparse(url).path)
        filename = raw_title.replace("File:", "").replace("/", "_")
        if not os.path.splitext(filename)[1]:
            ext = os.path.splitext(urlparse(url).path)[1]
            filename = filename + (ext or ".jpg")
        dst_path = os.path.join(out_dir, filename)
        base, ext = os.path.splitext(dst_path)
        counter = 1
        while os.path.exists(dst_path):
            dst_path = f"{base}_{counter}{ext}"
            counter += 1
        success = download_image(url, dst_path, session, user_agent)
        if not success:
            print(f"[!] Retry once for {url}")
            time.sleep(1)
            success = download_image(url, dst_path, session, user_agent)
            if not success:
                print(f"[X] Skipped {url}")
    session.close()
    print("[+] Download completed.")

# ========== CLI ==========

def main() -> None:
    parser = argparse.ArgumentParser(description="Fetch MediaWiki image metadata (anonymous).")
    parser.add_argument("url", nargs="?", help="Image or Category URL (if omitted uses DEFAULT_CATEGORIES from config)")
    parser.add_argument("--csv", "-c", metavar="OUTFILE", help="Save results to CSV (e.g., --csv results.csv)")
    parser.add_argument("--xlsx", "-x", metavar="OUTXLSX", help="Save results to XLSX (e.g., --xlsx results.xlsx)")
    parser.add_argument("--max", "-m", type=int, default=15, help="Max items for category recursive fetch")
    parser.add_argument("--width", type=int, help="Target width to filter images (optional)")
    parser.add_argument("--height", type=int, help="Target height to filter images (optional)")
    parser.add_argument("--download", "-d", metavar="DIR", help="Download image files into DIR (e.g., --download images/)")
    args = parser.parse_args()

    # Determine effective max: CLI overrides config only if provided non-default.
    effective_max = None
    if args.max is not None and args.max != 15:
        effective_max = args.max
    else:
        if isinstance(MAX_IMAGES, int) and MAX_IMAGES > 0:
            effective_max = MAX_IMAGES

    # Determine effective target_size: CLI (width & height) preferred; otherwise config.TARGET_RESOLUTION
    target_size = None
    if args.width and args.height:
        target_size = (args.width, args.height)
    else:
        if TARGET_RESOLUTION and isinstance(TARGET_RESOLUTION, (list, tuple)) and len(TARGET_RESOLUTION) == 2:
            target_size = (int(TARGET_RESOLUTION[0]), int(TARGET_RESOLUTION[1]))

    items: List[dict] = []

    if args.url:
        image_url = args.url
        if "Category:" in image_url:
            items = fetch_category_files_recursive(category_url=image_url, max_items=effective_max, target_size=target_size)
        else:
            metadata = get_image_metadata(image_url)
            if metadata and is_winner(metadata):
                items = [metadata]
            else:
                print("[X] Failed to fetch winning image metadata or image is not marked as a winner.")
                sys.exit(1)
    else:
        if not DEFAULT_CATEGORIES:
            print("[X] No URL provided and config.DEFAULT_CATEGORIES is empty. Please set DEFAULT_CATEGORIES in config.py or provide a URL.")
            sys.exit(1)
        remaining = None if effective_max is None else effective_max
        for cat in DEFAULT_CATEGORIES:
            if remaining == 0:
                break
            cat_max = None if remaining is None else remaining
            found = fetch_category_files_recursive(category_url=cat, max_items=cat_max, target_size=target_size)
            for f in found:
                items.append(f)
                if remaining is not None:
                    remaining -= 1
                    if remaining <= 0:
                        break
            if remaining is not None and remaining <= 0:
                break

    if not items:
        print("[X] No winning images found (after searching).")
        sys.exit(1)

    for it in items:
        display_metadata(it)

    if args.csv:
        save_items_to_csv(items, args.csv)

    xlsx_path = args.xlsx if args.xlsx else os.path.join(os.getcwd(), "results.xlsx")
    save_items_to_xlsx_with_hyperlinks(items, xlsx_path)

    if args.download:
        download_images(items, args.download, USER_AGENT)

if __name__ == "__main__":
    main()