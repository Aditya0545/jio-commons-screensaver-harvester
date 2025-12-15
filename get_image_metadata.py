#!/usr/bin/env python3
"""
Wikimedia Commons Quality Image Harvester - Set-Top Box Version
Priority-based harvesting: Winning images → Featured → Quality → General
100% config-driven, zero redundancy
"""

import sys
import os
import argparse
import re
from typing import Optional, List, Dict, Tuple, Set
from urllib.parse import urlparse, unquote, parse_qs
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

import config

# ========== LOAD ALL SETTINGS FROM CONFIG ONLY ==========
try:
    USER_AGENT = f"Wiki-Jio/1.0 (MediaWiki user: {config.MEDIAWIKI_USERNAME})"
except Exception:
    print("[X] config.py missing or invalid")
    sys.exit(1)

if not getattr(config, "MEDIAWIKI_USERNAME", "").strip():
    print("[X] MEDIAWIKI_USERNAME is empty in config.py")
    sys.exit(1)

# Everything from config - SINGLE SOURCE OF TRUTH
CATEGORIES = getattr(config, "DEFAULT_CATEGORIES", [])
TARGET_RESOLUTIONS = getattr(config, "TARGET_RESOLUTION", None)
COMMON_RESOLUTIONS = getattr(config, "COMMON_RESOLUTIONS", [])
ALLOW_COMMON = getattr(config, "ALLOW_COMMON_RESOLUTIONS", False)
MAX_IMAGES = getattr(config, "MAX_IMAGES", 50)
TOLERANCE = getattr(config, "TOLERANCE", 0.30)
MIN_SCORE = getattr(config, "MIN_RESOLUTION_SCORE", 50)
WINNING_KEYWORDS = getattr(config, "WINNING_KEYWORDS", [])
DEFAULT_OUTPUT = getattr(config, "DEFAULT_OUTPUT_FILE", "results.xlsx")
API_TIMEOUT = getattr(config, "API_TIMEOUT", 20)
THUMBNAIL_WORKERS = getattr(config, "THUMBNAIL_WORKERS", 10)
THUMBNAIL_WIDTH = getattr(config, "THUMBNAIL_WIDTH", 1920)
MIN_WIDTH = getattr(config, "MIN_WIDTH", 800)
MIN_HEIGHT = getattr(config, "MIN_HEIGHT", 600)
MAX_WIDTH = getattr(config, "MAX_WIDTH", 10000)
MAX_HEIGHT = getattr(config, "MAX_HEIGHT", 10000)


def api_request(base_url: str, params: dict, timeout: int = None) -> Optional[dict]:
    """Make API request with retries."""
    if timeout is None:
        timeout = API_TIMEOUT
    
    api_url = f"{base_url}/w/api.php"
    headers = {"User-Agent": USER_AGENT, "Accept": "application/json"}
    
    for attempt in range(2):
        try:
            resp = requests.get(api_url, params=params, headers=headers, timeout=timeout)
            resp.raise_for_status()
            return resp.json()
        except Exception:
            if attempt == 0:
                continue
            return None


def extract_category_title(url: str) -> Optional[Tuple[str, str]]:
    """Extract base_url and category title from URL."""
    parsed = urlparse(url)
    base_url = f"{parsed.scheme}://{parsed.netloc}"
    
    if "/wiki/" in parsed.path:
        cat_title = unquote(parsed.path.split("/wiki/", 1)[1])
    elif "title=" in parsed.query:
        query_params = parse_qs(parsed.query)
        cat_title = unquote(query_params.get("title", [""])[0])
    else:
        return None
    
    if not cat_title.startswith("Category:"):
        return None
    
    return base_url, cat_title


def is_winning_category(cat_title: str) -> bool:
    """Check if category contains winning/award images."""
    cat_lower = cat_title.lower()
    for keyword in WINNING_KEYWORDS:
        if keyword.lower() in cat_lower:
            return True
    return False


def resolution_score(width: int, height: int, targets: List[Tuple[int, int]], tolerance: float) -> float:
    """Score how well image resolution matches targets. Returns 0-100."""
    best_score = 0
    
    for tw, th in targets:
        w_diff = abs(width - tw) / tw
        h_diff = abs(height - th) / th
        
        if w_diff <= tolerance and h_diff <= tolerance:
            score = 100 - (w_diff + h_diff) * 50
            best_score = max(best_score, score)
        
        target_aspect = tw / th
        actual_aspect = width / height
        aspect_diff = abs(target_aspect - actual_aspect) / target_aspect
        
        if aspect_diff < 0.1:
            size_diff = abs(width * height - tw * th) / (tw * th)
            if size_diff < 0.5:
                score = 60 - size_diff * 30
                best_score = max(best_score, score)
    
    return best_score


def clean_html(text: str) -> str:
    """Remove HTML tags from text."""
    if not text:
        return ""
    text = re.sub(r'<[^>]+>', '', text)
    text = ' '.join(text.split())
    return text.strip()


def fetch_batch_metadata(base_url: str, titles: List[str]) -> Dict[str, dict]:
    """Fetch metadata for up to 50 images in one API call."""
    if not titles:
        return {}
    
    titles_str = "|".join(titles[:50])
    
    params = {
        "action": "query",
        "titles": titles_str,
        "prop": "imageinfo",
        "iiprop": "user|size|url|extmetadata|dimensions",
        "format": "json",
        "formatversion": "2",
    }
    
    data = api_request(base_url, params)
    if not data:
        return {}
    
    pages = data.get("query", {}).get("pages", [])
    result = {}
    
    for page in pages:
        if page.get("missing"):
            continue
        
        imageinfo = page.get("imageinfo", [])
        if not imageinfo:
            continue
        
        info = imageinfo[0]
        title = page.get("title", "")
        
        ext = info.get("extmetadata", {}) or {}
        
        def ext_value(key: str) -> Optional[str]:
            val = ext.get(key)
            if isinstance(val, dict):
                return val.get("value")
            return val
        
        width = info.get("width")
        height = info.get("height")
        
        if not width or not height:
            continue
        
        if width < MIN_WIDTH or height < MIN_HEIGHT or width > MAX_WIDTH or height > MAX_HEIGHT:
            continue
        
        page_url = f"{base_url}/wiki/{title}"
        file_url = info.get("url")
        
        author = clean_html(ext_value("Artist") or info.get("user") or "Unknown")
        license_type = clean_html(ext_value("LicenseShortName") or ext_value("License") or "Unknown")
        description = clean_html(ext_value("ImageDescription") or ext_value("ObjectName") or "")
        
        result[title] = {
            "page_url": page_url,
            "title": title,
            "url": file_url,
            "author": author,
            "size_bytes": info.get("size"),
            "resolution": f"{width}x{height}",
            "license_type": license_type,
            "description": description,
            "width": width,
            "height": height,
        }
    
    return result


def get_thumbnail_url(base_url: str, title: str) -> Optional[dict]:
    """Get thumbnail URL for a specific width."""
    params = {
        "action": "query",
        "titles": title,
        "prop": "imageinfo",
        "iiprop": "url|size",
        "iiurlwidth": str(THUMBNAIL_WIDTH),
        "format": "json",
        "formatversion": "2",
    }
    
    data = api_request(base_url, params, timeout=10)
    if not data:
        return None
    
    pages = data.get("query", {}).get("pages", [])
    if not pages:
        return None
    
    page = pages[0]
    imageinfo = page.get("imageinfo", [])
    if not imageinfo:
        return None
    
    info = imageinfo[0]
    thumb_url = info.get("thumburl") or info.get("url")
    
    return {
        "url": thumb_url,
        "width": info.get("thumbwidth"),
        "height": info.get("thumbheight"),
    }


def process_thumbnails_batch(base_url: str, titles: List[str]) -> Dict[str, dict]:
    """Get thumbnails in parallel."""
    results = {}
    
    with ThreadPoolExecutor(max_workers=THUMBNAIL_WORKERS) as executor:
        future_to_title = {
            executor.submit(get_thumbnail_url, base_url, title): title
            for title in titles
        }
        
        for future in as_completed(future_to_title):
            title = future_to_title[future]
            try:
                thumb = future.result()
                if thumb:
                    results[title] = thumb
            except Exception:
                pass
    
    return results


def load_existing_images(filename: str) -> Set[str]:
    """Load already harvested image URLs from existing Excel file."""
    if not os.path.exists(filename):
        return set()
    
    try:
        wb = load_workbook(filename, read_only=True)
        ws = wb.active
        
        existing = set()
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row and row[1]:
                existing.add(row[1])
        
        wb.close()
        return existing
    except Exception:
        return set()


def harvest_from_category(
    base_url: str,
    cat_title: str,
    max_items: int,
    target_sizes: Optional[List[Tuple[int, int]]],
    existing_urls: Set[str],
    is_priority: bool = False
) -> List[dict]:
    """Harvest from category, skipping already collected images."""
    items = []
    
    priority_marker = "🏆 PRIORITY" if is_priority else "📁"
    print(f"\n[*] {priority_marker} Harvesting: {cat_title[:65]}")
    
    params = {
        "action": "query",
        "list": "categorymembers",
        "cmtitle": cat_title,
        "cmtype": "file",
        "cmlimit": "500",
        "format": "json",
        "formatversion": "2",
    }
    
    data = api_request(base_url, params)
    if not data:
        return []
    
    members = data.get("query", {}).get("categorymembers", [])
    file_titles = [m.get("title") for m in members if isinstance(m, dict) and m.get("title")]
    
    if not file_titles:
        print(f"    ⊘ No files in this category")
        return []
    
    print(f"    📄 Found {len(file_titles)} files")
    
    for i in range(0, len(file_titles), 50):
        if len(items) >= max_items:
            break
        
        batch = file_titles[i:i+50]
        batch_meta = fetch_batch_metadata(base_url, batch)
        
        scored_items = []
        for title, meta in batch_meta.items():
            if len(items) >= max_items:
                break
            
            if meta["url"] in existing_urls:
                continue
            
            if target_sizes:
                score = resolution_score(meta["width"], meta["height"], target_sizes, TOLERANCE)
                if score >= MIN_SCORE:
                    meta["resolution_score"] = score
                    scored_items.append(meta)
            else:
                meta["resolution_score"] = 100
                scored_items.append(meta)
        
        scored_items.sort(key=lambda x: x["resolution_score"], reverse=True)
        
        for meta in scored_items:
            if len(items) >= max_items:
                break
            items.append(meta)
            marker = "🏆" if is_priority else "✓"
            print(f"    {marker} [{len(items)}/{max_items}] {meta['width']}×{meta['height']} - {meta['title'][:40]}")
    
    if items:
        print(f"    ✅ Collected {len(items)} new images")
    
    return items


def append_to_xlsx(items: List[dict], filename: str) -> None:
    """Append new images to existing Excel file or create new one."""
    if not items:
        print("[!] No new items to save")
        return
    
    os.makedirs(os.path.dirname(os.path.abspath(filename)) or ".", exist_ok=True)
    
    file_exists = os.path.exists(filename)
    
    if file_exists:
        wb = load_workbook(filename)
        ws = wb.active
        print(f"[*] Appending to existing file: {filename}")
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Images"
        
        headers = [
            "image_page_url", "file_url", "line1", "line2",
            "best_thumb_url", "best_res_under_1mb", "title",
            "description", "license", "author"
        ]
        
        ws.append(headers)
        for col_idx in range(1, len(headers) + 1):
            ws.cell(1, col_idx).font = Font(bold=True)
        
        print(f"[*] Creating new file: {filename}")
    
    for it in items:
        author = it.get("author", "")
        license_type = it.get("license_type", "")
        desc = it.get("description", "")
        
        line1 = f"by {author}, {license_type}"
        line2 = desc[:64] if desc else ""
        
        thumb_res = ""
        if it.get("best_thumb_width"):
            w = it['best_thumb_width']
            h = it['best_thumb_height']
            thumb_res = f"{w:,} × {h:,} pixels"
        
        row = [
            it.get("page_url", ""),
            it.get("url", ""),
            line1,
            line2,
            it.get("best_thumb_url", ""),
            thumb_res,
            it.get("title", ""),
            desc,
            license_type,
            author
        ]
        
        row_idx = ws.max_row + 1
        
        for col_idx, val in enumerate(row, 1):
            cell = ws.cell(row_idx, col_idx, val)
            
            if col_idx in [1, 2, 5] and val:
                cell.hyperlink = val
                cell.font = Font(color="0000FF", underline="single")
    
    if not file_exists:
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                try:
                    val = str(cell.value or "")
                except:
                    val = ""
                max_len = max(max_len, len(val))
            ws.column_dimensions[col_letter].width = min(max_len + 2, 100)
    
    try:
        wb.save(filename)
        print(f"[+] Saved to: {filename}")
    except Exception as e:
        print(f"[X] Failed to save: {e}")


def normalize_resolutions(config_value) -> Optional[List[Tuple[int, int]]]:
    """Normalize resolution config to list of tuples."""
    if config_value is None:
        return None
    
    if isinstance(config_value, list):
        result = []
        for item in config_value:
            if isinstance(item, (list, tuple)) and len(item) == 2:
                result.append((int(item[0]), int(item[1])))
        return result if result else None
    
    if isinstance(config_value, tuple) and len(config_value) == 2:
        return [(int(config_value[0]), int(config_value[1]))]
    
    return None


def main():
    parser = argparse.ArgumentParser(description="Wikimedia Commons Image Harvester for Set-Top Box")
    parser.add_argument("--output", "-o", default=DEFAULT_OUTPUT, help=f"Output Excel file (default: {DEFAULT_OUTPUT})")
    parser.add_argument("--max", "-m", type=int, help="Override MAX_IMAGES from config")
    
    args = parser.parse_args()
    
    max_items = args.max if args.max else MAX_IMAGES
    target_sizes = normalize_resolutions(TARGET_RESOLUTIONS)
    
    # Add common resolutions if allowed
    if target_sizes and ALLOW_COMMON and COMMON_RESOLUTIONS:
        common = normalize_resolutions(COMMON_RESOLUTIONS)
        if common:
            for res in common:
                if res not in target_sizes:
                    target_sizes.append(res)
    
    if not CATEGORIES:
        print("[X] No categories configured in config.py")
        sys.exit(1)
    
    print("=" * 70)
    print("🎨 WIKIMEDIA QUALITY IMAGE HARVESTER - SET-TOP BOX")
    print("=" * 70)
    print(f"🎯 Target: {max_items} images")
    print(f"📁 Output: {args.output}")
    
    if target_sizes:
        print(f"📐 Resolutions (±{int(TOLERANCE*100)}%):")
        for tw, th in target_sizes[:8]:
            print(f"   • {tw}×{th}")
        if len(target_sizes) > 8:
            print(f"   ... and {len(target_sizes) - 8} more")
    else:
        print(f"📐 Any resolution")
    
    existing_urls = load_existing_images(args.output)
    if existing_urls:
        print(f"📋 Found {len(existing_urls)} existing images (skipping duplicates)")
    
    print("=" * 70)
    
    # Harvest with PRIORITY
    all_items = []
    
    # Separate categories by priority
    priority_cats = []
    regular_cats = []
    
    for cat_url in CATEGORIES:
        extracted = extract_category_title(cat_url)
        if not extracted:
            continue
        
        base_url, cat_title = extracted
        
        if is_winning_category(cat_title):
            priority_cats.append((base_url, cat_title, cat_url))
        else:
            regular_cats.append((base_url, cat_title, cat_url))
    
    # Process priority categories FIRST
    if priority_cats:
        print(f"\n🏆 PRIORITY: Processing {len(priority_cats)} winning/award categories first")
        
        for base_url, cat_title, cat_url in priority_cats:
            if len(all_items) >= max_items:
                break
            
            remaining = max_items - len(all_items)
            items = harvest_from_category(
                base_url, cat_title, remaining, target_sizes, existing_urls, is_priority=True
            )
            all_items.extend(items)
    
    # Process regular categories if needed
    if len(all_items) < max_items and regular_cats:
        print(f"\n📁 Processing {len(regular_cats)} featured/quality categories")
        
        for base_url, cat_title, cat_url in regular_cats:
            if len(all_items) >= max_items:
                break
            
            remaining = max_items - len(all_items)
            items = harvest_from_category(
                base_url, cat_title, remaining, target_sizes, existing_urls, is_priority=False
            )
            all_items.extend(items)
    
    if not all_items:
        print("\n[!] No new images found")
        if existing_urls:
            print(f"[*] Already have {len(existing_urls)} images in {args.output}")
        sys.exit(0)
    
    # Add thumbnails
    print(f"\n🖼️  Generating thumbnails for {len(all_items)} images...")
    
    base_url = "https://commons.wikimedia.org"
    titles = [it["title"] for it in all_items]
    thumbs = process_thumbnails_batch(base_url, titles)
    
    for it in all_items:
        thumb = thumbs.get(it["title"])
        if thumb:
            it["best_thumb_url"] = thumb.get("url")
            it["best_thumb_width"] = thumb.get("width")
            it["best_thumb_height"] = thumb.get("height")
    
    # Save
    append_to_xlsx(all_items, args.output)
    
    print(f"\n{'='*70}")
    print(f"✅ Added {len(all_items)} new images")
    
    total = len(existing_urls) + len(all_items)
    print(f"📊 Total images in file: {total}")
    
    # Stats
    priority_count = sum(1 for it in all_items if it.get("resolution_score", 0) >= 80)
    print(f"🏆 High-quality matches: {priority_count}")
    
    res_counts = {}
    for it in all_items:
        res = f"{it['width']}×{it['height']}"
        res_counts[res] = res_counts.get(res, 0) + 1
    
    if res_counts:
        print(f"\n📐 New images by resolution:")
        for res, count in sorted(res_counts.items(), key=lambda x: -x[1])[:5]:
            print(f"   {res}: {count} images")
    
    print("=" * 70)


if __name__ == "__main__":
    main()