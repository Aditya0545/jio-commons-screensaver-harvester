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
CATEGORIES = config.DEFAULT_CATEGORIES
TARGET_RESOLUTIONS = getattr(config, "TARGET_RESOLUTION", None)
COMMON_RESOLUTIONS = getattr(config, "COMMON_RESOLUTIONS", [])
ALLOW_COMMON = getattr(config, "ALLOW_COMMON_RESOLUTIONS", False)
MAX_IMAGES = config.MAX_IMAGES
WINNING_KEYWORDS = config.WINNING_KEYWORDS
DEFAULT_OUTPUT = config.DEFAULT_OUTPUT_FILE

# Internal defaults (not in config - kept simple)
TOLERANCE = 0.0  # Exact match only - no tolerance for resolution matching
MIN_SCORE = 50  # Minimum resolution score to accept
MIN_WIDTH = 800
MIN_HEIGHT = 600
MAX_WIDTH = 10000
MAX_HEIGHT = 10000
THUMBNAIL_WIDTH = 1920
THUMBNAIL_WORKERS = 10
API_TIMEOUT = 20


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
    """Extract base_url and page title from URL.
    Supports both Category: and Commons: namespace pages.
    """
    parsed = urlparse(url)
    base_url = f"{parsed.scheme}://{parsed.netloc}"
    
    if "/wiki/" in parsed.path:
        page_title = unquote(parsed.path.split("/wiki/", 1)[1])
    elif "title=" in parsed.query:
        query_params = parse_qs(parsed.query)
        page_title = unquote(query_params.get("title", [""])[0])
    else:
        return None
    
    # Accept both Category: and Commons: namespace pages
    # Category: pages use categorymembers API
    # Commons: pages use images API to get linked images
    if not (page_title.startswith("Category:") or page_title.startswith("Commons:")):
        return None
    
    return base_url, page_title


def is_winning_category(cat_title: str) -> bool:
    """Check if category contains winning/award images."""
    cat_lower = cat_title.lower()
    for keyword in WINNING_KEYWORDS:
        if keyword.lower() in cat_lower:
            return True
    return False


def resolution_score(width: int, height: int, targets: List[Tuple[int, int]], tolerance: float) -> float:
    """Score how well image resolution matches targets. Returns 0-100.
    Only accepts EXACT matches - width and height must match exactly.
    """
    # Check for exact matches only
    for tw, th in targets:
        # Exact match required - width and height must be identical
        if width == tw and height == th:
            return 100
    
    # No match found - return 0
    return 0


def clean_html(text: str) -> str:
    """Remove HTML tags from text."""
    if not text:
        return ""
    text = re.sub(r'<[^>]+>', '', text)
    text = ' '.join(text.split())
    return text.strip()


def is_allowed_format(title: str) -> bool:
    """Check if file format is allowed (jpg, jpeg, png only)."""
    # MediaWiki titles are like "File:example.jpg"
    # Extract extension from title
    title_lower = title.lower()
    
    # Allowed formats
    allowed_extensions = ['.jpg', '.jpeg', '.png']
    
    # Check if title ends with any allowed extension
    for ext in allowed_extensions:
        if title_lower.endswith(ext):
            return True
    
    return False


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
        
        # Filter by file format - only allow jpg, jpeg, png
        if not is_allowed_format(title):
            continue
        
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
    
    data = api_request(base_url, params, timeout=API_TIMEOUT)
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


def load_existing_images(filename: str) -> Tuple[Set[str], int]:
    """Load already harvested image URLs from existing Excel file.
    Returns: (set of existing URLs, highest batch number found)
    """
    if not os.path.exists(filename):
        return set(), 0
    
    try:
        wb = load_workbook(filename, read_only=True)
        ws = wb.active
        
        existing = set()
        max_batch = 0
        
        # Check if batch_number column exists (first column)
        # If file has headers, check row 1 to see column structure
        headers = []
        if ws.max_row >= 1:
            for cell in ws[1]:
                headers.append(cell.value)
        
        # Determine column indices
        # New format: batch_number (col 0), image_page_url (col 1), file_url (col 2)
        # Old format: image_page_url (col 0), file_url (col 1), no batch_number
        has_batch_column = headers and headers[0] == "batch_number"
        
        if has_batch_column:
            # New format: batch_number in col 0, image_page_url in col 1, file_url in col 2
            for row in ws.iter_rows(min_row=2, values_only=True):
                if row and len(row) > 2:
                    batch_val = row[0]
                    file_url = row[2]  # file_url is in column 2 in new format
                    
                    # Track max batch number
                    if batch_val and isinstance(batch_val, (int, float)):
                        max_batch = max(max_batch, int(batch_val))
                    
                    # Track existing URLs
                    if file_url:
                        existing.add(file_url)
        else:
            # Old format: image_page_url in col 0, file_url in col 1
            for row in ws.iter_rows(min_row=2, values_only=True):
                if row and len(row) > 1 and row[1]:
                    existing.add(row[1])  # file_url is in column 1 in old format
            # For old files, start batch from 1
            max_batch = 1 if existing else 0
        
        wb.close()
        return existing, max_batch
    except Exception:
        return set(), 0


def get_subcategories(base_url: str, cat_title: str) -> List[str]:
    """Get all subcategories from a category.
    Returns list of subcategory titles (Category:...).
    """
    if not cat_title.startswith("Category:"):
        return []
    
    params = {
        "action": "query",
        "list": "categorymembers",
        "cmtitle": cat_title,
        "cmtype": "subcat",  # Only get subcategories
        "cmlimit": "500",
        "format": "json",
        "formatversion": "2",
    }
    
    data = api_request(base_url, params)
    if not data:
        return []
    
    members = data.get("query", {}).get("categorymembers", [])
    subcategories = [
        m.get("title") for m in members 
        if isinstance(m, dict) and m.get("title") and m.get("title").startswith("Category:")
    ]
    
    return subcategories


def harvest_from_category(
    base_url: str,
    cat_title: str,
    max_items: int,
    target_sizes: Optional[List[Tuple[int, int]]],
    existing_urls: Set[str],
    is_priority: bool = False,
    processed_cats: Optional[Set[str]] = None,
    depth: int = 0,
    max_depth: int = 2
) -> List[dict]:
    """Harvest from category or Commons page, skipping already collected images.
    Recursively processes subcategories up to max_depth levels.
    Uses categorymembers API for Category: pages, images API for Commons: pages.
    """
    # Track processed categories to avoid infinite loops
    if processed_cats is None:
        processed_cats = set()
    
    # Avoid processing same category twice
    if cat_title in processed_cats:
        return []
    
    processed_cats.add(cat_title)
    
    items = []
    
    # Indent output based on depth for better readability
    indent = "  " * depth
    priority_marker = "🏆 PRIORITY" if is_priority else "📁"
    print(f"\n[*] {priority_marker} {indent}Harvesting: {cat_title[:65]}")
    
    file_titles = []
    
    # Use different API methods based on page type
    if cat_title.startswith("Category:"):
        # Standard category - use categorymembers API
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
        if data:
            members = data.get("query", {}).get("categorymembers", [])
            file_titles = [m.get("title") for m in members if isinstance(m, dict) and m.get("title")]
    elif cat_title.startswith("Commons:"):
        # Commons namespace page - use images API to get linked images
        params = {
            "action": "query",
            "titles": cat_title,
            "prop": "images",
            "imlimit": "500",
            "format": "json",
            "formatversion": "2",
        }
        
        data = api_request(base_url, params)
        if data:
            pages = data.get("query", {}).get("pages", [])
            for page in pages:
                images = page.get("images", [])
                # Filter to only include File: namespace pages (actual image files)
                file_titles = [
                    img.get("title") for img in images 
                    if isinstance(img, dict) and img.get("title") and img.get("title").startswith("File:")
                ]
                break  # Only process first page
    
    # Process files in this category
    if file_titles:
        print(f"    {indent}📄 Found {len(file_titles)} files")
        
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
                print(f"    {indent}{marker} [{len(items)}/{max_items}] {meta['width']}×{meta['height']} - {meta['title'][:40]}")
    
    # Recursively process subcategories if depth limit not reached
    if len(items) < max_items and depth < max_depth and cat_title.startswith("Category:"):
        subcategories = get_subcategories(base_url, cat_title)
        
        if subcategories:
            print(f"    {indent}📂 Found {len(subcategories)} subcategories")
            
            for subcat in subcategories:
                if len(items) >= max_items:
                    break
                
                # Recursively harvest from subcategory
                sub_items = harvest_from_category(
                    base_url,
                    subcat,
                    max_items - len(items),
                    target_sizes,
                    existing_urls,
                    is_priority,
                    processed_cats,
                    depth + 1,
                    max_depth
                )
                items.extend(sub_items)
    
    if items:
        print(f"    {indent}✅ Collected {len(items)} new images from {cat_title[:50]}")
    elif not file_titles and depth == 0:
        print(f"    {indent}⊘ No files found")
    
    return items


def append_to_xlsx(items: List[dict], filename: str, batch_number: int) -> None:
    """Append new images to existing Excel file or create new one.
    Adds batch_number as first column to track which run added each image.
    """
    if not items:
        print("[!] No new items to save")
        return
    
    os.makedirs(os.path.dirname(os.path.abspath(filename)) or ".", exist_ok=True)
    
    file_exists = os.path.exists(filename)
    
    if file_exists:
        wb = load_workbook(filename)
        ws = wb.active
        
        # Check if batch_number column exists
        headers = []
        if ws.max_row >= 1:
            for cell in ws[1]:
                headers.append(cell.value)
        
        has_batch_column = headers and headers[0] == "batch_number"
        
        if not has_batch_column:
            # Old format file - need to add batch_number column
            # Insert new first column for batch_number
            ws.insert_cols(1)
            ws.cell(1, 1, "batch_number").font = Font(bold=True)
            
            # Add batch number 1 to all existing rows
            for row_idx in range(2, ws.max_row + 1):
                ws.cell(row_idx, 1, 1)
            
            print(f"[*] Updated existing file format: added batch_number column")
        
        print(f"[*] Appending batch #{batch_number} to existing file: {filename}")
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Images"
        
        # New format: batch_number is first column
        headers = [
            "batch_number", "image_page_url", "file_url", "line1", "line2",
            "best_thumb_url", "best_res_under_1mb", "title",
            "description", "license", "author"
        ]
        
        ws.append(headers)
        for col_idx in range(1, len(headers) + 1):
            ws.cell(1, col_idx).font = Font(bold=True)
        
        print(f"[*] Creating new file: {filename}")
    
    # Add new items with batch number
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
        
        # Row with batch_number as first column
        row = [
            batch_number,  # Batch number in first column
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
            
            # Hyperlink columns shifted by 1 (was 1,2,5, now 2,3,6)
            if col_idx in [2, 3, 6] and val:
                cell.hyperlink = val
                cell.font = Font(color="0000FF", underline="single")
    
    # Auto-size columns for new files only
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
        print(f"[+] Saved batch #{batch_number} to: {filename}")
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
        if TOLERANCE == 0.0:
            print(f"📐 Resolutions (EXACT MATCH ONLY):")
        else:
            print(f"📐 Resolutions (±{int(TOLERANCE*100)}%):")
        for tw, th in target_sizes[:8]:
            print(f"   • {tw}×{th}")
        if len(target_sizes) > 8:
            print(f"   ... and {len(target_sizes) - 8} more")
    else:
        print(f"📐 Any resolution")
    
    existing_urls, max_batch = load_existing_images(args.output)
    next_batch = max_batch + 1
    
    if existing_urls:
        print(f"📋 Found {len(existing_urls)} existing images (skipping duplicates)")
        print(f"📦 Last batch: #{max_batch}, Starting new batch: #{next_batch}")
    else:
        print(f"📦 Starting batch: #{next_batch}")
    
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
    
    # Track processed categories across all runs to avoid duplicates
    processed_categories = set()
    
    # Process priority categories FIRST
    if priority_cats:
        print(f"\n🏆 PRIORITY: Processing {len(priority_cats)} winning/award categories first")
        
        for base_url, cat_title, cat_url in priority_cats:
            if len(all_items) >= max_items:
                break
            
            remaining = max_items - len(all_items)
            items = harvest_from_category(
                base_url, cat_title, remaining, target_sizes, existing_urls, 
                is_priority=True, processed_cats=processed_categories
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
                base_url, cat_title, remaining, target_sizes, existing_urls, 
                is_priority=False, processed_cats=processed_categories
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
    
    # Save with batch number
    append_to_xlsx(all_items, args.output, next_batch)
    
    print(f"\n{'='*70}")
    print(f"✅ Added {len(all_items)} new images in batch #{next_batch}")
    
    total = len(existing_urls) + len(all_items)
    print(f"📊 Total images in file: {total}")
    print(f"📦 Batch #{next_batch} complete")
    
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