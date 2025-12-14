#!/usr/bin/env python3
"""
Smart WikiMedia Commons Quality Image Harvester - OPTIMIZED FAST VERSION
Fetches high-quality images matching your demo format.

Features:
- Targets Featured/Quality images only
- SMART resolution matching - finds nearby resolutions too
- FAST parallel processing - no waiting
- Exact Excel format from demo
- Smart thumbnail generation (<1MB)
"""

import sys
import os
import argparse
import re
from typing import Optional, List, Dict, Tuple
from urllib.parse import urlparse, unquote, parse_qs
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

import config

# ========== CONFIG ==========
try:
    USER_AGENT = f"Wiki-Jio/1.0 (MediaWiki user: {config.MEDIAWIKI_USERNAME})"
except Exception:
    print("[X] config.py missing. Set MEDIAWIKI_USERNAME")
    sys.exit(1)

if not getattr(config, "MEDIAWIKI_USERNAME", "").strip():
    print("[X] MEDIAWIKI_USERNAME is empty in config.py")
    print("[!] Add your Wikimedia Commons username:")
    print('    MEDIAWIKI_USERNAME = "YourUsername"')
    sys.exit(1)

DEFAULT_CATEGORIES = getattr(config, "DEFAULT_CATEGORIES", [])
TARGET_RESOLUTION = getattr(config, "TARGET_RESOLUTION", None)
MAX_IMAGES = getattr(config, "MAX_IMAGES", 10)

# Popular resolutions to target if user's exact resolution is hard to find
COMMON_RESOLUTIONS = [
    (3840, 2160),  # 4K (16:9)
    (2560, 1440),  # QHD (16:9)
    (1920, 1080),  # Full HD (16:9)
    (1280, 720),   # HD (16:9)
    (1280, 853),   # Common camera resolution (3:2)
    (1920, 1280),  # Common DSLR (3:2)
    (2048, 1365),  # Common photo resolution (3:2)
    (3840, 2400),  # 4K widescreen (16:10)
    (2560, 1600),  # WQXGA (16:10)
    (1920, 1200),  # WUXGA (16:10)
    (4000, 3000),  # 12MP photo (4:3)
    (3000, 2000),  # 6MP photo (3:2)
    (2400, 1600),  # Common photo (3:2)
]


def api_request(base_url: str, params: dict, timeout: int = 20) -> Optional[dict]:
    """Make API request with retries."""
    api_url = f"{base_url}/w/api.php"
    headers = {"User-Agent": USER_AGENT, "Accept": "application/json"}
    
    for attempt in range(2):
        try:
            resp = requests.get(api_url, params=params, headers=headers, timeout=timeout)
            resp.raise_for_status()
            return resp.json()
        except Exception as e:
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


def resolution_score(width: int, height: int, targets: List[Tuple[int, int]], tolerance: float = 0.30) -> float:
    """
    Score how well an image resolution matches target resolutions.
    Returns score 0-100. Higher is better.
    Uses RELAXED 30% tolerance by default for faster results.
    """
    best_score = 0
    
    for tw, th in targets:
        # Calculate % difference
        w_diff = abs(width - tw) / tw
        h_diff = abs(height - th) / th
        
        # If within tolerance, score based on how close
        if w_diff <= tolerance and h_diff <= tolerance:
            score = 100 - (w_diff + h_diff) * 50
            best_score = max(best_score, score)
        
        # Also check if aspect ratio is similar (even if size is different)
        target_aspect = tw / th
        actual_aspect = width / height
        aspect_diff = abs(target_aspect - actual_aspect) / target_aspect
        
        if aspect_diff < 0.1:  # Same aspect ratio
            size_diff = abs(width * height - tw * th) / (tw * th)
            if size_diff < 0.5:  # Similar total pixels
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
        
        # Skip very small or very large images
        if width < 800 or height < 600:
            continue
        if width > 10000 or height > 10000:
            continue
        
        page_url = f"{base_url}/wiki/{title}"
        file_url = info.get("url")
        
        author = ext_value("Artist") or info.get("user") or "Unknown"
        author = clean_html(author)
        
        license_type = ext_value("LicenseShortName") or ext_value("License") or "Unknown"
        license_type = clean_html(license_type)
        
        description = ext_value("ImageDescription") or ext_value("ObjectName") or ""
        description = clean_html(description)
        
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


def get_thumbnail_url(base_url: str, title: str, target_width: int = 1920) -> Optional[dict]:
    """Get thumbnail URL for a specific width."""
    params = {
        "action": "query",
        "titles": title,
        "prop": "imageinfo",
        "iiprop": "url|size",
        "iiurlwidth": str(target_width),
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
    
    with ThreadPoolExecutor(max_workers=10) as executor:
        future_to_title = {
            executor.submit(get_thumbnail_url, base_url, title, 1920): title
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


def harvest_from_category_fast(
    base_url: str,
    cat_title: str,
    max_items: int,
    target_sizes: Optional[List[Tuple[int, int]]] = None,
    tolerance: float = 0.30,
    min_score: float = 50
) -> List[dict]:
    """
    Fast harvesting - only checks direct category members.
    Uses scoring system to accept nearby resolutions.
    """
    items = []
    seen_files = set()
    
    print(f"\n[*] Harvesting: {cat_title[:70]}")
    
    # Get files from this category only (fast!)
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
    
    print(f"    📄 Found {len(file_titles)} files, processing...")
    
    # Process in batches
    for i in range(0, len(file_titles), 50):
        if len(items) >= max_items:
            break
        
        batch = file_titles[i:i+50]
        new_titles = [t for t in batch if t not in seen_files]
        
        if not new_titles:
            continue
        
        for t in new_titles:
            seen_files.add(t)
        
        # Fetch metadata
        batch_meta = fetch_batch_metadata(base_url, new_titles)
        
        # Score and filter
        scored_items = []
        for title, meta in batch_meta.items():
            if len(items) >= max_items:
                break
            
            if target_sizes:
                score = resolution_score(meta["width"], meta["height"], target_sizes, tolerance)
                if score >= min_score:
                    meta["resolution_score"] = score
                    scored_items.append(meta)
            else:
                meta["resolution_score"] = 100
                scored_items.append(meta)
        
        # Sort by score (best matches first)
        scored_items.sort(key=lambda x: x["resolution_score"], reverse=True)
        
        for meta in scored_items:
            if len(items) >= max_items:
                break
            items.append(meta)
            print(f"    ✓ [{len(items)}/{max_items}] {meta['width']}×{meta['height']} (score: {meta['resolution_score']:.0f}) - {meta['title'][:35]}")
    
    print(f"    → Collected {len(items)} images")
    return items


def get_featured_categories_direct() -> List[str]:
    """
    Get direct featured/quality category URLs.
    Much faster than scanning!
    """
    return [
        # Featured Pictures - main categories
        "https://commons.wikimedia.org/wiki/Category:Featured_pictures_on_Wikimedia_Commons",
        
        # Quality Images
        "https://commons.wikimedia.org/wiki/Category:Quality_images",
        
        # Pictures of the Day
        "https://commons.wikimedia.org/wiki/Category:Pictures_of_the_day_(Wikimedia_Commons)",
        
        # Valued Images
        "https://commons.wikimedia.org/wiki/Category:Valued_images",
    ]


def harvest_from_parent_category(
    category_url: str,
    max_items: int,
    target_sizes: Optional[List[Tuple[int, int]]] = None,
    tolerance: float = 0.30
) -> List[dict]:
    """
    Fast harvesting - only direct category files.
    """
    extracted = extract_category_title(category_url)
    if not extracted:
        print(f"[X] Invalid category URL: {category_url}")
        return []
    
    base_url, cat_title = extracted
    
    items = harvest_from_category_fast(base_url, cat_title, max_items, target_sizes, tolerance)
    
    return items


def save_to_xlsx(items: List[dict], filename: str) -> None:
    """Save to Excel with hyperlinks - matching demo format exactly."""
    if not items:
        print("[!] No items to save")
        return
    
    os.makedirs(os.path.dirname(os.path.abspath(filename)) or ".", exist_ok=True)
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Images"
    
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
        "author"
    ]
    
    ws.append(headers)
    for col_idx in range(1, len(headers) + 1):
        ws.cell(1, col_idx).font = Font(bold=True)
    
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
        print(f"\n[+] Saved to: {filename}")
    except Exception as e:
        print(f"[X] Failed to save: {e}")


def normalize_target_resolutions(config_value) -> Optional[List[Tuple[int, int]]]:
    """Normalize TARGET_RESOLUTION from config."""
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


def expand_target_resolutions(targets: List[Tuple[int, int]]) -> List[Tuple[int, int]]:
    """
    Expand target resolutions with nearby common resolutions.
    This helps find images faster!
    """
    expanded = list(targets)
    
    # Add common resolutions that are close to targets
    for common_res in COMMON_RESOLUTIONS:
        if common_res not in expanded:
            expanded.append(common_res)
    
    return expanded


def main():
    parser = argparse.ArgumentParser(description="Fast harvest of quality images from Wikimedia Commons")
    parser.add_argument("url", nargs="?", help="Category URL")
    parser.add_argument("--max", "-m", type=int, help="Max images (default: from config)")
    parser.add_argument("--xlsx", "-x", default="results.xlsx", help="Output file")
    parser.add_argument("--width", type=int, help="Target width")
    parser.add_argument("--height", type=int, help="Target height")
    parser.add_argument("--tolerance", type=float, default=0.30, help="Resolution tolerance (default: 0.30 = ±30%%)")
    parser.add_argument("--strict", action="store_true", help="Use strict resolution matching (slower)")
    
    args = parser.parse_args()
    
    max_items = args.max or MAX_IMAGES or 10
    
    # Handle target resolutions
    target_sizes = None
    if args.width and args.height:
        target_sizes = [(args.width, args.height)]
    else:
        target_sizes = normalize_target_resolutions(TARGET_RESOLUTION)
    
    # Expand with common resolutions for faster results (unless strict mode)
    if target_sizes and not args.strict:
        target_sizes = expand_target_resolutions(target_sizes)
        print(f"[*] Expanded to {len(target_sizes)} target resolutions for faster results")
    
    print("=" * 70)
    print("🎨 WIKIMEDIA QUALITY IMAGE HARVESTER - FAST MODE")
    print("=" * 70)
    print(f"🎯 Max images: {max_items}")
    
    if target_sizes:
        tolerance_pct = args.tolerance * 100
        print(f"📐 Target resolutions (±{tolerance_pct:.0f}% tolerance, accepts nearby):")
        for tw, th in target_sizes[:5]:
            print(f"   • {tw}×{th}")
        if len(target_sizes) > 5:
            print(f"   ... and {len(target_sizes) - 5} more")
    else:
        print(f"📐 Resolution: Any")
    
    print("=" * 70)
    
    # Harvest
    all_items = []
    
    if args.url:
        items = harvest_from_parent_category(args.url, max_items, target_sizes, args.tolerance)
        all_items.extend(items)
    else:
        categories_to_use = DEFAULT_CATEGORIES if DEFAULT_CATEGORIES else None
        
        if not categories_to_use:
            print("[!] No categories in config, using default featured categories...")
            categories_to_use = get_featured_categories_direct()
        
        for cat_url in categories_to_use:
            if len(all_items) >= max_items:
                break
            
            print(f"\n{'─'*70}")
            
            remaining = max_items - len(all_items)
            items = harvest_from_parent_category(cat_url, remaining, target_sizes, args.tolerance)
            all_items.extend(items)
    
    if not all_items:
        print("\n[X] No images found")
        print("\n💡 TIPS:")
        print("   1. Remove resolution filter: TARGET_RESOLUTION = None")
        print("   2. Use --tolerance 0.5 for more relaxed matching")
        print("   3. Try different categories in config.py")
        sys.exit(1)
    
    # Add thumbnails
    print(f"\n{'='*70}")
    print(f"🖼️  Generating thumbnails for {len(all_items)} images...")
    print("=" * 70)
    
    base_url = "https://commons.wikimedia.org"
    titles = [it["title"] for it in all_items]
    thumbs = process_thumbnails_batch(base_url, titles)
    
    for it in all_items:
        thumb = thumbs.get(it["title"])
        if thumb:
            it["best_thumb_url"] = thumb.get("url")
            it["best_thumb_width"] = thumb.get("width")
            it["best_thumb_height"] = thumb.get("height")
    
    print(f"\n{'='*70}")
    print(f"✅ Harvested {len(all_items)} quality images")
    print("=" * 70)
    
    # Resolution breakdown
    res_counts = {}
    for it in all_items:
        res = f"{it['width']}×{it['height']}"
        res_counts[res] = res_counts.get(res, 0) + 1
    
    print(f"\n📊 Resolution breakdown:")
    for res, count in sorted(res_counts.items(), key=lambda x: -x[1])[:10]:
        print(f"   {res}: {count} images")
    
    print(f"\n📋 Sample images:")
    for idx, it in enumerate(all_items[:3], 1):
        print(f"\n{idx}. {it['title'][:60]}")
        print(f"   {it['width']:,}×{it['height']:,} pixels")
        print(f"   {it['author'][:40]}")
        print(f"   {it['license_type']}")
    
    if len(all_items) > 3:
        print(f"\n... and {len(all_items) - 3} more")
    
    # Save
    save_to_xlsx(all_items, args.xlsx)
    
    print(f"\n{'='*70}")
    print(f"✅ Done! Check {args.xlsx}")
    print("=" * 70)


if __name__ == "__main__":
    main()