"""
Helper to fetch image metadata from a MediaWiki site anonymously (no login).

Usage examples:
    python get_image_metadata.py "https://commons.wikimedia.org/wiki/File:Example.jpg"
    python get_image_metadata.py
"""

import sys
from typing import Optional
from urllib.parse import urlparse, unquote, parse_qs

import requests

# =============================================================================
# CONFIGURATION
# =============================================================================

# Optional default URL. Set to the target category so the script runs without prompt.
DEFAULT_IMAGE_URL = (
    "https://commons.wikimedia.org/wiki/Category:"
    "Winners_of_Wiki_Loves_Monuments_2024_by_country"
)


# =============================================================================
# CORE LOGIC
# =============================================================================


def extract_file_title(image_url: str) -> Optional[tuple[str, str]]:
    """
    Extract base site URL and file title from a MediaWiki image/file URL.

    Keeps the logic explicit and commented so the caller understands how the
    title is derived. Returns (base_url, file_title) or None on failure.
    """
    parsed = urlparse(image_url)
    base_url = f"{parsed.scheme}://{parsed.netloc}"

    # Try common /wiki/File:... format first.
    if "/wiki/" in parsed.path:
        maybe_title = parsed.path.split("/wiki/", 1)[1]
        file_title = unquote(maybe_title)
    # Fallback: look for title=File:... in query string (old format).
    elif "title=" in parsed.query:
        query_params = parse_qs(parsed.query)
        file_title = unquote(query_params.get("title", [""])[0])
    else:
        # Last path segment as a last resort.
        parts = parsed.path.split("/")
        file_title = unquote(parts[-1]) if parts else ""

    if not file_title:
        return None
    return base_url, file_title


def extract_category_title(category_url: str) -> Optional[tuple[str, str]]:
    """
    Extract base site URL and category title from a MediaWiki category URL.

    This mirrors extract_file_title so the caller can route correctly.
    Returns (base_url, category_title) or None when parsing fails.
    """
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
    """
    Run a MediaWiki API request anonymously with polite headers.
    """
    api_url = f"{base_url}/w/api.php"
    session = requests.Session()
    session.headers.update(
        {
            "User-Agent": "Wiki-Jio/1.0 (contact: Aditya-wiki-0545)",
            "Accept": "application/json",
        }
    )
    try:
        resp = session.get(api_url, params=params, timeout=10)
        resp.raise_for_status()
        return resp.json()
    except Exception as error:  # noqa: BLE001
        print(f"[X] API request failed: {error}")
        return None


def _fetch_image_metadata(base_url: str, title: str) -> Optional[dict]:
    """
    Fetch image metadata anonymously from a file title.
    """
    api_params = {
        "action": "query",
        "titles": title,
        "prop": "imageinfo|categories",
        "iiprop": "timestamp|user|size|url|metadata|extmetadata|dimensions",
        "format": "json",
        "formatversion": "2",
        "redirects": "true",
        "converttitles": "true",
        "cllimit": "50",
    }

    data = api_request(base_url, api_params)
    if not data:
        return None

    if "error" in data:
        print(f"[X] MediaWiki error: {data['error'].get('info')}")
        return None

    pages = data.get("query", {}).get("pages", [])
    if not pages:
        print("[X] No page data returned.")
        return None

    page = pages[0]
    if page.get("missing"):
        print("[X] File not found.")
        return None

    imageinfo = page.get("imageinfo", [])
    if not imageinfo:
        print("[X] No imageinfo returned.")
        return None

    info = imageinfo[0]
    ext = info.get("extmetadata", {}) or {}

    # Helper to extract extmetadata values safely.
    def ext_value(key: str) -> Optional[str]:
        val = ext.get(key)
        if isinstance(val, dict):
            return val.get("value")
        return val

    # Derive required fields with sensible fallbacks.
    title = page.get("title", title)
    file_url = info.get("url")
    author = ext_value("Artist") or info.get("user")
    size_bytes = info.get("size")
    license_type = ext_value("LicenseShortName") or ext_value("License")
    description = ext_value("ImageDescription") or ext_value("ObjectName")
    creation_date = ext_value("DateTimeOriginal") or ext_value("DateTime")
    width = info.get("width")
    height = info.get("height")

    categories_raw = page.get("categories", []) or []
    categories = [c.get("title") for c in categories_raw if isinstance(c, dict)]

    return {
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
    }


def get_image_metadata(image_url: str) -> Optional[dict]:
    """
    Public wrapper: parse URL then fetch metadata.
    """
    extracted = extract_file_title(image_url)
    if not extracted:
        print("[X] Could not parse file title from URL.")
        return None

    base_url, file_title = extracted
    return _fetch_image_metadata(base_url, file_title)


def fetch_category_files(category_url: str) -> list[dict]:
    """
    Fetch file metadata for a category (direct children only).
    """
    extracted = extract_category_title(category_url)
    if not extracted:
        print("[X] Could not parse category title from URL.")
        return []

    base_url, category_title = extracted

    files: list[str] = []
    cont: dict | None = None
    while True:
        params = {
            "action": "query",
            "list": "categorymembers",
            "cmtype": "file",
            "cmlimit": "100",
            "cmtitle": category_title,
            "format": "json",
            "formatversion": "2",
        }
        if cont:
            params.update(cont)

        data = api_request(base_url, params)
        if not data:
            break

        members = data.get("query", {}).get("categorymembers", []) or []
        files.extend([m.get("title") for m in members if isinstance(m, dict)])
        cont = data.get("continue")
        if not cont:
            break

    items: list[dict] = []
    for title in files:
        if not title:
            continue
        metadata = _fetch_image_metadata(base_url, title)
        if metadata:
            items.append(metadata)
    return items


def fetch_category_files_recursive(
    category_url: str,
    max_items: int | None = None,
    target_size: tuple[int, int] | None = None,
) -> list[dict]:
    """
    Recursively fetch file metadata from a category and all subcategories.
    Emits progress lines for each category and file batch.
    """
    extracted = extract_category_title(category_url)
    if not extracted:
        print("[X] Could not parse category title from URL.")
        return []

    base_url, category_title = extracted

    seen_cats: set[str] = set()
    seen_files: set[str] = set()
    items: list[dict] = []

    def fetch_files(file_titles: list[str], depth: int) -> None:
        if not file_titles:
            return
        indent = "  " * depth
        print(f"{indent}  [*] Fetching {len(file_titles)} image(s)...")
        for title in file_titles:
            if title in seen_files:
                continue
            seen_files.add(title)
            metadata = _fetch_image_metadata(base_url, title)
            if metadata:
                if target_size:
                    tw, th = target_size
                    if metadata.get("width") != tw or metadata.get("height") != th:
                        continue
                items.append(metadata)
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
                "action": "query",
                "list": "categorymembers",
                "cmtitle": cat_title,
                "cmlimit": "100",
                "cmtype": "file|subcat",
                "format": "json",
                "formatversion": "2",
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
                # Fallback on namespace if type is absent.
                if mtype is None:
                    ns = m.get("ns")
                    if ns == 14:
                        mtype = "subcat"
                    elif ns == 6:
                        mtype = "file"
                if mtype == "subcat":
                    walk(mtitle, depth + 1)
                elif mtype == "file":
                    local_files.append(mtitle)

            cont = data.get("continue")
            if not cont:
                break

        if local_files:
            fetch_files(local_files, depth)
            if max_items is not None and len(items) >= max_items:
                return

    walk(category_title)
    return items


# =============================================================================
# PRESENTATION
# =============================================================================

def display_metadata(metadata: dict) -> None:
    """Pretty-print the metadata with simple formatting."""
    print("=" * 60)
    print("IMAGE METADATA (anonymous API call)")
    print("=" * 60)
    print(f"Title:       {metadata.get('title', 'Unknown')}")
    print(f"URL:         {metadata.get('url', 'Unknown')}")
    print(f"Author:      {metadata.get('author', 'Unknown')}")
    print(f"Size bytes:  {metadata.get('size_bytes', 'Unknown')}")
    print(f"License:     {metadata.get('license_type', 'Unknown')}")
    print(f"Description: {metadata.get('description', 'Unknown')}")
    print(f"Created:     {metadata.get('creation_date', 'Unknown')}")
    print(f"Resolution:  {metadata.get('resolution', 'Unknown')}")
    categories = metadata.get("categories") or []
    print(f"Categories:  {', '.join(categories) if categories else 'None'}")


# =============================================================================
# CLI ENTRY POINT
# =============================================================================

def main() -> None:
    """Handle inputs and print metadata."""
    # Priority: CLI arg -> default constant -> prompt.
    if len(sys.argv) > 1:
        image_url = sys.argv[1]
    elif DEFAULT_IMAGE_URL:
        image_url = DEFAULT_IMAGE_URL
    else:
        print("Enter the image/file URL (e.g., https://en.wikipedia.org/wiki/File:Example.jpg)")
        image_url = input("URL: ").strip()

    if not image_url:
        print("[X] No URL provided.")
        sys.exit(1)

    if not image_url.startswith(("http://", "https://")):
        print("[!] URL does not start with http/https; using as-is.")

    # Route based on URL type so the CLI stays predictable.
    if "Category:" in image_url:
        results = fetch_category_files_recursive(
            category_url=image_url,
            max_items=20,
            target_size=(4288, 2848),
        )
        if not results:
            print("[X] No images found in category (including subcategories).")
            sys.exit(1)
        for item in results:
            display_metadata(item)
    else:
        metadata = get_image_metadata(image_url)
        if metadata:
            display_metadata(metadata)
        else:
            print("[X] Failed to fetch image metadata.")
            sys.exit(1)


if __name__ == "__main__":
    main()

