# =============================================================================
# WIKIMEDIA QUALITY IMAGE HARVESTER - CONFIGURATION
# Focused on: Wiki Loves Monuments 2025, Wiki Loves Folklore, Wiki Loves Birds
# =============================================================================

# -----------------------------------------------------------------------------
# MediaWiki Username (REQUIRED)
# Used only for User-Agent compliance
# -----------------------------------------------------------------------------
MEDIAWIKI_USERNAME = ""

# -----------------------------------------------------------------------------
# STRICT Target Resolutions (Set-Top Box)
# Only these resolutions will be accepted (±20 pixels tolerance)
# -----------------------------------------------------------------------------
TARGET_RESOLUTION = [
    (1280, 720),   # HD Ready (16:9)
    (2560, 1440),  # QHD (16:9)
]

# -----------------------------------------------------------------------------
# Resolution Matching Behavior - PIXEL TOLERANCE MODE
# Images within ±20 pixels of target dimensions will be accepted
# -----------------------------------------------------------------------------
PIXEL_TOLERANCE = 20  # ±20 pixels tolerance
EXACT_DIMENSIONS_ONLY = False  # Use pixel tolerance mode
MIN_RESOLUTION_SCORE = 95  # Minimum score to accept image

# Legacy percentage tolerance (not used when PIXEL_TOLERANCE > 0)
TOLERANCE = 0.015
ALLOW_COMMON_RESOLUTIONS = False
COMMON_RESOLUTIONS = [
    (1280, 720),
    (2560, 1440),
]

# -----------------------------------------------------------------------------
# Image Limits
# -----------------------------------------------------------------------------
MAX_IMAGES = 1000

# -----------------------------------------------------------------------------
# Image Dimension Safety Limits
# -----------------------------------------------------------------------------
MIN_WIDTH = 600
MIN_HEIGHT = 500
MAX_WIDTH = 10000
MAX_HEIGHT = 10000

# -----------------------------------------------------------------------------
# Output & Performance
# -----------------------------------------------------------------------------
DEFAULT_OUTPUT_FILE = "results.xlsx"
VIEWING_OUTPUT_FILE = "result_viewing.xlsx"  # Safe copy for viewing
API_TIMEOUT = 30
THUMBNAIL_WORKERS = 10
THUMBNAIL_WIDTH = 1920
BATCH_SIZE = 10  # Save every 10 images

# =============================================================================
# CATEGORY STRATEGY - FOCUSED ON 3 CAMPAIGNS ONLY
# -----------------------------------------------------------------------------
# 1️⃣ Wiki Loves Monuments 2025
# 2️⃣ Wiki Loves Folklore (all years - old ongoing event)
# 3️⃣ Wiki Loves Birds (all editions)
# =============================================================================

DEFAULT_CATEGORIES = [
    # =========================================================================
    # WIKI LOVES MONUMENTS 2025
    # =========================================================================
    "https://commons.wikimedia.org/wiki/Category:Images_from_Wiki_Loves_Monuments_2025",
    "https://commons.wikimedia.org/wiki/Category:Wiki_Loves_Monuments_2025",
    "https://commons.wikimedia.org/wiki/Category:Images_from_Wiki_Loves_Monuments_2025_in_India",
    "https://commons.wikimedia.org/wiki/Category:Images_from_Wiki_Loves_Monuments_2025_in_Germany",
    "https://commons.wikimedia.org/wiki/Category:Images_from_Wiki_Loves_Monuments_2025_in_France",
    "https://commons.wikimedia.org/wiki/Category:Images_from_Wiki_Loves_Monuments_2025_in_Italy",
    "https://commons.wikimedia.org/wiki/Category:Images_from_Wiki_Loves_Monuments_2025_in_Spain",
    "https://commons.wikimedia.org/wiki/Category:Images_from_Wiki_Loves_Monuments_2025_in_Poland",
    
    # =========================================================================
    # WIKI LOVES FOLKLORE (ALL YEARS - Ongoing event since many years)
    # =========================================================================
    "https://commons.wikimedia.org/wiki/Category:Images_from_Wiki_Loves_Folklore",
    "https://commons.wikimedia.org/wiki/Category:Wiki_Loves_Folklore",
    "https://commons.wikimedia.org/wiki/Category:Images_from_Wiki_Loves_Folklore_2025",
    "https://commons.wikimedia.org/wiki/Category:Images_from_Wiki_Loves_Folklore_2024",
    "https://commons.wikimedia.org/wiki/Category:Images_from_Wiki_Loves_Folklore_2023",
    "https://commons.wikimedia.org/wiki/Category:Images_from_Wiki_Loves_Folklore_2022",
    "https://commons.wikimedia.org/wiki/Category:Images_from_Wiki_Loves_Folklore_2021",
    "https://commons.wikimedia.org/wiki/Category:Images_from_Wiki_Loves_Folklore_2020",
    "https://commons.wikimedia.org/wiki/Category:Images_from_Wiki_Loves_Folklore_2019",
    "https://commons.wikimedia.org/wiki/Category:Wiki_Loves_Folklore_winning_images",
    
    # =========================================================================
    # WIKI LOVES BIRDS (ALL EDITIONS)
    # =========================================================================
    "https://commons.wikimedia.org/wiki/Category:Wiki_Loves_Birds",
    "https://commons.wikimedia.org/wiki/Category:Wiki_Loves_Birds_winning_images",
    "https://commons.wikimedia.org/wiki/Category:Images_from_Wiki_Loves_Birds",
    "https://commons.wikimedia.org/wiki/Category:Wiki_Loves_Birds_2025",
    "https://commons.wikimedia.org/wiki/Category:Wiki_Loves_Birds_2024",
    "https://commons.wikimedia.org/wiki/Category:Wiki_Loves_Birds_2023",
    "https://commons.wikimedia.org/wiki/Category:Wiki_Loves_Birds_2022",
    "https://commons.wikimedia.org/wiki/Category:Wiki_Loves_Birds_2021",
    "https://commons.wikimedia.org/wiki/Category:Wiki_Loves_Birds_India",
    "https://commons.wikimedia.org/wiki/Category:Wiki_Loves_Birds_India_2024",
    "https://commons.wikimedia.org/wiki/Category:Wiki_Loves_Birds_India_2023",
]

# -----------------------------------------------------------------------------
# Keywords to auto-detect priority subcategories during traversal
# -----------------------------------------------------------------------------
WINNING_KEYWORDS = [
    "winning", "winner", "award", "awarded",
    "best", "picture of the year",
    "competition", "prize", "featured"
]

# =============================================================================
# NOTES
# -----------------------------------------------------------------------------
# • ONLY 3 campaigns: Wiki Loves Monuments 2025, Wiki Loves Folklore, Wiki Loves Birds
# • ±20 pixel tolerance for resolution matching
# • Dual output: results.xlsx (main) + result_viewing.xlsx (safe to open)
# • Batch save every 10 images
# =============================================================================
