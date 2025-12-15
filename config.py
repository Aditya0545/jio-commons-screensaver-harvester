# =============================================================================
# WIKIMEDIA QUALITY IMAGE HARVESTER - CONFIGURATION
# Single source of truth - all settings controlled here only
# =============================================================================

# -----------------------------------------------------------------------------
# MediaWiki Username (REQUIRED)
# -----------------------------------------------------------------------------
MEDIAWIKI_USERNAME = ""

# -----------------------------------------------------------------------------
# Target Resolutions for Set-Top Box Screensaver
# Script will ONLY use these resolutions (no auto-expansion)
# Optimized for common TV/monitor aspect ratios
# 
# Format: List of (width, height) tuples
# Keep to 6-8 resolutions for optimal performance
# -----------------------------------------------------------------------------
TARGET_RESOLUTION = [
    (3840, 2160),  # 4K UHD (16:9) - Premium displays
    (2560, 1440),  # QHD (16:9) - High-end monitors
    (1920, 1080),  # Full HD (16:9) - Standard HD
    (1280, 720),   # HD Ready (16:9) - Basic HD
    (1920, 1280),  # DSLR format (3:2) - Photography
    (1280, 853),   # Common photo (3:2) - Cameras
    (2560, 1600),  # WQXGA (16:10) - Widescreen
    (1920, 1200),  # WUXGA (16:10) - Professional
]

# Set to None to accept ANY resolution (fastest but less specific)
# TARGET_RESOLUTION = None

# -----------------------------------------------------------------------------
# Common Resolutions Fallback
# If target resolutions are too strict, script can use these as backup
# Only used if ALLOW_COMMON_RESOLUTIONS = True
# -----------------------------------------------------------------------------
COMMON_RESOLUTIONS = [
    (3840, 2160),
    (2560, 1440),
    (1920, 1080),
    (1280, 720),
    (1920, 1280),
    (2048, 1365),
    (1280, 853),
]

# Enable fallback to common resolutions if target not found
ALLOW_COMMON_RESOLUTIONS = True

# -----------------------------------------------------------------------------
# Maximum Images Per Run
# -----------------------------------------------------------------------------
MAX_IMAGES = 50

# -----------------------------------------------------------------------------
# Resolution Matching Settings
# -----------------------------------------------------------------------------
TOLERANCE = 0.30  # ±30% tolerance (0.30 = 30%, 0.50 = 50%)
MIN_RESOLUTION_SCORE = 50  # Minimum score 0-100 (lower = more flexible)

# -----------------------------------------------------------------------------
# Categories to Harvest From - PRIORITY ORDER
# 
# Categories are processed in order listed below
# PRIORITY 1: Winning/Award categories (processed FIRST)
# PRIORITY 2: Featured/Quality categories (processed if needed)
# PRIORITY 3: General quality categories (backup)
#
# Keywords for PRIORITY categories: "winning", "winner", "award", 
# "picture of the year", "photo of the year"
# -----------------------------------------------------------------------------

# PRIORITY 1: Award-winning and competition winners (HIGHEST PRIORITY)
PRIORITY_CATEGORIES = [
    "https://commons.wikimedia.org/wiki/Category:Pictures_of_the_Year",
    "https://commons.wikimedia.org/wiki/Category:Wiki_Loves_Earth_winning_images",
    "https://commons.wikimedia.org/wiki/Category:Wiki_Loves_Monuments_winning_images",
    "https://commons.wikimedia.org/wiki/Category:Picture_of_the_Day_winners",
]

# PRIORITY 2: Featured and quality images (HIGH PRIORITY)
FEATURED_CATEGORIES = [
    "https://commons.wikimedia.org/wiki/Category:Featured_pictures_on_Wikimedia_Commons",
    "https://commons.wikimedia.org/wiki/Category:Quality_images",
    "https://commons.wikimedia.org/wiki/Category:Valued_images",
]

# PRIORITY 3: General quality categories (BACKUP)
GENERAL_CATEGORIES = [
    "https://commons.wikimedia.org/wiki/Category:Pictures_of_the_day_(Wikimedia_Commons)",
]

# Combine all categories in priority order
DEFAULT_CATEGORIES = PRIORITY_CATEGORIES + FEATURED_CATEGORIES + GENERAL_CATEGORIES

# -----------------------------------------------------------------------------
# Priority Keywords for Automatic Category Detection
# If script finds subcategories with these keywords, it prioritizes them
# -----------------------------------------------------------------------------
WINNING_KEYWORDS = [
    "winning", "winner", "winners", "award", "awards",
    "picture of the year", "photo of the year",
    "pictures of the year", "competition",
    "prize", "best of"
]

# =============================================================================
# ALTERNATIVE CONFIGURATIONS FOR SPECIFIC THEMES
# =============================================================================

# For Nature/Landscape Screensavers:
# TARGET_RESOLUTION = [(3840, 2160), (2560, 1440), (1920, 1080)]
# DEFAULT_CATEGORIES = [
#     "https://commons.wikimedia.org/wiki/Category:Wiki_Loves_Earth_winning_images",
#     "https://commons.wikimedia.org/wiki/Category:Featured_pictures_of_landscapes",
#     "https://commons.wikimedia.org/wiki/Category:Quality_images_of_nature",
# ]

# For Architecture/Monuments:
# DEFAULT_CATEGORIES = [
#     "https://commons.wikimedia.org/wiki/Category:Wiki_Loves_Monuments_winning_images",
#     "https://commons.wikimedia.org/wiki/Category:Featured_pictures_of_architecture",
# ]

# For Wildlife/Animals:
# DEFAULT_CATEGORIES = [
#     "https://commons.wikimedia.org/wiki/Category:Featured_pictures_of_animals",
#     "https://commons.wikimedia.org/wiki/Category:Quality_images_of_mammals",
# ]

# =============================================================================
# ADVANCED SETTINGS (Optional)
# =============================================================================

# Output file name (can be overridden via command line)
DEFAULT_OUTPUT_FILE = "results.xlsx"

# API request timeout in seconds
API_TIMEOUT = 20

# Number of parallel thumbnail downloads
THUMBNAIL_WORKERS = 10

# Thumbnail target width (pixels)
THUMBNAIL_WIDTH = 1920

# Minimum image dimensions to consider
MIN_WIDTH = 800
MIN_HEIGHT = 600

# Maximum image dimensions to consider
MAX_WIDTH = 10000
MAX_HEIGHT = 10000