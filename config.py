# =============================================================================
# CONFIGURATION FOR WIKIMEDIA QUALITY IMAGE HARVESTER
# =============================================================================

# -----------------------------------------------------------------------------
# MediaWiki Username (REQUIRED)
# Used only for User-Agent - must be your Wikimedia Commons username
# -----------------------------------------------------------------------------
MEDIAWIKI_USERNAME = "Sanskardubedev"

# -----------------------------------------------------------------------------
# Target Resolution(s)
# The script will AUTO-EXPAND these to include nearby resolutions for faster results!
# 
# Option 1: Single resolution (most common)
# TARGET_RESOLUTION = (2560, 1440)
#
# Option 2: Multiple resolutions (matches ANY of these + nearby resolutions)
# TARGET_RESOLUTION = [(2560, 1440), (1280, 720), (1280, 853)]
#
# Option 3: Any resolution (fastest - no filtering)
# TARGET_RESOLUTION = None
# -----------------------------------------------------------------------------

# RECOMMENDED: Include common photo resolutions for better results
TARGET_RESOLUTION = [
    (2560, 1440),  # QHD monitor
    (1920, 1080),  # Full HD
    (1280, 853),   # Common photo (3:2 ratio)
    (1280, 720),   # HD
]

# If you want exact matches only (slower but more precise)
# TARGET_RESOLUTION = (2560, 1440)

# -----------------------------------------------------------------------------
# Maximum Number of Images
# -----------------------------------------------------------------------------
MAX_IMAGES = 50

# -----------------------------------------------------------------------------
# Default Categories - FAST HIGH-QUALITY SOURCES
# These are direct featured/quality categories that contain actual files
# -----------------------------------------------------------------------------
DEFAULT_CATEGORIES = [
    # Main featured collections (most curated, highest quality)
    "https://commons.wikimedia.org/wiki/Category:Featured_pictures_on_Wikimedia_Commons",
    
    # Quality images (verified good quality)
    "https://commons.wikimedia.org/wiki/Category:Quality_images",
    
    # Pictures of the Day (one selected each day)
    "https://commons.wikimedia.org/wiki/Category:Pictures_of_the_day_(Wikimedia_Commons)",
    
    # Valued images (most valuable in their scope)
    "https://commons.wikimedia.org/wiki/Category:Valued_images",
]

# =============================================================================
# ALTERNATIVE CATEGORY CONFIGURATIONS
# =============================================================================

# For nature/landscape photos:
# DEFAULT_CATEGORIES = [
#     "https://commons.wikimedia.org/wiki/Category:Featured_pictures_of_landscapes",
#     "https://commons.wikimedia.org/wiki/Category:Quality_images_of_nature",
# ]

# For architecture:
# DEFAULT_CATEGORIES = [
#     "https://commons.wikimedia.org/wiki/Category:Featured_pictures_of_architecture",
#     "https://commons.wikimedia.org/wiki/Category:Quality_images_of_buildings",
# ]

# For animals/wildlife:
# DEFAULT_CATEGORIES = [
#     "https://commons.wikimedia.org/wiki/Category:Featured_pictures_of_animals",
#     "https://commons.wikimedia.org/wiki/Category:Quality_images_of_mammals",
# ]

# For space/astronomy:
# DEFAULT_CATEGORIES = [
#     "https://commons.wikimedia.org/wiki/Category:Featured_pictures_of_space",
#     "https://commons.wikimedia.org/wiki/Category:Quality_images_of_astronomy",
# ]

# =============================================================================
# ADVANCED SETTINGS (Optional)
# =============================================================================

# Resolution tolerance (default: 0.30 = ±30%)
# Lower = stricter matching (slower), Higher = more flexible (faster)
# TOLERANCE = 0.30

# Minimum resolution score to accept (0-100, default: 50)
# Higher = only better matches, Lower = accept more variety
# MIN_RESOLUTION_SCORE = 50

# =============================================================================
# TIPS FOR FASTER RESULTS
# =============================================================================
# 
# 1. Use TARGET_RESOLUTION = None for fastest harvesting (no filtering)
# 
# 2. Include multiple resolutions in TARGET_RESOLUTION list:
#    TARGET_RESOLUTION = [(2560, 1440), (1920, 1080), (1280, 853)]
#    Script will find ANY of these (+ nearby sizes)
# 
# 3. Use main category URLs (Featured/Quality) - these are fastest:
#    DEFAULT_CATEGORIES = [
#        "https://commons.wikimedia.org/wiki/Category:Featured_pictures_on_Wikimedia_Commons",
#    ]
# 
# 4. Increase MAX_IMAGES if you want more variety:
#    MAX_IMAGES = 100
# 
# 5. Use --tolerance flag for more flexibility:
#    python get_image_metadata.py --tolerance 0.5
#
# =============================================================================