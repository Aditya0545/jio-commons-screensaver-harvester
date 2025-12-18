# =============================================================================
# WIKIMEDIA QUALITY IMAGE HARVESTER - CONFIGURATION
# =============================================================================

# MediaWiki Username
MEDIAWIKI_USERNAME = "Adityakumar0545"

# Maximum Images Per Run
MAX_IMAGES = 100

# Target Resolutions
TARGET_RESOLUTION = [
    (2560, 1440),
    (1280, 720),
]

# Output File
DEFAULT_OUTPUT_FILE = "results.xlsx"

# Winning/Award Category Keywords (for priority processing)
WINNING_KEYWORDS = ["winner", "winners", "award", "awards", "prize"]

# Categories to Harvest From
# Supports both Category: and Commons: namespace pages
DEFAULT_CATEGORIES = [
    # Priority: Winning/Award categories (processed first)
    "https://commons.wikimedia.org/wiki/Commons:Wiki_Loves_Earth_2025/Winners",
    
    # Featured Pictures - High quality images selected by community
    "https://commons.wikimedia.org/wiki/Category:Featured_pictures_on_Wikimedia_Commons",
    
    # Quality Images - Another high-quality category
    "https://commons.wikimedia.org/wiki/Category:Quality_images",
    
    # Picture of the Year - Annual competition winners
    "https://commons.wikimedia.org/wiki/Category:Picture_of_the_Year",
    
    # Recent Featured Pictures (2024-2025)
    "https://commons.wikimedia.org/wiki/Category:Featured_pictures_2024",
    "https://commons.wikimedia.org/wiki/Category:Featured_pictures_2025",
    
    # Wiki Loves Earth categories
    "https://commons.wikimedia.org/wiki/Category:Wiki_Loves_Earth_2025_in_Indonesia",
    "https://commons.wikimedia.org/wiki/Category:Wiki_Loves_Earth_2024",
    "https://commons.wikimedia.org/wiki/Category:Wiki_Loves_Earth_2023",
]   
