# -------------------------------------------------
# MediaWiki account username (public, non-secret)
# Used ONLY to construct a compliant User-Agent.
# Example: "SanskarDubey"
# -------------------------------------------------
MEDIAWIKI_USERNAME = ""

# -------------------------------------------------
# Resolution filter
# - Leave as None to fetch ALL images
# - Or set as a tuple: (width, height)
#   Example: (1920, 1080)
# -------------------------------------------------
TARGET_RESOLUTION = None

# -------------------------------------------------
# Number of images to fetch
# - 0 means NO LIMIT
# - Any positive integer limits results
# -------------------------------------------------
MAX_IMAGES = 100