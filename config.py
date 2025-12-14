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
TARGET_RESOLUTION = (6000, 4000)

# -------------------------------------------------
# Number of images to fetch
# - 0 means NO LIMIT
# - Any positive integer limits results
# -------------------------------------------------
MAX_IMAGES = 10

# Default categories to run when no URL is provided.
# Use full category URLs. Add/remove categories here for easy updates.
DEFAULT_CATEGORIES = [
    "https://commons.wikimedia.org/wiki/Category:Wiki_Loves_Earth",
    "https://commons.wikimedia.org/wiki/Category:Wiki_Loves_Birds"
]

# Now run the script
# Now it will traverse all sub categories of parent categories that you can see 

# if it will find given resolution image in india then it will show result if not then it will search other country
# it will search more because 10 images has not fetched yet with the given resolution so it will search all other category
# it is still searching more, so here we have got result 
# now check results.xlsx
# Thank You