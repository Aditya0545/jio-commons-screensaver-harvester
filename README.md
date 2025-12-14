# Wiki-Jio

A Python tool to fetch image metadata from MediaWiki sites (like Wikimedia Commons). This tool can extract image information, search through categories, filter by resolution, and save results to CSV or Excel files.

<img width="1921" height="649" alt="image" src="https://github.com/user-attachments/assets/3d5c69da-3505-4a88-ae10-a3cfd98aa062" />

---

## What This Tool Does

This tool helps you:
- Get detailed information about images from Wikimedia Commons or other MediaWiki sites
- Search through categories and their subcategories automatically
- Filter images by specific resolution (width and height)
- Save results to CSV or Excel files with clickable links
- Download images to your computer

---

## Tech Stack

This project uses:
- **Python 3** - The programming language
- **requests** - To make HTTP requests to MediaWiki API
- **tqdm** - To show progress bars while downloading
- **openpyxl** - To create Excel files with clickable hyperlinks
- **MediaWiki API** - The official API from Wikimedia to get image data

---

## How It Works

### Step-by-Step Process

1. **Input**: You provide a URL (either an image page or a category page)
2. **URL Parsing**: The tool extracts the image or category name from the URL
3. **API Requests**: It makes requests to the MediaWiki API to get image information
4. **Category Traversal**: If it's a category, it recursively searches through all subcategories
5. **Filtering**: Optionally filters images by resolution (width and height)
6. **Output**: Saves the results to CSV and Excel files with clickable links

### Key Features

- **Recursive Category Search**: Automatically finds images in subcategories
- **Resolution Filtering**: Can filter images by exact width and height
- **Progress Tracking**: Shows progress bars when downloading images
- **Clickable Links**: Excel files contain clickable hyperlinks to images
- **Error Handling**: Handles missing images and network errors gracefully

---

## Installation

### Prerequisites

- Python 3.7 or higher installed on your computer
- pip (Python package installer)

### Step 1: Clone or Download This Project

Download this project to your computer and navigate to the project folder.

### Step 2: Create a Virtual Environment (Recommended)

Open a terminal/command prompt in the project folder and run:

```bash
python -m venv venv
```

**On Windows:**
```bash
venv\Scripts\activate
```

**On Mac/Linux:**
```bash
source venv/bin/activate
```

### Step 3: Install Required Packages

```bash
pip install -r requirements.txt
```

This will install:
- requests (version 2.31.0)
- tqdm (version 4.65.0 or higher)
- openpyxl (version 3.1.0 or higher)

---

## Configuration

Before running the script, you can customize settings in `config.py`:

### Open `config.py` and Edit:

```python
# Your MediaWiki username (used for User-Agent header)
MEDIAWIKI_USERNAME = "YourUsername"

# Filter by resolution (set to None to fetch all images)
# Example: (1920, 1080) means width=1920, height=1080
TARGET_RESOLUTION = (6000, 4000)

# Maximum number of images to fetch
# Set to 0 for no limit
MAX_IMAGES = 10
```

**Important Notes:**
- `MEDIAWIKI_USERNAME`: This is just for identification. It doesn't need to be a real account.
- `TARGET_RESOLUTION`: Set to `None` if you want all images regardless of size
- `MAX_IMAGES`: Set to `0` if you want to fetch all matching images

---

## How to Run

### Basic Usage

**Fetch images from the default category:**
```bash
python get_image_metadata.py
```

**Fetch images from a specific category URL:**
```bash
python get_image_metadata.py "https://commons.wikimedia.org/wiki/Category:Your_Category_Name"
```

**Fetch a single image:**
```bash
python get_image_metadata.py "https://commons.wikimedia.org/wiki/File:Image_Name.jpg"
```

### Advanced Options

**Save to a custom CSV file:**
```bash
python get_image_metadata.py --csv my_results.csv
```

**Save to a custom Excel file:**
```bash
python get_image_metadata.py --xlsx my_results.xlsx
```

**Limit the number of images fetched:**
```bash
python get_image_metadata.py --max 50
```

**Filter by specific resolution:**
```bash
python get_image_metadata.py --width 1920 --height 1080
```

**Download images to a folder:**
```bash
python get_image_metadata.py --download images/
```

### Complete Example

```bash
python get_image_metadata.py "https://commons.wikimedia.org/wiki/Category:Winners_of_Wiki_Loves_Monuments_2024_by_country" --max 20 --xlsx winners.xlsx --download downloaded_images/
```

This command will:
- Fetch up to 20 images from the category
- Save results to `winners.xlsx`
- Download images to the `downloaded_images/` folder

---

## Command Line Arguments

| Argument | Short Form | Description | Example |
|----------|------------|-------------|---------|
| `url` | - | Image or Category URL (optional) | `"https://commons.wikimedia.org/wiki/Category:..."` |
| `--csv` | `-c` | Save results to CSV file | `--csv results.csv` |
| `--xlsx` | `-x` | Save results to Excel file | `--xlsx results.xlsx` |
| `--max` | `-m` | Maximum number of images to fetch | `--max 50` |
| `--width` | - | Filter by image width (requires --height) | `--width 1920` |
| `--height` | - | Filter by image height (requires --width) | `--height 1080` |
| `--download` | `-d` | Download images to directory | `--download images/` |

---

## Output Files

### CSV File
- Contains two columns: `title` and `url`
- Simple text format, can be opened in Excel or any text editor
- Example: `results.csv`

### Excel File (XLSX)
- Contains two columns: `title` and `url`
- URLs are clickable hyperlinks (blue and underlined)
- Automatically adjusts column widths
- Default filename: `results.xlsx` (if not specified)

### Downloaded Images
- Images are saved with their original filenames
- If a file already exists, a number is appended (e.g., `image_1.jpg`)
- Progress bars show download status for each image

---

## Understanding the Output

### What Information is Collected?

For each image, the tool collects:
- **Title**: The file name on Wikimedia
- **URL**: Direct link to the image file
- **Author**: Who uploaded/created the image
- **Size**: File size in bytes
- **Resolution**: Width and height in pixels
- **License**: Type of license (if available)
- **Description**: Image description (if available)
- **Categories**: All categories the image belongs to

### Example Output

When you run the script, you'll see:
```
[*] Scanning category: Category:Winners_of_Wiki_Loves_Monuments_2024_by_country
  [*] Fetching 5 image(s)...
============================================================
WINNING IMAGE (file name and URL)
============================================================
File Name:   File:Example_Image.jpg
URL:         https://upload.wikimedia.org/wikipedia/commons/...
[+] Saved 5 rows to XLSX: results.xlsx
```

---

## Troubleshooting

### Common Issues

**Problem: "Permission denied" when saving Excel file**
- **Solution**: Close the Excel file if it's open, then run the script again

**Problem: "No images found"**
- **Solution**: 
  - Check if the URL is correct
  - Verify the category exists
  - Try removing resolution filters (set `TARGET_RESOLUTION = None` in config.py)

**Problem: "API request failed"**
- **Solution**: 
  - Check your internet connection
  - The MediaWiki site might be temporarily unavailable
  - Wait a few minutes and try again

**Problem: Images not downloading**
- **Solution**: 
  - Check if the download directory path is correct
  - Ensure you have write permissions in that directory
  - Check your internet connection

**Problem: "Module not found" errors**
- **Solution**: 
  - Make sure you activated the virtual environment
  - Run `pip install -r requirements.txt` again

---

## How the Recursive Category Search Works

1. The tool starts with the main category you provide
2. It lists all files and subcategories in that category
3. For each subcategory, it repeats the process (recursion)
4. It continues until it finds enough images or runs out of categories
5. Images are filtered by resolution if specified
6. Results are collected and saved

**Example:**
```
Category: Winners
  ├── Category: India (subcategory)
  │   ├── Image1.jpg
  │   └── Image2.jpg
  ├── Category: USA (subcategory)
  │   └── Image3.jpg
  └── Image4.jpg
```

The tool will find: Image1, Image2, Image3, and Image4.

---

## Tips for Best Results

1. **Start Small**: Test with `--max 5` first to see if it works
2. **Use Specific Categories**: More specific categories return better results
3. **Check Resolution**: Some categories may not have images in your target resolution
4. **Be Patient**: Large categories can take several minutes to process
5. **Save Regularly**: Results are automatically saved, but you can specify custom filenames

---

## Example Use Cases

### Use Case 1: Find High-Resolution Images
```bash
python get_image_metadata.py --width 6000 --height 4000 --max 10
```

### Use Case 2: Download All Images from a Category
```bash
python get_image_metadata.py "https://commons.wikimedia.org/wiki/Category:Nature" --max 0 --download nature_images/
```

### Use Case 3: Create a Spreadsheet of Images
```bash
python get_image_metadata.py --csv images.csv --xlsx images.xlsx
```

---

## Technical Details

### API Endpoints Used
- `/w/api.php?action=query` - Main API endpoint for MediaWiki
- Uses `categorymembers` to list category contents
- Uses `imageinfo` to get image metadata

### Rate Limiting
- The tool includes delays between requests to be respectful to the API
- No authentication required (works anonymously)

### File Formats Supported
- All image formats supported by MediaWiki (JPG, PNG, GIF, etc.)

---

## License

This tool is provided as-is for educational and personal use.

---

## Support

If you encounter issues:
1. Check the Troubleshooting section above
2. Verify your Python version: `python --version` (should be 3.7+)
3. Ensure all packages are installed: `pip list`
4. Check the MediaWiki site is accessible in your browser

---

## Contributing

Feel free to improve this tool by:
- Adding more features
- Fixing bugs
- Improving documentation
- Making the code more efficient

---

**Happy Image Fetching! 🎉**
