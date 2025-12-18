# 🎨 Wikimedia Commons Quality Image Harvester

<div align="center">

[![Python](https://img.shields.io/badge/Python-3.7+-blue.svg)](https://www.python.org/downloads/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![MediaWiki API](https://img.shields.io/badge/API-MediaWiki-orange.svg)](https://www.mediawiki.org/wiki/API:Main_page)

**A high-performance Python tool for harvesting quality images from Wikimedia Commons**

*Optimized for Set-Top Box screensavers with intelligent resolution matching*

</div>

---

## ✨ Features

### 🎯 Smart Resolution Matching
- **Pixel Tolerance Mode** - Match images within ±N pixels of target resolutions
- **Multiple Target Resolutions** - Support for HD (1280×720), QHD (2560×1440), and more
- **Intelligent Scoring** - Images ranked by how closely they match target dimensions

### 📁 Focused Category Campaigns
Pre-configured to harvest from premium Wikimedia competition categories:
- 🏛️ **Wiki Loves Monuments 2025** - Architectural heritage photography
- 🎭 **Wiki Loves Folklore** - Cultural and traditional imagery (all years)
- 🐦 **Wiki Loves Birds** - Wildlife and ornithology photography

### 📊 Dual-File Output System
- **`results.xlsx`** - Primary output file (may be locked during writes)
- **`result_viewing.xlsx`** - Safe viewing copy (open anytime without stopping the script!)
- **Batch Saving** - Saves every 10 images to prevent data loss

### ⚡ Performance Optimizations
- **Connection Pooling** - Reuses HTTP connections with `requests.Session()`
- **Gzip Compression** - Reduced bandwidth with compressed API responses
- **Server-Friendly** - Uses `maxlag` parameter to respect Wikimedia server load
- **Retry Logic** - Automatic retries with exponential backoff on failures
- **Parallel Thumbnails** - Multi-threaded thumbnail generation

### 🛡️ Robust Error Handling
- Graceful fallbacks instead of crashes
- Continues processing on individual image failures
- Handles file permission errors with backup file creation
- Safe file operations that won't corrupt your data

---

## 🚀 Quick Start

### Prerequisites

- Python 3.7 or higher
- pip (Python package installer)

### Installation

```bash
# Clone the repository
git clone https://github.com/Aditya0545/jio-commons-screensaver-harvester.git
cd jio-commons-screensaver-harvester

# Create virtual environment (recommended)
python -m venv venv

# Activate virtual environment
# On Windows:
venv\Scripts\activate
# On Mac/Linux:
source venv/bin/activate

# Install dependencies
pip install -r requirements.txt
```

### Basic Usage

```bash
# Run with default settings (1000 images max)
python get_image_metadata.py

# Harvest specific number of images
python get_image_metadata.py --max 100

# Custom output file
python get_image_metadata.py -o my_images.xlsx --max 50
```

---

## ⚙️ Configuration

All settings are managed in `config.py`:

### Resolution Settings

```python
# Target resolutions for Set-Top Box displays
TARGET_RESOLUTION = [
    (1280, 720),   # HD Ready (16:9)
    (2560, 1440),  # QHD (16:9)
]

# ±20 pixel tolerance (accepts 1260-1300 × 700-740 for HD)
PIXEL_TOLERANCE = 20

# Minimum score to accept an image (95 = very close match)
MIN_RESOLUTION_SCORE = 95
```

### Output Settings

```python
# Maximum images to harvest
MAX_IMAGES = 1000

# Primary output file
DEFAULT_OUTPUT_FILE = "results.xlsx"

# Safe viewing copy (open while script runs)
VIEWING_OUTPUT_FILE = "result_viewing.xlsx"

# Images saved per batch
BATCH_SIZE = 10
```

### MediaWiki Authentication

```python
# Your MediaWiki username (for User-Agent header)
# Required for API compliance - doesn't need to be a real account
MEDIAWIKI_USERNAME = "YourUsername"
```

---

## 📋 Command Line Options

| Option | Short | Description | Default |
|--------|-------|-------------|---------|
| `--output` | `-o` | Output Excel file path | `results.xlsx` |
| `--max` | `-m` | Maximum images to harvest | `1000` (from config) |

### Examples

```bash
# Harvest 50 images to custom file
python get_image_metadata.py --max 50 -o my_collection.xlsx

# Use all default settings
python get_image_metadata.py

# Override max images from config
python get_image_metadata.py -m 200
```

---

## 📁 Output Format

### Excel Columns

| Column | Description |
|--------|-------------|
| `image_page_url` | Link to Wikimedia Commons page |
| `file_url` | Direct link to full-resolution image |
| `line1` | Attribution line (author + license) |
| `line2` | Short description (max 64 chars) |
| `best_thumb_url` | Optimized thumbnail URL (1920px width) |
| `best_res_under_1mb` | Resolution in "W × H pixels" format |
| `title` | Original file title on Commons |
| `description` | Full image description |
| `license` | License type (CC BY-SA, etc.) |
| `author` | Image creator/uploader |

### Sample Output

```
✅ Added 100 new images
📊 Total images in file: 100
📁 Results saved to: results.xlsx
👁️  Viewing copy at: result_viewing.xlsx
🏆 High-quality matches: 100

📐 New images by resolution:
   1280×720: 45 images
   1280×719: 23 images
   1279×720: 18 images
   2560×1440: 14 images
```

---

## 🔧 How It Works

### Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│                    Wikimedia Commons API                        │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│  Connection Pool (requests.Session)                             │
│  • Gzip compression • Retries • maxlag compliance               │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│  Category Queue (BFS with priority ordering)                    │
│  • Winning categories first • Subcategory traversal             │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│  Resolution Filter (±20px tolerance)                            │
│  • Batch size queries • Score-based ranking                     │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│  Metadata Enrichment                                            │
│  • Author • License • Description • Thumbnails                  │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│  Batch Writer (every 10 images)                                 │
│  results.xlsx ──copy──▶ result_viewing.xlsx                     │
└─────────────────────────────────────────────────────────────────┘
```

### Processing Flow

1. **Load Categories** - Configured campaigns from `config.py`
2. **Priority Queue** - Winning/award categories processed first
3. **Batch Size Query** - Fetch dimensions for 50 images at once
4. **Resolution Scoring** - Score each image against target resolutions
5. **Metadata Fetch** - Get full metadata for matching images
6. **Thumbnail Generation** - Parallel thumbnail URL generation
7. **Batch Save** - Write every 10 images to Excel + viewing copy
8. **Subcategory Traversal** - Queue subcategories for more images

---

## 🎯 Supported Categories

### Wiki Loves Monuments 2025
```
Category:Images_from_Wiki_Loves_Monuments_2025
Category:Images_from_Wiki_Loves_Monuments_2025_in_India
Category:Images_from_Wiki_Loves_Monuments_2025_in_Germany
Category:Images_from_Wiki_Loves_Monuments_2025_in_France
Category:Images_from_Wiki_Loves_Monuments_2025_in_Italy
Category:Images_from_Wiki_Loves_Monuments_2025_in_Spain
Category:Images_from_Wiki_Loves_Monuments_2025_in_Poland
```

### Wiki Loves Folklore (All Years)
```
Category:Images_from_Wiki_Loves_Folklore
Category:Images_from_Wiki_Loves_Folklore_2025
Category:Images_from_Wiki_Loves_Folklore_2024
Category:Images_from_Wiki_Loves_Folklore_2023
Category:Images_from_Wiki_Loves_Folklore_2022
Category:Images_from_Wiki_Loves_Folklore_2021
Category:Images_from_Wiki_Loves_Folklore_2020
Category:Images_from_Wiki_Loves_Folklore_2019
Category:Wiki_Loves_Folklore_winning_images
```

### Wiki Loves Birds (All Editions)
```
Category:Wiki_Loves_Birds
Category:Wiki_Loves_Birds_winning_images
Category:Wiki_Loves_Birds_2025
Category:Wiki_Loves_Birds_2024
Category:Wiki_Loves_Birds_2023
Category:Wiki_Loves_Birds_India
Category:Wiki_Loves_Birds_India_2024
```

---

## 🛠️ Troubleshooting

### Common Issues

| Problem | Solution |
|---------|----------|
| **"Permission denied"** | Close `results.xlsx` if open. Script auto-creates backup file. |
| **"No images found"** | Check internet connection. Verify categories have images at target resolution. |
| **"API timeout"** | Script retries automatically. If persistent, increase `API_TIMEOUT` in config. |
| **Slow performance** | Normal for large categories. Progress is shown in real-time. |
| **Empty viewing file** | Wait for first batch of 10 images to complete. |

### Safe Viewing

The `result_viewing.xlsx` file is specifically designed for safe viewing:

1. Open `result_viewing.xlsx` anytime
2. Script continues writing to `results.xlsx`
3. Viewing file updates after each batch
4. No file locking issues!

---

## 📊 Performance Tips

1. **Start Small** - Test with `--max 10` first
2. **Check Resolution** - Not all categories have images at target resolutions
3. **Be Patient** - Large categories take time to scan
4. **Monitor Progress** - Watch console output for real-time status
5. **Use Viewing File** - Check `result_viewing.xlsx` without stopping the script

---

## 🔌 API Compliance

This tool follows MediaWiki API best practices:

- ✅ **User-Agent Header** - Identifies requests properly
- ✅ **maxlag Parameter** - Respects server load
- ✅ **Gzip Compression** - Reduces bandwidth
- ✅ **Connection Reuse** - Efficient session handling
- ✅ **Rate Limiting** - No parallel bombardment
- ✅ **Retry Logic** - Graceful failure handling

---

## 📜 License

This project is provided as-is for educational and personal use.

---

## 🤝 Contributing

Contributions are welcome! Feel free to:

- 🐛 Report bugs
- 💡 Suggest features
- 📝 Improve documentation
- 🔧 Submit pull requests

---

## 🙏 Acknowledgments

- [Wikimedia Commons](https://commons.wikimedia.org/) - For the amazing image repository
- [MediaWiki API](https://www.mediawiki.org/wiki/API:Main_page) - For the powerful API
- All photographers contributing to Wiki Loves competitions

---

<div align="center">

**Made with ❤️ for quality screensaver content**

*Happy Harvesting! 🎉*

</div>
