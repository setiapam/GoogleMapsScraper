
---

# Google Maps Business Data Scraper & Database Updater

Python-based automation tool designed to scrape business data from Google Maps via a local API and perform a direct, intelligent update to an existing Excel database (`data.xlsx`). The script is built for stability, featuring **Smart Resume**, **Fuzzy Logic Matching**, and a **Dual-Logic** filtering system to ensure no duplicate work or data loss.

## 🚀 Key Features

* **Dynamic Keyword Construction**: Automatically generates search queries by combining: `nama_usaha`, `nmdesa`, `nmkec`, and `nmkab`.
* **Dual-Logic Resume & Skip**: Skips rows that already have an `input_id` or a `hasilgc` value.
* **Incremental Saving (Crash-Safe)**: Saves the Excel file after every successful row update to prevent data loss during interruptions.
* **Fuzzy Similarity Comparison**: Calculates a percentage score in the `perbandingan` column based on Name and Address similarity.
* **Automated "Not Found" Marking**: Marks failed or empty results as **"TIDAK DITEMUKAN"** to avoid re-processing the same failed keywords.

## 🛠 Prerequisites

### 1. Backend: Google Maps Scraper API

This script requires the [gosom/google-maps-scraper](https://github.com/gosom/google-maps-scraper) running in server mode.

You can run the backend in two ways:

* **Via Docker (Recommended):**
```bash
mkdir -p gmapsdata && docker run -v $PWD/gmapsdata:/gmapsdata -p 8080:8080 gosom/google-maps-scraper -data-folder /gmapsdata

```

* **Via Executable Release:**
If you prefer not to use Docker, you can download the **standalone executable version** available for **Windows, macOS, and Linux** from the [GitHub Releases page](https://github.com/gosom/google-maps-scraper/releases). Simply run the binary with the `server` argument.

> **Note:** Ensure the API is accessible at `http://localhost:8080` before running the Python script. Whether you use Docker or the executable version, the scraper server must be active to receive requests.

### 2. Python Environment

Install the required libraries:

```bash
pip install pandas requests openpyxl fuzzywuzzy python-Levenshtein

```

## 📂 File Requirements

The script expects a file named `data.xlsx` in the same directory with the following columns:

* `perusahaan_id`: Unique identifier.
* `nama_usaha`, `alamat_usaha`: Internal data for comparison.
* `nmdesa`, `nmkec`, `nmkab`: Used for search query construction.
* `input_id`, `hasilgc`: Status and results storage.

## 📖 Usage Instructions

1. **Prepare your data**: Place your `data.xlsx` file in the script folder.
2. **Start the Scraper Server**: Ensure your Docker container (gosom/google-maps-scraper) is running and the WebUI/API is accessible at port 8080.
3. **Run the script**:
```bash
python your_script_name.py

```


4. **Monitor Logs**: The terminal will show the progress, job IDs, and similarity scores for each business found.
5. **Stopping/Resuming**: You can stop the script at any time using `Ctrl+C`. When restarted, it will automatically resume from the last unfinished row.

## ⚙️ Configuration

| Variable | Description |
| --- | --- |
| `BASE_URL` | Set to `http://localhost:8080/api/v1`. |
| `TIMEOUT_WORKING` | Max seconds to wait for a job to complete (Default: 300s). |
| `MAPPING_KOLOM` | Mapping API CSV fields to your Excel column names. |

---
