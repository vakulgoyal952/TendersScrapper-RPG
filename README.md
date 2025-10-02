# 🏗️ eProc Scraper Tool

A fully automated web scraper for extracting government construction tenders from regional eProcurement portals across India. Built for a construction company to streamline the tender search process, analyze opportunities, and make informed bidding decisions.

---

## 🚀 Features

- ✅ Scrapes regional government eProcurement portals
- ✅ Saves tender data in clean, structured CSV format
- ✅ Includes filters for:
  - Organisation
  - Project location/area
  - Minimum tender amount
- ✅ Provides tender selection insights and basic cost analysis
- ✅ Fully automated via CLI and Jupyter Notebook
- ✅ UNIX executable for bulk tender fetching

---

## 🗂️ Repository Structure

├── Eproc_Crawler.py # Main Python script to scrape tender data
├── Eproc_crawler.ipynb # Jupyter Notebook for interactive filtering and analysis
├── eproc_tool/ # Contains compiled UNIX executable for automated scraping
├── build/ # Build files/scripts for UNIX tool compilation
├── logs/ # Runtime logs for monitoring and debugging
├── tenders_YYYY-MM-DD/ # Folder with date-based scrapped CSV files
└── README.md # You're here!


---

## ⚙️ Requirements

- Python 3.8+
- Recommended packages (see `requirements.txt` if included):
  - `requests`
  - `beautifulsoup4`
  - `pandas`
  - `tqdm`
  - `lxml` (optional, for faster parsing)

---

## 🧠 How It Works

### 1. **Data Scraping**
- `Eproc_Crawler.py` connects to government tender portals and scrapes project data.
- Results are saved in CSV files inside `tenders_YYYY-MM-DD/`.

### 2. **Filtering & Analysis**
- Use `Eproc_crawler.ipynb` to:
  - Filter tenders by organization, area, amount
  - Perform high-level analysis for selection

### 3. **Automated Execution**
- `eproc_tool` (UNIX executable) can be run via cron or shell for scheduled scraping.

---

## 🛠️ Usage

### 📌 Run Python Script
```bash
python3 Eproc_Crawler.py


---

### ✅ Next Steps:
- Add a `requirements.txt` file if you haven’t already.
- Include example CSV outputs or screenshots (optional).
- If the UNIX tool is compiled from C/C++ or Rust, consider adding build instructions in the `build/` folder or another section in the README.

Would you like help generating the `requirements.txt` file or adding usage examples to the notebook?
