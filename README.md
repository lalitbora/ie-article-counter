# Indian Express Group — Multi-Property Story Count & Article Analyzer

A complete, high-performance software solution to extract, calculate, and analyze published stories directly from official XML sitemaps across **ALL Indian Express Group publications**:

1. **The Indian Express** (`https://indianexpress.com/`)
2. **Financial Express** (`https://www.financialexpress.com/`)
3. **Loksatta** (`https://www.loksatta.com/`)
4. **Jansatta** (`https://www.jansatta.com/`)
5. **IE Tamil** (`https://tamil.indianexpress.com/`)
6. **IE Bengali** (`https://bengali.indianexpress.com/`)
7. **IE Gujarati** (`https://gujarati.indianexpress.com/`)
8. **IE Malayalam** (`https://malayalam.indianexpress.com/`)

---

## 🌟 Key Features

1. **Multi-Property Dropdown Selection:**
   - Easily switch between any of the 8 Indian Express publications from a single dropdown menu.
2. **Daily Story Count (Date-wise):**
   - Number of stories published each day, day-of-week breakdown, and daily trend visualization.
3. **Weekly Story Count (Week-wise):**
   - Grouped into ISO calendar weeks (e.g. `Week 48 (25 Nov - 01 Dec 2024)`), active days in range, total stories, daily average, and top publishing desk.
4. **Monthly Story Count (Month-wise):**
   - Active days, total stories published per month, and daily average.
5. **Section-wise Story Count (Desk-wise):**
   - Counts per desk/section and percentage share.
   - Includes Daily Section Matrix (pivot table: Dates on Rows, Sections on Columns).
6. **Full Article Metadata Extraction:**
   - Every single article's headline/title, clickable URL, section, sub-section, first publish time, and last modified ISO timestamp.
7. **Custom Output Selectors (Checkboxes):**
   - Check or uncheck any module (Daily, Weekly, Monthly, Sections, Metadata) in both the Web Dashboard and Terminal CLI.
8. **Multi-Sheet Styled Excel Export (.xlsx):**
   - Generates executive-formatted Excel workbooks honoring the selected property and modules.

---

## 🚀 How to Run the Software

### Method 1: Launch Software Menu (.bat - Recommended)
- Double-click **`launch_software.bat`** in the `IE Script` folder.
- An interactive black screen menu appears with options:
  - `[1]` Run in Terminal (Black Screen Interactive Mode with Property & Module pickers)
  - `[2]` Launch Web Dashboard in Browser (`http://localhost:5000`)
  - `[3]` Generate Public Live URL to Share with Team (100% Cloudflare Immune domestic tunnel)
  - `[4]` Quick Analysis: Last 7 Days
  - `[5]` Quick Analysis: Current Month to Date
  - `[6]` Push Latest Updates to GitHub (Sync Streamlit Cloud)
  - `[7]` Exit

### Method 2: Local Web Dashboard (Browser)
- Double-click **`start_web_app.bat`** or run `python web_app.py`.
- Open `http://localhost:5000/` in any browser.
- Select your desired publication from the dropdown menu, choose dates, toggle checkboxes, and click **Calculate & Analyze**.
- Download formatted Excel (.xlsx) with one click.

### Method 3: Cloud / Online Link (Streamlit Cloud)
- Access from any device/mobile 24/7 without needing your computer turned on:
  👉 **`https://ie-article-counter-2iqi73m6nnxwbczdb7cgjy.streamlit.app`**

### Method 4: Terminal Command Line (CLI)
```bash
# Interactive mode (prompts for publication, dates, and modules)
python count_articles.py

# CLI with specific publication and date range
python count_articles.py --property financialexpress --start 2024-12-01 --end 2024-12-05
python count_articles.py --property loksatta --start 2024-12-01 --end 2024-12-05
python count_articles.py --property ietamil --start 2024-12-01 --end 2024-12-05

# Custom output file
python count_articles.py --property indianexpress --start 01-12-2024 --end 15-12-2024 --output FE_Report.xlsx
```
