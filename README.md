# The Indian Express — Story Count & Article Analyzer Software

A complete software solution to extract, calculate, and analyze published stories directly from **The Indian Express (`indianexpress.com`)** sitemaps.

---

## 🌟 Key Features

1. **Daily Story Count (Date-wise):**
   - Number of stories published each day, day-of-week breakdown, and daily trend visualization.
2. **Weekly Story Count (Week-wise):**
   - Grouped into ISO calendar weeks (e.g. `Week 45 (03 Nov - 09 Nov 2026)`), active days in range, total stories, daily average, and top publishing desk.
3. **Monthly Story Count (Month-wise):**
   - Active days, total stories published per month, and daily average.
4. **Section-wise Story Count (Desk-wise):**
   - Counts per desk (Cities, India, Entertainment, Sports, Lifestyle, Opinion, Explained, Business, etc.) and % share.
   - Includes Daily Section Matrix (pivot table: Dates on Rows, Sections on Columns).
5. **Full Article Metadata Extraction:**
   - Every single article's headline/title, clickable URL, section, sub-section, first publish time, and last modified ISO timestamp.
6. **Custom Output Selectors (Checkboxes):**
   - Check or uncheck any module (Daily, Weekly, Monthly, Sections, Metadata) both in the Web UI and Terminal CLI.
7. **Multi-Sheet Styled Excel Export (.xlsx):**
   - Generates executive-formatted Excel workbooks honoring only the selected modules.

---

## 🚀 How to Run the Software

### Method 1: Launch Software Menu (.bat - Recommended)
- Double-click **`launch_software.bat`** on your desktop in the `IE Script` folder.
- An interactive black screen menu appears with options:
  - `[1]` Run in Terminal (Black Screen Interactive Mode)
  - `[2]` Launch Web Dashboard in Browser (`http://localhost:5000`)
  - `[3]` Quick Analysis: Last 7 Days
  - `[4]` Quick Analysis: Current Month to Date
  - `[5]` Sync Updates to GitHub & Streamlit Cloud

### Method 2: Local Web Dashboard (Browser)
- Double-click **`start_web_app.bat`** or run `python web_app.py`.
- Open `http://localhost:5000/` in any browser.
- Check/uncheck options, pick dates, click **Calculate & Analyze**, and download Excel.

### Method 3: Cloud / Online Link (Streamlit Cloud)
- Access from any device/mobile 24/7 without needing your laptop turned on:
  👉 **`https://ie-article-counter-2iqi73m6nnxwbczdb7cgjy.streamlit.app`**

### Method 4: Terminal Command Line
```bash
# Interactive mode
python count_articles.py

# CLI with date range
python count_articles.py --start 2026-11-01 --end 2026-11-30

# Custom output file
python count_articles.py --start 01-11-2026 --end 30-11-2026 --output November_Report.xlsx
```
