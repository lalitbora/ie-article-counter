#!/usr/bin/env python3
"""
The Indian Express - Article Counter & Analyzer
================================================
Counts daily articles section-wise, daily total, and monthly total directly
from The Indian Express (indianexpress.com) official sitemaps.

Outputs:
  - Terminal: Formatted summary tables (Monthly, Daily, Section-wise).
  - Excel (.xlsx): Executive multi-sheet report with pivot matrix & raw data.
"""

import sys
import os
import re
import argparse
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.parse import urlparse
import xml.etree.ElementTree as ET
import urllib.request
import urllib.error

# Ensure Windows terminal doesn't crash on unicode characters
try:
    if sys.stdout.encoding != 'utf-8':
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Common friendly section names mapping
SECTION_TITLE_MAP = {
    'cities': 'Cities',
    'entertainment': 'Entertainment',
    'india': 'India',
    'sports': 'Sports',
    'lifestyle': 'Lifestyle',
    'opinion': 'Opinion',
    'world': 'World',
    'business': 'Business',
    'explained': 'Explained',
    'technology': 'Technology',
    'political-pulse': 'Political Pulse',
    'education': 'Education',
    'legal-news': 'Legal News',
    'trending': 'Trending',
    'upsc-current-affairs': 'UPSC Current Affairs',
    'health-wellness': 'Health & Wellness',
    'photos': 'Photo Galleries',
    'books-and-literature': 'Books & Literature',
    'web-stories': 'Web Stories',
    'delhi-confidential': 'Delhi Confidential',
    'north-east-india': 'North East India',
    'jobs': 'Jobs',
    'science': 'Science',
    'audio': 'Podcasts / Audio',
    'videos': 'Videos',
    'long-reads': 'Long Reads',
    'express-sunday-eye': 'Express Sunday Eye',
    'fresh-take': 'Fresh Take',
    'news-today': 'News Today'
}

# Slugs that belong to Cities
CITY_SLUGS = {
    'delhi', 'mumbai', 'chandigarh', 'pune', 'bangalore', 'bengaluru',
    'kolkata', 'lucknow', 'ahmedabad', 'hyderabad', 'chennai', 'jaipur',
    'patna', 'bhopal', 'shimla', 'dehradun', 'thiruvananthapuram', 'kochi',
    'nagpur', 'varanasi', 'kanpur', 'surat', 'ludhiana', 'amritsar',
    'gurgaon', 'noida', 'ghaziabad', 'faridabad'
}

# Slugs that belong to World
WORLD_SLUGS = {'pakistan', 'world'}

HEADERS = {
    'User-Agent': (
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
        'AppleWebKit/537.36 (KHTML, like Gecko) '
        'Chrome/124.0.0.0 Safari/537.36'
    ),
    'Accept': 'application/xml,text/xml,*/*',
    'Accept-Language': 'en-US,en;q=0.9',
}


def parse_flexible_date(date_str: str) -> datetime.date:
    """Parse flexible date strings (YYYY-MM-DD, DD-MM-YYYY, DD/MM/YYYY, etc.)."""
    date_str = date_str.strip()
    formats = [
        '%Y-%m-%d',
        '%d-%m-%Y',
        '%d/%m/%Y',
        '%Y/%m/%d',
        '%d.%m.%Y',
        '%Y.%m.%d',
        '%d %b %Y',
        '%d %B %Y',
        '%Y%m%d',
    ]
    for fmt in formats:
        try:
            return datetime.strptime(date_str, fmt).date()
        except ValueError:
            pass
    raise ValueError(
        f"Invalid date format: '{date_str}'. Please use YYYY-MM-DD or DD-MM-YYYY (e.g. 2026-11-01 or 01-11-2026)."
    )


def extract_title_from_slug(slug: str) -> str:
    """Generate a readable title from article URL slug."""
    if not slug:
        return ""
    clean_slug = re.sub(r'-\d+$', '', slug)
    words = clean_slug.replace('-', ' ').split()
    return ' '.join(word.capitalize() for word in words)


def parse_article_url(url: str):
    """
    Parse section, sub-section, and title from Indian Express article URL.
    Returns: (section_clean, sub_section_clean, title)
    """
    parsed = urlparse(url)
    parts = [p for p in parsed.path.strip('/').split('/') if p]

    if not parts:
        return 'Other', 'General', ''

    section_key = 'other'
    sub_section = '-'
    slug = ''

    if parts[0] == 'article':
        if len(parts) >= 4:
            section_key = parts[1].lower()
            sub_section = parts[2]
            slug = parts[3]
        elif len(parts) == 3:
            section_key = parts[1].lower()
            sub_section = 'General'
            slug = parts[2]
        elif len(parts) == 2:
            section_key = parts[1].lower()
            sub_section = 'General'
            slug = ''
        else:
            section_key = 'article'
            sub_section = 'General'
            slug = ''
    else:
        section_key = parts[0].lower()
        if len(parts) > 2:
            sub_section = parts[1]
            slug = parts[-1]
        elif len(parts) == 2:
            sub_section = 'General'
            slug = parts[1]
        else:
            sub_section = 'General'
            slug = ''

    # Normalize cities
    if section_key in CITY_SLUGS:
        sub_section = section_key.capitalize()
        section_key = 'cities'
    elif section_key in WORLD_SLUGS and section_key != 'world':
        sub_section = section_key.capitalize()
        section_key = 'world'

    section_name = SECTION_TITLE_MAP.get(
        section_key,
        section_key.replace('-', ' ').title()
    )
    sub_section_name = sub_section.replace('-', ' ').title() if sub_section != '-' else '-'
    title = extract_title_from_slug(slug)

    return section_name, sub_section_name, title


def fetch_sitemap_for_date(target_date: datetime.date, max_retries: int = 3):
    """Fetch and parse daily Indian Express sitemap for a given date."""
    yyyy = target_date.year
    mm = f"{target_date.month:02d}"
    dd = f"{target_date.day:02d}"
    sitemap_url = f"https://indianexpress.com/sitemap.xml?yyyy={yyyy}&mm={mm}&dd={dd}"

    articles = []
    error_msg = None

    for attempt in range(1, max_retries + 1):
        try:
            req = urllib.request.Request(sitemap_url, headers=HEADERS)
            with urllib.request.urlopen(req, timeout=20) as response:
                content = response.read()

            root = ET.fromstring(content)
            urls = root.findall('.//{http://www.sitemaps.org/schemas/sitemap/0.9}url')

            for u in urls:
                loc_elem = u.find('{http://www.sitemaps.org/schemas/sitemap/0.9}loc')
                lastmod_elem = u.find('{http://www.sitemaps.org/schemas/sitemap/0.9}lastmod')

                if loc_elem is not None and loc_elem.text:
                    url_str = loc_elem.text.strip()
                    lastmod_str = lastmod_elem.text.strip() if lastmod_elem is not None and lastmod_elem.text else ''
                    
                    pub_time = ''
                    if 'T' in lastmod_str:
                        time_part = lastmod_str.split('T')[1]
                        pub_time = time_part[:8]

                    sec, sub_sec, title = parse_article_url(url_str)

                    articles.append({
                        'date': target_date.strftime('%Y-%m-%d'),
                        'month': target_date.strftime('%Y-%m'),
                        'month_name': target_date.strftime('%B %Y'),
                        'day_of_week': target_date.strftime('%A'),
                        'time': pub_time,
                        'section': sec,
                        'sub_section': sub_sec,
                        'title': title,
                        'url': url_str,
                        'lastmod': lastmod_str
                    })

            error_msg = None
            break

        except urllib.error.HTTPError as e:
            if e.code == 404:
                error_msg = f"No sitemap found (404) for {target_date}"
                break
            error_msg = f"HTTP Error {e.code}: {e.reason}"
        except Exception as e:
            error_msg = str(e)

        if attempt < max_retries:
            import time
            time.sleep(1.0 * attempt)

    return target_date, articles, error_msg


def fetch_all_dates(start_date: datetime.date, end_date: datetime.date, max_workers: int = 6):
    """Fetch articles for all dates in the range concurrently."""
    total_days = (end_date - start_date).days + 1
    dates = [start_date + timedelta(days=i) for i in range(total_days)]

    print(f"\n" + "=" * 65)
    print(f"  Fetching article data from The Indian Express sitemaps...")
    print(f"  Period  : {start_date.strftime('%d-%b-%Y')} to {end_date.strftime('%d-%b-%Y')} ({total_days} days)")
    print(f"  Threads : {max_workers} concurrent requests")
    print("=" * 65)

    all_articles = []
    date_status = {}
    completed_count = 0

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_to_date = {executor.submit(fetch_sitemap_for_date, d): d for d in dates}

        for future in as_completed(future_to_date):
            d = future_to_date[future]
            completed_count += 1
            try:
                target_date, articles, error_msg = future.result()
                if error_msg:
                    date_status[target_date] = f"Error: {error_msg}"
                    print(f"  [{completed_count:3d}/{total_days:3d}] {target_date} -> 0 articles ({error_msg})")
                else:
                    date_status[target_date] = f"{len(articles)} articles"
                    all_articles.extend(articles)
                    print(f"  [{completed_count:3d}/{total_days:3d}] {target_date} -> {len(articles):3d} articles")
            except Exception as exc:
                date_status[d] = f"Exception: {exc}"
                print(f"  [{completed_count:3d}/{total_days:3d}] {d} -> Failed: {exc}")

    return all_articles, date_status


def print_ascii_table(title: str, headers: list, rows: list, alignments: list = None):
    """Print a clean formatted terminal table."""
    if not rows:
        print(f"\n{title}: No data available.")
        return

    col_widths = [len(h) for h in headers]
    for row in rows:
        for idx, val in enumerate(row):
            col_widths[idx] = max(col_widths[idx], len(str(val)))

    if alignments is None:
        alignments = ['left'] + ['right'] * (len(headers) - 1)

    sep_line = "+" + "+".join("-" * (w + 2) for w in col_widths) + "+"
    double_sep = "=" * len(sep_line)

    print("\n" + double_sep)
    print(f"  {title.upper()}")
    print(double_sep)
    print(sep_line)

    header_str = "|"
    for h, w in zip(headers, col_widths):
        header_str += f" {h.center(w)} |"
    print(header_str)
    print(sep_line)

    for r in rows:
        row_str = "|"
        for val, w, align in zip(r, col_widths, alignments):
            val_str = str(val)
            if align == 'right':
                row_str += f" {val_str.rjust(w)} |"
            elif align == 'center':
                row_str += f" {val_str.center(w)} |"
            else:
                row_str += f" {val_str.ljust(w)} |"
        print(row_str)

    print(sep_line)


def display_terminal_summary(df: pd.DataFrame, start_date: datetime.date, end_date: datetime.date):
    """Display Monthly, Daily, and Section-wise counts in Terminal."""
    if df.empty:
        print("\nNo articles found for the specified date range.")
        return

    total_articles = len(df)
    total_days = (end_date - start_date).days + 1
    daily_avg = total_articles / total_days if total_days > 0 else 0

    print("\n" + "=" * 65)
    print("           INDIAN EXPRESS ARTICLE COUNT REPORT")
    print("=" * 65)
    print(f"  Period        : {start_date.strftime('%d-%b-%Y')} to {end_date.strftime('%d-%b-%Y')}")
    print(f"  Total Days    : {total_days}")
    print(f"  Total Articles: {total_articles:,}")
    print(f"  Daily Average : {daily_avg:.1f} articles/day")
    print("=" * 65)

    # 1. MONTHLY COUNT
    monthly_rows = []
    month_groups = df.groupby(['month', 'month_name'])
    for (m_code, m_name), m_df in month_groups:
        m_count = len(m_df)
        m_days = m_df['date'].nunique()
        m_avg = m_count / m_days if m_days > 0 else 0
        monthly_rows.append([m_name, m_days, f"{m_count:,}", f"{m_avg:.1f}"])
    
    monthly_rows.append([
        "TOTAL",
        df['date'].nunique(),
        f"{total_articles:,}",
        f"{daily_avg:.1f}"
    ])
    print_ascii_table(
        "1. Monthly Article Count Summary",
        ["Month", "Days Active", "Total Articles", "Daily Average"],
        monthly_rows,
        alignments=['left', 'right', 'right', 'right']
    )

    # 2. DAILY COUNT
    daily_rows = []
    daily_groups = df.groupby(['date', 'day_of_week']).size().reset_index(name='count').sort_values('date')
    for _, row in daily_groups.iterrows():
        daily_rows.append([
            row['date'],
            row['day_of_week'],
            f"{row['count']:,}"
        ])
    daily_rows.append(["GRAND TOTAL", f"{total_days} Days", f"{total_articles:,}"])
    print_ascii_table(
        "2. Daily Total Article Count",
        ["Date", "Day of Week", "Total Articles"],
        daily_rows,
        alignments=['left', 'left', 'right']
    )

    # 3. SECTION-WISE COUNT
    section_rows = []
    sec_counts = df['section'].value_counts()
    for sec, count in sec_counts.items():
        pct = (count / total_articles) * 100
        section_rows.append([sec, f"{count:,}", f"{pct:.1f}%"])
    section_rows.append(["TOTAL", f"{total_articles:,}", "100.0%"])
    print_ascii_table(
        "3. Section-Wise Article Count (Overall)",
        ["Section", "Article Count", "Share (%)"],
        section_rows,
        alignments=['left', 'right', 'right']
    )


def export_to_excel(df: pd.DataFrame, start_date: datetime.date, end_date: datetime.date, filepath: str):
    """Generate an executive, professional multi-tab Excel workbook."""
    if df.empty:
        print("No data to export.")
        return

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # Styling Palette
    navy_header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    total_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    total_font = Font(name="Calibri", size=11, bold=True, color="000000")
    regular_font = Font(name="Calibri", size=11)
    title_font = Font(name="Calibri", size=14, bold=True, color="1F4E78")
    subtitle_font = Font(name="Calibri", size=10, italic=True, color="595959")
    
    thin_border_side = Side(border_style="thin", color="D9D9D9")
    thin_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    
    total_top_side = Side(border_style="thin", color="000000")
    total_bottom_side = Side(border_style="double", color="000000")
    total_border = Border(top=total_top_side, bottom=total_bottom_side, left=thin_border_side, right=thin_border_side)

    total_articles = len(df)
    total_days = (end_date - start_date).days + 1
    daily_avg = total_articles / total_days if total_days > 0 else 0

    # -------------------------------------------------------------
    # SHEET 1: Monthly Summary
    # -------------------------------------------------------------
    ws_month = wb.create_sheet(title="Monthly Summary")
    ws_month.views.sheetView[0].showGridLines = True

    ws_month["A1"] = "The Indian Express - Article Count Report"
    ws_month["A1"].font = title_font
    ws_month["A2"] = f"Period: {start_date.strftime('%d-%b-%Y')} to {end_date.strftime('%d-%b-%Y')} | Generated: {datetime.now().strftime('%d-%b-%Y %I:%M %p')}"
    ws_month["A2"].font = subtitle_font

    headers_month = ["Month", "Days Active", "Total Articles", "Daily Average"]
    for col_idx, h in enumerate(headers_month, 1):
        cell = ws_month.cell(row=4, column=col_idx, value=h)
        cell.fill = navy_header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center" if col_idx > 1 else "left", vertical="center")
    ws_month.row_dimensions[4].height = 24

    row_idx = 5
    month_groups = df.groupby(['month', 'month_name'])
    for (m_code, m_name), m_df in month_groups:
        m_count = len(m_df)
        m_days = m_df['date'].nunique()
        m_avg = m_count / m_days if m_days > 0 else 0

        ws_month.cell(row=row_idx, column=1, value=m_name).alignment = Alignment(horizontal="left")
        ws_month.cell(row=row_idx, column=2, value=m_days).alignment = Alignment(horizontal="right")
        ws_month.cell(row=row_idx, column=3, value=m_count).alignment = Alignment(horizontal="right")
        ws_month.cell(row=row_idx, column=4, value=round(m_avg, 1)).alignment = Alignment(horizontal="right")

        for col_idx in range(1, 5):
            c = ws_month.cell(row=row_idx, column=col_idx)
            c.font = regular_font
            c.border = thin_border
            if col_idx in [2, 3]:
                c.number_format = "#,##0"
            elif col_idx == 4:
                c.number_format = "#,##0.0"
        row_idx += 1

    # Total row
    ws_month.cell(row=row_idx, column=1, value="Grand Total").alignment = Alignment(horizontal="left")
    ws_month.cell(row=row_idx, column=2, value=df['date'].nunique()).alignment = Alignment(horizontal="right")
    ws_month.cell(row=row_idx, column=3, value=total_articles).alignment = Alignment(horizontal="right")
    ws_month.cell(row=row_idx, column=4, value=round(daily_avg, 1)).alignment = Alignment(horizontal="right")
    for col_idx in range(1, 5):
        c = ws_month.cell(row=row_idx, column=col_idx)
        c.fill = total_fill
        c.font = total_font
        c.border = total_border
        if col_idx in [2, 3]:
            c.number_format = "#,##0"
        elif col_idx == 4:
            c.number_format = "#,##0.0"

    # -------------------------------------------------------------
    # SHEET 2: Daily Count
    # -------------------------------------------------------------
    ws_daily = wb.create_sheet(title="Daily Total Count")
    ws_daily.views.sheetView[0].showGridLines = True

    ws_daily["A1"] = "Daily Total Article Count"
    ws_daily["A1"].font = title_font
    ws_daily["A2"] = f"Period: {start_date} to {end_date}"
    ws_daily["A2"].font = subtitle_font

    headers_daily = ["Date", "Day of Week", "Total Articles"]
    for col_idx, h in enumerate(headers_daily, 1):
        cell = ws_daily.cell(row=4, column=col_idx, value=h)
        cell.fill = navy_header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center" if col_idx > 1 else "left", vertical="center")
    ws_daily.row_dimensions[4].height = 24

    daily_groups = df.groupby(['date', 'day_of_week']).size().reset_index(name='count').sort_values('date')
    row_idx = 5
    for _, r in daily_groups.iterrows():
        ws_daily.cell(row=row_idx, column=1, value=r['date']).alignment = Alignment(horizontal="center")
        ws_daily.cell(row=row_idx, column=2, value=r['day_of_week']).alignment = Alignment(horizontal="left")
        ws_daily.cell(row=row_idx, column=3, value=r['count']).alignment = Alignment(horizontal="right")
        for col_idx in range(1, 4):
            c = ws_daily.cell(row=row_idx, column=col_idx)
            c.font = regular_font
            c.border = thin_border
            if col_idx == 3:
                c.number_format = "#,##0"
        row_idx += 1

    # Daily Total
    ws_daily.cell(row=row_idx, column=1, value="Grand Total").alignment = Alignment(horizontal="left")
    ws_daily.cell(row=row_idx, column=2, value=f"{total_days} Days").alignment = Alignment(horizontal="left")
    ws_daily.cell(row=row_idx, column=3, value=total_articles).alignment = Alignment(horizontal="right")
    for col_idx in range(1, 4):
        c = ws_daily.cell(row=row_idx, column=col_idx)
        c.fill = total_fill
        c.font = total_font
        c.border = total_border
        if col_idx == 3:
            c.number_format = "#,##0"

    # -------------------------------------------------------------
    # SHEET 3: Section Summary
    # -------------------------------------------------------------
    ws_sec = wb.create_sheet(title="Section Summary")
    ws_sec.views.sheetView[0].showGridLines = True

    ws_sec["A1"] = "Overall Section-Wise Article Count"
    ws_sec["A1"].font = title_font
    ws_sec["A2"] = f"Total Articles: {total_articles:,} across {total_days} day(s)"
    ws_sec["A2"].font = subtitle_font

    headers_sec = ["Rank", "Section Name", "Article Count", "Share (%)"]
    for col_idx, h in enumerate(headers_sec, 1):
        cell = ws_sec.cell(row=4, column=col_idx, value=h)
        cell.fill = navy_header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center" if col_idx in [1, 4] else "left" if col_idx == 2 else "right", vertical="center")
    ws_sec.row_dimensions[4].height = 24

    sec_counts = df['section'].value_counts()
    row_idx = 5
    for rank, (sec, count) in enumerate(sec_counts.items(), 1):
        pct = count / total_articles if total_articles > 0 else 0
        ws_sec.cell(row=row_idx, column=1, value=rank).alignment = Alignment(horizontal="center")
        ws_sec.cell(row=row_idx, column=2, value=sec).alignment = Alignment(horizontal="left")
        ws_sec.cell(row=row_idx, column=3, value=count).alignment = Alignment(horizontal="right")
        ws_sec.cell(row=row_idx, column=4, value=pct).alignment = Alignment(horizontal="right")

        for col_idx in range(1, 5):
            c = ws_sec.cell(row=row_idx, column=col_idx)
            c.font = regular_font
            c.border = thin_border
            if col_idx == 3:
                c.number_format = "#,##0"
            elif col_idx == 4:
                c.number_format = "0.0%"
        row_idx += 1

    # Section Total
    ws_sec.cell(row=row_idx, column=1, value="").alignment = Alignment(horizontal="center")
    ws_sec.cell(row=row_idx, column=2, value="Grand Total").alignment = Alignment(horizontal="left")
    ws_sec.cell(row=row_idx, column=3, value=total_articles).alignment = Alignment(horizontal="right")
    ws_sec.cell(row=row_idx, column=4, value=1.0).alignment = Alignment(horizontal="right")
    for col_idx in range(1, 5):
        c = ws_sec.cell(row=row_idx, column=col_idx)
        c.fill = total_fill
        c.font = total_font
        c.border = total_border
        if col_idx == 3:
            c.number_format = "#,##0"
        elif col_idx == 4:
            c.number_format = "0.0%"

    # -------------------------------------------------------------
    # SHEET 4: Daily Section Matrix (Pivot Table)
    # -------------------------------------------------------------
    ws_matrix = wb.create_sheet(title="Daily Section Matrix")
    ws_matrix.views.sheetView[0].showGridLines = True

    ws_matrix["A1"] = "Daily Section-Wise Matrix"
    ws_matrix["A1"].font = title_font
    ws_matrix["A2"] = "Breakdown of articles published in each section by date"
    ws_matrix["A2"].font = subtitle_font

    pivot = pd.crosstab(
        index=df['date'],
        columns=df['section'],
        margins=True,
        margins_name="Total"
    )

    matrix_headers = ["Date"] + list(pivot.columns)
    for col_idx, h in enumerate(matrix_headers, 1):
        cell = ws_matrix.cell(row=4, column=col_idx, value=h)
        cell.fill = navy_header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center" if col_idx == 1 else "right", vertical="center")
    ws_matrix.row_dimensions[4].height = 24

    row_idx = 5
    for dt_val, row_data in pivot.iterrows():
        is_total_row = (dt_val == "Total")
        date_label = "Grand Total" if is_total_row else str(dt_val)

        ws_matrix.cell(row=row_idx, column=1, value=date_label).alignment = Alignment(horizontal="center" if not is_total_row else "left")

        for col_idx, count_val in enumerate(row_data, 2):
            cell = ws_matrix.cell(row=row_idx, column=col_idx, value=int(count_val))
            cell.alignment = Alignment(horizontal="right")
            cell.number_format = "#,##0"

        for col_idx in range(1, len(matrix_headers) + 1):
            c = ws_matrix.cell(row=row_idx, column=col_idx)
            if is_total_row:
                c.fill = total_fill
                c.font = total_font
                c.border = total_border
            else:
                c.font = regular_font
                c.border = thin_border
                if col_idx == len(matrix_headers):
                    c.font = Font(name="Calibri", size=11, bold=True)
        row_idx += 1

    # -------------------------------------------------------------
    # SHEET 5: Raw Article Data
    # -------------------------------------------------------------
    ws_data = wb.create_sheet(title="All Articles Data")
    ws_data.views.sheetView[0].showGridLines = True

    headers_data = ["Date", "Day", "Time (IST)", "Section", "Sub-Section", "Article Title", "URL"]
    for col_idx, h in enumerate(headers_data, 1):
        cell = ws_data.cell(row=1, column=col_idx, value=h)
        cell.fill = navy_header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center" if col_idx <= 3 else "left", vertical="center")
    ws_data.row_dimensions[1].height = 22

    for r_idx, art in enumerate(df.to_dict('records'), 2):
        ws_data.cell(row=r_idx, column=1, value=art['date']).alignment = Alignment(horizontal="center")
        ws_data.cell(row=r_idx, column=2, value=art['day_of_week']).alignment = Alignment(horizontal="left")
        ws_data.cell(row=r_idx, column=3, value=art['time']).alignment = Alignment(horizontal="center")
        ws_data.cell(row=r_idx, column=4, value=art['section']).alignment = Alignment(horizontal="left")
        ws_data.cell(row=r_idx, column=5, value=art['sub_section']).alignment = Alignment(horizontal="left")
        ws_data.cell(row=r_idx, column=6, value=art['title']).alignment = Alignment(horizontal="left")
        
        url_cell = ws_data.cell(row=r_idx, column=7, value=art['url'])
        url_cell.alignment = Alignment(horizontal="left")
        url_cell.font = Font(name="Calibri", size=10, color="0563C1", underline="single")

        for c_idx in range(1, 7):
            ws_data.cell(row=r_idx, column=c_idx).font = regular_font

    # Auto-adjust column widths
    for ws in wb.worksheets:
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.row in [1, 2] and ws.title != "All Articles Data":
                    continue
                val = str(cell.value or '')
                max_len = max(max_len, len(val))
            ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 65)

    wb.save(filepath)
    if isinstance(filepath, str):
        print(f"\nExcel Report successfully saved to:\n  -> {filepath}")


def run_pipeline(start_date_str: str, end_date_str: str, output_path: str = None, max_workers: int = 6):
    """Core execution flow from date strings to reports."""
    start_date = parse_flexible_date(start_date_str)
    end_date = parse_flexible_date(end_date_str)

    if start_date > end_date:
        print(f"\n[Notice] Start Date ({start_date}) was after End Date ({end_date}). Swapping automatically.")
        start_date, end_date = end_date, start_date

    all_articles, date_status = fetch_all_dates(start_date, end_date, max_workers=max_workers)

    if not all_articles:
        print("\nNo articles could be fetched for this date range.")
        return

    df = pd.DataFrame(all_articles)

    # 1. Terminal display
    display_terminal_summary(df, start_date, end_date)

    # 2. Excel export
    if not output_path:
        filename = f"IE_Article_Count_{start_date.strftime('%Y%m%d')}_to_{end_date.strftime('%Y%m%d')}.xlsx"
        output_path = os.path.join(os.getcwd(), filename)

    export_to_excel(df, start_date, end_date, output_path)


def main():
    parser = argparse.ArgumentParser(
        description="The Indian Express - Daily Article Counter & Analyzer"
    )
    parser.add_argument(
        "--start", "-s",
        type=str,
        help="Start Date (e.g. 2026-11-01 or 01-11-2026)"
    )
    parser.add_argument(
        "--end", "-e",
        type=str,
        help="End Date (e.g. 2026-11-15 or 15-11-2026)"
    )
    parser.add_argument(
        "--output", "-o",
        type=str,
        help="Custom output Excel filename / path (.xlsx)"
    )
    parser.add_argument(
        "--workers", "-w",
        type=int,
        default=6,
        help="Number of concurrent worker threads (default: 6)"
    )

    args = parser.parse_args()

    # If arguments not provided via CLI, prompt interactively
    if not args.start or not args.end:
        print("=" * 65)
        print("     THE INDIAN EXPRESS - ARTICLE COUNTER & ANALYZER     ")
        print("=" * 65)
        print("Format: YYYY-MM-DD (e.g. 2026-11-01) ya DD-MM-YYYY (e.g. 01-11-2026)\n")

        while True:
            start_input = input("Enter Start Date: ").strip()
            try:
                parse_flexible_date(start_input)
                break
            except ValueError as e:
                print(f"  {e}")

        while True:
            end_input = input("Enter End Date:   ").strip()
            try:
                parse_flexible_date(end_input)
                break
            except ValueError as e:
                print(f"  {e}")

        run_pipeline(start_input, end_input, output_path=args.output, max_workers=args.workers)
    else:
        run_pipeline(args.start, args.end, output_path=args.output, max_workers=args.workers)


if __name__ == "__main__":
    main()
