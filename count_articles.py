#!/usr/bin/env python3
"""
The Indian Express - Article & Story Count Analyzer
====================================================
Comprehensive sitemap analyzer that fetches published articles and generates:
  - Daily Story Count (Date-wise)
  - Weekly Story Count (Week-wise with start/end dates)
  - Monthly Story Count (Month-wise with daily averages)
  - Section-wise Story Count (Overall, Daily, Weekly, and Monthly matrices)
  - Full Article Metadata (URLs, Headlines, Publish Date/Time, Last Modified Time)
  - Multi-sheet Styled Excel Workbook (.xlsx)

Author: Antigravity Assistant
Environment: Python 3.8+
"""

import sys
import os
import io
import re
import argparse
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.parse import urlparse
import xml.etree.ElementTree as ET

# Ensure standard output uses UTF-8 to prevent charmap encoding errors on Windows
try:
    if sys.stdout.encoding != 'utf-8':
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Use curl_cffi with Chrome TLS impersonation to bypass Cloudflare on cloud servers (AWS/Streamlit Cloud)
try:
    from curl_cffi import requests as cffi_requests
    _USE_CURL_CFFI = True
    _SESSION = None
except ImportError:
    _USE_CURL_CFFI = False
    try:
        import cloudscraper
        _SESSION = cloudscraper.create_scraper(
            browser={
                'browser': 'chrome',
                'platform': 'windows',
                'desktop': True
            }
        )
    except ImportError:
        import requests
        _SESSION = requests.Session()

# Realistic Chrome headers for HTTP requests
HEADERS = {
    'User-Agent': (
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
        'AppleWebKit/537.36 (KHTML, like Gecko) '
        'Chrome/124.0.0.0 Safari/537.36'
    ),
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
    'Accept-Language': 'en-US,en;q=0.9',
    'Accept-Encoding': 'gzip, deflate',
    'Connection': 'keep-alive',
    'Sec-Ch-Ua': '"Chromium";v="124", "Google Chrome";v="124", "Not-A.Brand";v="99"',
    'Sec-Ch-Ua-Mobile': '?0',
    'Sec-Ch-Ua-Platform': '"Windows"',
    'Sec-Fetch-Dest': 'document',
    'Sec-Fetch-Mode': 'navigate',
    'Sec-Fetch-Site': 'none',
    'Sec-Fetch-User': '?1',
    'Upgrade-Insecure-Requests': '1',
}

# Mapping of section URL slugs to canonical display titles
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

# Slugs that map directly under the 'Cities' desk
CITY_SLUGS = {
    'delhi', 'mumbai', 'chandigarh', 'pune', 'bangalore', 'bengaluru',
    'kolkata', 'lucknow', 'ahmedabad', 'hyderabad', 'chennai', 'jaipur',
    'patna', 'bhopal', 'shimla', 'dehradun', 'thiruvananthapuram', 'kochi',
    'nagpur', 'varanasi', 'kanpur', 'surat', 'ludhiana', 'amritsar',
    'gurgaon', 'noida', 'ghaziabad', 'faridabad'
}

# Slugs that map directly under the 'World' desk
WORLD_SLUGS = {'pakistan', 'world'}


def parse_flexible_date(date_str: str) -> datetime.date:
    """
    Parse a user-entered date string in multiple common formats.
    Supported: YYYY-MM-DD, DD-MM-YYYY, DD/MM/YYYY, YYYY/MM/DD, DD.MM.YYYY, etc.
    """
    date_str = str(date_str).strip()
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
    """
    Generate a clean, readable headline from an article's URL slug.
    Removes trailing article IDs (e.g., -10963395) and converts dashes to title case.
    """
    if not slug:
        return ""
    clean_slug = re.sub(r'-\d+$', '', slug)
    words = clean_slug.replace('-', ' ').split()
    return ' '.join(word.capitalize() for word in words)


def parse_article_url(url: str):
    """
    Parse section, sub-section, and headline/title from an Indian Express URL.
    Returns: (section_clean, sub_section_clean, headline)
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

    # Normalize individual city paths to Cities desk
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
    headline = extract_title_from_slug(slug)

    return section_name, sub_section_name, headline


def compute_week_info(date_obj: datetime.date):
    """
    Calculate calendar week start (Monday), week end (Sunday), and a descriptive label.
    Example: 'Week 46 (10-Nov to 16-Nov 2026)'
    """
    iso_year, iso_week, _ = date_obj.isocalendar()
    monday = date_obj - timedelta(days=date_obj.weekday())
    sunday = monday + timedelta(days=6)
    week_key = f"{iso_year}-W{iso_week:02d}"
    week_label = f"Week {iso_week:02d} ({monday.strftime('%d-%b')} to {sunday.strftime('%d-%b-%Y')})"
    return week_key, week_label, monday, sunday


def fetch_sitemap_for_date(target_date: datetime.date, max_retries: int = 3):
    """
    Fetch and parse the daily XML sitemap for a specific date from Indian Express.
    Returns: (target_date, list_of_articles, error_msg)
    """
    yyyy = target_date.year
    mm = f"{target_date.month:02d}"
    dd = f"{target_date.day:02d}"
    sitemap_url = f"https://indianexpress.com/sitemap.xml?yyyy={yyyy}&mm={mm}&dd={dd}"

    articles = []
    error_msg = None

    week_key, week_label, monday, sunday = compute_week_info(target_date)

    for attempt in range(1, max_retries + 1):
        try:
            if _USE_CURL_CFFI:
                resp = cffi_requests.get(sitemap_url, headers=HEADERS, impersonate="chrome124", timeout=20)
            else:
                resp = _SESSION.get(sitemap_url, headers=HEADERS, timeout=20)
            
            if resp.status_code == 404:
                error_msg = f"No sitemap found (404) for {target_date}"
                break
            elif resp.status_code == 403:
                error_msg = f"HTTP 403 Forbidden (Cloudflare bot block)"
                break
            elif resp.status_code != 200:
                error_msg = f"HTTP {resp.status_code}"
                continue

            content = resp.content
            root = ET.fromstring(content)
            urls = root.findall('.//{http://www.sitemaps.org/schemas/sitemap/0.9}url')

            for u in urls:
                loc_elem = u.find('{http://www.sitemaps.org/schemas/sitemap/0.9}loc')
                lastmod_elem = u.find('{http://www.sitemaps.org/schemas/sitemap/0.9}lastmod')

                if loc_elem is not None and loc_elem.text:
                    url_str = loc_elem.text.strip()
                    lastmod_str = lastmod_elem.text.strip() if lastmod_elem is not None and lastmod_elem.text else ''
                    
                    # Extract time portions from ISO lastmod timestamp
                    pub_time = ''
                    if 'T' in lastmod_str:
                        time_part = lastmod_str.split('T')[1]
                        pub_time = time_part[:8]

                    sec, sub_sec, headline = parse_article_url(url_str)

                    articles.append({
                        'date': target_date.strftime('%Y-%m-%d'),
                        'publish_date': target_date.strftime('%Y-%m-%d'),
                        'first_publish_time': pub_time if pub_time else '00:00:00',
                        'last_modified_time': lastmod_str,
                        'day_of_week': target_date.strftime('%A'),
                        'week_key': week_key,
                        'week_label': week_label,
                        'month': target_date.strftime('%Y-%m'),
                        'month_name': target_date.strftime('%B %Y'),
                        'section': sec,
                        'sub_section': sub_sec,
                        'title': headline,
                        'headline': headline,
                        'url': url_str
                    })

            error_msg = None
            break

        except Exception as e:
            error_msg = str(e)

        if attempt < max_retries:
            import time
            time.sleep(1.0 * attempt)

    return target_date, articles, error_msg


def fetch_all_dates(start_date: datetime.date, end_date: datetime.date, max_workers: int = 6):
    """
    Fetch articles for all dates in the range concurrently using a ThreadPoolExecutor.
    """
    total_days = (end_date - start_date).days + 1
    dates = [start_date + timedelta(days=i) for i in range(total_days)]

    print(f"\n" + "=" * 68)
    print(f"  FETCHING STORIES FROM THE INDIAN EXPRESS SITEMAPS")
    print(f"  Period  : {start_date.strftime('%d-%b-%Y')} to {end_date.strftime('%d-%b-%Y')} ({total_days} days)")
    print(f"  Threads : {max_workers} concurrent network workers")
    print("=" * 68)

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
    """Print an aligned ASCII table to standard output."""
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


def display_terminal_summary(df: pd.DataFrame, start_date: datetime.date, end_date: datetime.date,
                             include_daily: bool = True, include_weekly: bool = True,
                             include_monthly: bool = True, include_sections: bool = True):
    """
    Display comprehensive terminal tables based on selected checkboxes.
    """
    if df.empty:
        print("\nNo articles found for the specified date range.")
        return

    total_articles = len(df)
    total_days = (end_date - start_date).days + 1
    daily_avg = total_articles / total_days if total_days > 0 else 0

    print("\n" + "=" * 68)
    print("        THE INDIAN EXPRESS — STORY COUNT ANALYSIS REPORT")
    print("=" * 68)
    print(f"  Period        : {start_date.strftime('%d-%b-%Y')} to {end_date.strftime('%d-%b-%Y')}")
    print(f"  Total Days    : {total_days}")
    print(f"  Total Stories : {total_articles:,}")
    print(f"  Daily Average : {daily_avg:.1f} stories/day")
    print("=" * 68)

    # 1. Monthly Summary Table
    if include_monthly:
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
            "1. Monthly Story Count Summary",
            ["Month", "Days Active", "Total Stories", "Daily Average"],
            monthly_rows,
            alignments=['left', 'right', 'right', 'right']
        )

    # 2. Weekly Summary Table
    if include_weekly:
        weekly_rows = []
        week_groups = df.groupby(['week_key', 'week_label'])
        for (w_key, w_label), w_df in week_groups:
            w_count = len(w_df)
            w_days = w_df['date'].nunique()
            w_avg = w_count / w_days if w_days > 0 else 0
            top_sec = w_df['section'].value_counts().index[0] if not w_df.empty else "-"
            weekly_rows.append([w_label, w_days, f"{w_count:,}", f"{w_avg:.1f}", top_sec])
        
        weekly_rows.append([
            "GRAND TOTAL",
            df['date'].nunique(),
            f"{total_articles:,}",
            f"{daily_avg:.1f}",
            "-"
        ])
        print_ascii_table(
            "2. Weekly Story Count Summary",
            ["Calendar Week", "Days in Range", "Total Stories", "Daily Average", "Top Section"],
            weekly_rows,
            alignments=['left', 'right', 'right', 'right', 'left']
        )

    # 3. Daily Summary Table
    if include_daily:
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
            "3. Daily Total Story Count",
            ["Date", "Day of Week", "Total Stories"],
            daily_rows,
            alignments=['left', 'left', 'right']
        )

    # 4. Section-wise Overall Table
    if include_sections:
        section_rows = []
        sec_counts = df['section'].value_counts()
        for sec, count in sec_counts.items():
            pct = (count / total_articles) * 100
            section_rows.append([sec, f"{count:,}", f"{pct:.1f}%"])
        section_rows.append(["TOTAL", f"{total_articles:,}", "100.0%"])
        print_ascii_table(
            "4. Section-Wise Story Count (Overall)",
            ["Section Name", "Story Count", "Share (%)"],
            section_rows,
            alignments=['left', 'right', 'right']
        )


def export_to_excel(df: pd.DataFrame, start_date: datetime.date, end_date: datetime.date, filepath,
                    include_daily: bool = True, include_weekly: bool = True,
                    include_monthly: bool = True, include_sections: bool = True,
                    include_details: bool = True):
    """
    Generate an executive, professional multi-sheet Excel workbook (.xlsx)
    including only the sheets selected by the user.
    """
    if df.empty:
        print("No data to export.")
        return

    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # Remove default blank sheet

    # Corporate Color Palette & Styling
    navy_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    total_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    total_font = Font(name="Calibri", size=11, bold=True, color="000000")
    regular_font = Font(name="Calibri", size=11)
    title_font = Font(name="Calibri", size=14, bold=True, color="1F4E78")
    subtitle_font = Font(name="Calibri", size=10, italic=True, color="595959")
    
    thin_side = Side(border_style="thin", color="D9D9D9")
    thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    
    total_top = Side(border_style="thin", color="000000")
    total_bottom = Side(border_style="double", color="000000")
    total_border = Border(top=total_top, bottom=total_bottom, left=thin_side, right=thin_side)

    total_articles = len(df)
    total_days = (end_date - start_date).days + 1
    daily_avg = total_articles / total_days if total_days > 0 else 0

    # -------------------------------------------------------------
    # SHEET 1: Monthly Summary
    # -------------------------------------------------------------
    if include_monthly:
        ws_m = wb.create_sheet(title="Monthly Summary")
        ws_m.views.sheetView[0].showGridLines = True

        ws_m["A1"] = "The Indian Express — Monthly Story Count Summary"
        ws_m["A1"].font = title_font
        ws_m["A2"] = f"Period: {start_date.strftime('%d-%b-%Y')} to {end_date.strftime('%d-%b-%Y')} | Generated: {datetime.now().strftime('%d-%b-%Y %I:%M %p')}"
        ws_m["A2"].font = subtitle_font

        headers_m = ["Month", "Days Active", "Total Stories", "Daily Average"]
        for col_idx, h in enumerate(headers_m, 1):
            c = ws_m.cell(row=4, column=col_idx, value=h)
            c.fill = navy_fill
            c.font = header_font
            c.alignment = Alignment(horizontal="center" if col_idx > 1 else "left", vertical="center")
        ws_m.row_dimensions[4].height = 24

        row_idx = 5
        month_groups = df.groupby(['month', 'month_name'])
        for (m_code, m_name), m_df in month_groups:
            m_count = len(m_df)
            m_days = m_df['date'].nunique()
            m_avg = m_count / m_days if m_days > 0 else 0

            ws_m.cell(row=row_idx, column=1, value=m_name).alignment = Alignment(horizontal="left")
            ws_m.cell(row=row_idx, column=2, value=m_days).alignment = Alignment(horizontal="right")
            ws_m.cell(row=row_idx, column=3, value=m_count).alignment = Alignment(horizontal="right")
            ws_m.cell(row=row_idx, column=4, value=round(m_avg, 1)).alignment = Alignment(horizontal="right")

            for col_idx in range(1, 5):
                cell = ws_m.cell(row=row_idx, column=col_idx)
                cell.font = regular_font
                cell.border = thin_border
                if col_idx in [2, 3]:
                    cell.number_format = "#,##0"
                elif col_idx == 4:
                    cell.number_format = "#,##0.0"
            row_idx += 1

        # Grand Total Row
        ws_m.cell(row=row_idx, column=1, value="Grand Total").alignment = Alignment(horizontal="left")
        ws_m.cell(row=row_idx, column=2, value=df['date'].nunique()).alignment = Alignment(horizontal="right")
        ws_m.cell(row=row_idx, column=3, value=total_articles).alignment = Alignment(horizontal="right")
        ws_m.cell(row=row_idx, column=4, value=round(daily_avg, 1)).alignment = Alignment(horizontal="right")
        for col_idx in range(1, 5):
            cell = ws_m.cell(row=row_idx, column=col_idx)
            cell.fill = total_fill
            cell.font = total_font
            cell.border = total_border
            if col_idx in [2, 3]:
                cell.number_format = "#,##0"
            elif col_idx == 4:
                cell.number_format = "#,##0.0"

    # -------------------------------------------------------------
    # SHEET 2: Weekly Summary
    # -------------------------------------------------------------
    if include_weekly:
        ws_w = wb.create_sheet(title="Weekly Summary")
        ws_w.views.sheetView[0].showGridLines = True

        ws_w["A1"] = "The Indian Express — Weekly Story Count Summary"
        ws_w["A1"].font = title_font
        ws_w["A2"] = f"Period: {start_date.strftime('%d-%b-%Y')} to {end_date.strftime('%d-%b-%Y')} (ISO Calendar Weeks)"
        ws_w["A2"].font = subtitle_font

        headers_w = ["Calendar Week", "Days in Range", "Total Stories", "Daily Average", "Top Publishing Section"]
        for col_idx, h in enumerate(headers_w, 1):
            c = ws_w.cell(row=4, column=col_idx, value=h)
            c.fill = navy_fill
            c.font = header_font
            c.alignment = Alignment(horizontal="center" if col_idx in [2, 3, 4] else "left", vertical="center")
        ws_w.row_dimensions[4].height = 24

        row_idx = 5
        week_groups = df.groupby(['week_key', 'week_label'])
        for (w_key, w_label), w_df in week_groups:
            w_count = len(w_df)
            w_days = w_df['date'].nunique()
            w_avg = w_count / w_days if w_days > 0 else 0
            top_sec = w_df['section'].value_counts().index[0] if not w_df.empty else "-"

            ws_w.cell(row=row_idx, column=1, value=w_label).alignment = Alignment(horizontal="left")
            ws_w.cell(row=row_idx, column=2, value=w_days).alignment = Alignment(horizontal="right")
            ws_w.cell(row=row_idx, column=3, value=w_count).alignment = Alignment(horizontal="right")
            ws_w.cell(row=row_idx, column=4, value=round(w_avg, 1)).alignment = Alignment(horizontal="right")
            ws_w.cell(row=row_idx, column=5, value=top_sec).alignment = Alignment(horizontal="left")

            for col_idx in range(1, 6):
                cell = ws_w.cell(row=row_idx, column=col_idx)
                cell.font = regular_font
                cell.border = thin_border
                if col_idx in [2, 3]:
                    cell.number_format = "#,##0"
                elif col_idx == 4:
                    cell.number_format = "#,##0.0"
            row_idx += 1

        # Grand Total Row
        ws_w.cell(row=row_idx, column=1, value="Grand Total").alignment = Alignment(horizontal="left")
        ws_w.cell(row=row_idx, column=2, value=df['date'].nunique()).alignment = Alignment(horizontal="right")
        ws_w.cell(row=row_idx, column=3, value=total_articles).alignment = Alignment(horizontal="right")
        ws_w.cell(row=row_idx, column=4, value=round(daily_avg, 1)).alignment = Alignment(horizontal="right")
        ws_w.cell(row=row_idx, column=5, value="All Sections").alignment = Alignment(horizontal="left")
        for col_idx in range(1, 6):
            cell = ws_w.cell(row=row_idx, column=col_idx)
            cell.fill = total_fill
            cell.font = total_font
            cell.border = total_border
            if col_idx in [2, 3]:
                cell.number_format = "#,##0"
            elif col_idx == 4:
                cell.number_format = "#,##0.0"

    # -------------------------------------------------------------
    # SHEET 3: Daily Total Count
    # -------------------------------------------------------------
    if include_daily:
        ws_d = wb.create_sheet(title="Daily Total Count")
        ws_d.views.sheetView[0].showGridLines = True

        ws_d["A1"] = "The Indian Express — Daily Total Story Count"
        ws_d["A1"].font = title_font
        ws_d["A2"] = f"Period: {start_date} to {end_date}"
        ws_d["A2"].font = subtitle_font

        headers_d = ["Date", "Day of Week", "Total Stories Published"]
        for col_idx, h in enumerate(headers_d, 1):
            c = ws_d.cell(row=4, column=col_idx, value=h)
            c.fill = navy_fill
            c.font = header_font
            c.alignment = Alignment(horizontal="center" if col_idx == 1 else "left" if col_idx == 2 else "right", vertical="center")
        ws_d.row_dimensions[4].height = 24

        daily_groups = df.groupby(['date', 'day_of_week']).size().reset_index(name='count').sort_values('date')
        row_idx = 5
        for _, r in daily_groups.iterrows():
            ws_d.cell(row=row_idx, column=1, value=r['date']).alignment = Alignment(horizontal="center")
            ws_d.cell(row=row_idx, column=2, value=r['day_of_week']).alignment = Alignment(horizontal="left")
            ws_d.cell(row=row_idx, column=3, value=r['count']).alignment = Alignment(horizontal="right")
            for col_idx in range(1, 4):
                cell = ws_d.cell(row=row_idx, column=col_idx)
                cell.font = regular_font
                cell.border = thin_border
                if col_idx == 3:
                    cell.number_format = "#,##0"
            row_idx += 1

        # Daily Grand Total
        ws_d.cell(row=row_idx, column=1, value="Grand Total").alignment = Alignment(horizontal="left")
        ws_d.cell(row=row_idx, column=2, value=f"{total_days} Days").alignment = Alignment(horizontal="left")
        ws_d.cell(row=row_idx, column=3, value=total_articles).alignment = Alignment(horizontal="right")
        for col_idx in range(1, 4):
            cell = ws_d.cell(row=row_idx, column=col_idx)
            cell.fill = total_fill
            cell.font = total_font
            cell.border = total_border
            if col_idx == 3:
                cell.number_format = "#,##0"

    # -------------------------------------------------------------
    # SHEET 4: Section Summary & Matrices
    # -------------------------------------------------------------
    if include_sections:
        ws_s = wb.create_sheet(title="Section Summary")
        ws_s.views.sheetView[0].showGridLines = True

        ws_s["A1"] = "The Indian Express — Section-Wise Story Count"
        ws_s["A1"].font = title_font
        ws_s["A2"] = f"Total Stories: {total_articles:,} across {total_days} day(s)"
        ws_s["A2"].font = subtitle_font

        headers_s = ["Rank", "Section / Desk Name", "Total Stories", "Share (%)"]
        for col_idx, h in enumerate(headers_s, 1):
            c = ws_s.cell(row=4, column=col_idx, value=h)
            c.fill = navy_fill
            c.font = header_font
            c.alignment = Alignment(horizontal="center" if col_idx in [1, 4] else "left" if col_idx == 2 else "right", vertical="center")
        ws_s.row_dimensions[4].height = 24

        sec_counts = df['section'].value_counts()
        row_idx = 5
        for rank, (sec, count) in enumerate(sec_counts.items(), 1):
            pct = count / total_articles if total_articles > 0 else 0
            ws_s.cell(row=row_idx, column=1, value=rank).alignment = Alignment(horizontal="center")
            ws_s.cell(row=row_idx, column=2, value=sec).alignment = Alignment(horizontal="left")
            ws_s.cell(row=row_idx, column=3, value=count).alignment = Alignment(horizontal="right")
            ws_s.cell(row=row_idx, column=4, value=pct).alignment = Alignment(horizontal="right")

            for col_idx in range(1, 5):
                cell = ws_s.cell(row=row_idx, column=col_idx)
                cell.font = regular_font
                cell.border = thin_border
                if col_idx == 3:
                    cell.number_format = "#,##0"
                elif col_idx == 4:
                    cell.number_format = "0.0%"
            row_idx += 1

        # Section Grand Total
        ws_s.cell(row=row_idx, column=1, value="").alignment = Alignment(horizontal="center")
        ws_s.cell(row=row_idx, column=2, value="Grand Total").alignment = Alignment(horizontal="left")
        ws_s.cell(row=row_idx, column=3, value=total_articles).alignment = Alignment(horizontal="right")
        ws_s.cell(row=row_idx, column=4, value=1.0).alignment = Alignment(horizontal="right")
        for col_idx in range(1, 5):
            cell = ws_s.cell(row=row_idx, column=col_idx)
            cell.fill = total_fill
            cell.font = total_font
            cell.border = total_border
            if col_idx == 3:
                cell.number_format = "#,##0"
            elif col_idx == 4:
                cell.number_format = "0.0%"

        # SHEET 5: Daily Section Matrix (Pivot Table)
        ws_matrix = wb.create_sheet(title="Daily Section Matrix")
        ws_matrix.views.sheetView[0].showGridLines = True

        ws_matrix["A1"] = "The Indian Express — Daily Section-Wise Matrix"
        ws_matrix["A1"].font = title_font
        ws_matrix["A2"] = "Breakdown of published stories per section by date"
        ws_matrix["A2"].font = subtitle_font

        pivot_d = pd.crosstab(index=df['date'], columns=df['section'], margins=True, margins_name="Total")
        matrix_headers = ["Date"] + list(pivot_d.columns)
        for col_idx, h in enumerate(matrix_headers, 1):
            c = ws_matrix.cell(row=4, column=col_idx, value=h)
            c.fill = navy_fill
            c.font = header_font
            c.alignment = Alignment(horizontal="center" if col_idx == 1 else "right", vertical="center")
        ws_matrix.row_dimensions[4].height = 24

        row_idx = 5
        for dt_val, row_data in pivot_d.iterrows():
            is_total = (dt_val == "Total")
            date_label = "Grand Total" if is_total else str(dt_val)
            ws_matrix.cell(row=row_idx, column=1, value=date_label).alignment = Alignment(horizontal="center" if not is_total else "left")

            for col_idx, count_val in enumerate(row_data, 2):
                cell = ws_matrix.cell(row=row_idx, column=col_idx, value=int(count_val))
                cell.alignment = Alignment(horizontal="right")
                cell.number_format = "#,##0"

            for col_idx in range(1, len(matrix_headers) + 1):
                cell = ws_matrix.cell(row=row_idx, column=col_idx)
                if is_total:
                    cell.fill = total_fill
                    cell.font = total_font
                    cell.border = total_border
                else:
                    cell.font = regular_font
                    cell.border = thin_border
                    if col_idx == len(matrix_headers):
                        cell.font = Font(name="Calibri", size=11, bold=True)
            row_idx += 1

    # -------------------------------------------------------------
    # SHEET 6: Full Article Metadata (URLs, Headlines, Publish Times)
    # -------------------------------------------------------------
    if include_details:
        ws_data = wb.create_sheet(title="All Articles Metadata")
        ws_data.views.sheetView[0].showGridLines = True

        headers_data = [
            "S.No",
            "Publish Date",
            "Day of Week",
            "First Publish Time",
            "Last Modified (ISO)",
            "Calendar Week",
            "Month",
            "Section",
            "Sub-Section",
            "Article Headline / Title",
            "Article URL"
        ]
        for col_idx, h in enumerate(headers_data, 1):
            c = ws_data.cell(row=1, column=col_idx, value=h)
            c.fill = navy_fill
            c.font = header_font
            c.alignment = Alignment(horizontal="center" if col_idx in [1, 2, 4] else "left", vertical="center")
        ws_data.row_dimensions[1].height = 22

        for r_idx, art in enumerate(df.to_dict('records'), 2):
            ws_data.cell(row=r_idx, column=1, value=r_idx - 1).alignment = Alignment(horizontal="center")
            ws_data.cell(row=r_idx, column=2, value=art['publish_date']).alignment = Alignment(horizontal="center")
            ws_data.cell(row=r_idx, column=3, value=art['day_of_week']).alignment = Alignment(horizontal="left")
            ws_data.cell(row=r_idx, column=4, value=art['first_publish_time']).alignment = Alignment(horizontal="center")
            ws_data.cell(row=r_idx, column=5, value=art['last_modified_time']).alignment = Alignment(horizontal="left")
            ws_data.cell(row=r_idx, column=6, value=art['week_label']).alignment = Alignment(horizontal="left")
            ws_data.cell(row=r_idx, column=7, value=art['month_name']).alignment = Alignment(horizontal="left")
            ws_data.cell(row=r_idx, column=8, value=art['section']).alignment = Alignment(horizontal="left")
            ws_data.cell(row=r_idx, column=9, value=art['sub_section']).alignment = Alignment(horizontal="left")
            ws_data.cell(row=r_idx, column=10, value=art['headline']).alignment = Alignment(horizontal="left")
            
            url_cell = ws_data.cell(row=r_idx, column=11, value=art['url'])
            url_cell.alignment = Alignment(horizontal="left")
            url_cell.font = Font(name="Calibri", size=10, color="0563C1", underline="single")

            for c_idx in range(1, 11):
                ws_data.cell(row=r_idx, column=c_idx).font = regular_font

    # Automatically set proportional column widths on all generated sheets
    for ws in wb.worksheets:
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.row in [1, 2] and ws.title != "All Articles Metadata":
                    continue
                val = str(cell.value or '')
                max_len = max(max_len, len(val))
            ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 65)

    wb.save(filepath)
    if isinstance(filepath, str):
        print(f"\nExcel Report successfully saved to:\n  -> {filepath}")


def run_pipeline(start_date_str: str, end_date_str: str, output_path: str = None, max_workers: int = 6,
                 include_daily: bool = True, include_weekly: bool = True,
                 include_monthly: bool = True, include_sections: bool = True,
                 include_details: bool = True):
    """
    Core end-to-end processing pipeline from date strings to terminal summaries and Excel reports.
    """
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

    # 1. Print terminal summaries
    display_terminal_summary(
        df, start_date, end_date,
        include_daily=include_daily,
        include_weekly=include_weekly,
        include_monthly=include_monthly,
        include_sections=include_sections
    )

    # 2. Export Excel workbook
    if not output_path:
        filename = f"IE_Story_Count_{start_date.strftime('%Y%m%d')}_to_{end_date.strftime('%Y%m%d')}.xlsx"
        output_path = os.path.join(os.getcwd(), filename)

    export_to_excel(
        df, start_date, end_date, output_path,
        include_daily=include_daily,
        include_weekly=include_weekly,
        include_monthly=include_monthly,
        include_sections=include_sections,
        include_details=include_details
    )


def main():
    parser = argparse.ArgumentParser(
        description="The Indian Express — Daily, Weekly, Monthly & Section-wise Story Count Software"
    )
    parser.add_argument("--start", "-s", type=str, help="Start Date (e.g. 2026-11-01 or 01-11-2026)")
    parser.add_argument("--end", "-e", type=str, help="End Date (e.g. 2026-11-30 or 30-11-2026)")
    parser.add_argument("--output", "-o", type=str, help="Custom output Excel filename / path (.xlsx)")
    parser.add_argument("--workers", "-w", type=int, default=6, help="Number of concurrent worker threads (default: 6)")
    parser.add_argument("--no-daily", action="store_true", help="Exclude Daily count")
    parser.add_argument("--no-weekly", action="store_true", help="Exclude Weekly count")
    parser.add_argument("--no-monthly", action="store_true", help="Exclude Monthly count")
    parser.add_argument("--no-sections", action="store_true", help="Exclude Section-wise breakdown")
    parser.add_argument("--no-details", action="store_true", help="Exclude Article metadata & URLs sheet")

    args = parser.parse_args()

    # Interactive prompt mode when run without arguments
    if not args.start or not args.end:
        print("=" * 68)
        print("   THE INDIAN EXPRESS — STORY COUNT & ARTICLE ANALYZER   ")
        print("=" * 68)
        print("Supported date formats: YYYY-MM-DD or DD-MM-YYYY (e.g. 2026-11-01 or 01-11-2026)\n")

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

        print("\n--------------------------------------------------------------------")
        print(" SELECT STORY COUNT MODULES (Select 2 or more, e.g. 1,2 or 1,2,3):")
        print("--------------------------------------------------------------------")
        print("  [1] Daily Story Count (Date-wise)")
        print("  [2] Weekly Story Count (Week-wise)")
        print("  [3] Monthly Story Count (Month-wise)")
        print("  [4] Section-Wise Breakdown (Desk counts & Daily Matrix)")
        print("  [5] Full Article Metadata (URLs, Headlines, Publish & Modified Times)")
        print("  [A] Select ALL Modules (Default - Press Enter)")
        print("--------------------------------------------------------------------")
        selection = input("Enter your choices (e.g. 1,2 or 1,2,4 or press Enter for ALL): ").strip()

        if not selection or selection.lower() == 'a':
            inc_daily = True
            inc_weekly = True
            inc_monthly = True
            inc_sections = True
            inc_details = True
        else:
            chosen = [item.strip() for item in re.split(r'[, ]+', selection) if item.strip()]
            inc_daily = '1' in chosen
            inc_weekly = '2' in chosen
            inc_monthly = '3' in chosen
            inc_sections = '4' in chosen
            inc_details = '5' in chosen

            if not any([inc_daily, inc_weekly, inc_monthly, inc_sections, inc_details]):
                print("No numbers recognized. Defaulting to ALL modules.")
                inc_daily = inc_weekly = inc_monthly = inc_sections = inc_details = True

        run_pipeline(
            start_input, end_input,
            output_path=args.output,
            max_workers=args.workers,
            include_daily=inc_daily,
            include_weekly=inc_weekly,
            include_monthly=inc_monthly,
            include_sections=inc_sections,
            include_details=inc_details
        )
    else:
        run_pipeline(
            args.start, args.end,
            output_path=args.output,
            max_workers=args.workers,
            include_daily=not args.no_daily,
            include_weekly=not args.no_weekly,
            include_monthly=not args.no_monthly,
            include_sections=not args.no_sections,
            include_details=not args.no_details
        )


if __name__ == "__main__":
    main()
