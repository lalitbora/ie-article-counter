#!/usr/bin/env python3
"""
Indian Express Group — Multi-Property Story Count & Article Analyzer
=====================================================================
Comprehensive sitemap analyzer supporting all Indian Express Group publications:
  1. The Indian Express (https://indianexpress.com/)
  2. Financial Express (https://www.financialexpress.com/)
  3. Loksatta (https://www.loksatta.com/)
  4. Jansatta (https://www.jansatta.com/)
  5. IE Tamil (https://tamil.indianexpress.com/)
  6. IE Bengali (https://bengali.indianexpress.com/)
  7. IE Gujarati (https://gujarati.indianexpress.com/)
  8. IE Malayalam (https://malayalam.indianexpress.com/)

Calculates:
  - Daily Story Count (Date-wise)
  - Weekly Story Count (ISO calendar week-wise)
  - Monthly Story Count (Month-wise)
  - Section-wise Story Count (Desk breakdowns and daily matrix)
  - Full Article Metadata (URLs, Headlines, Publish Date/Time, Last Modified Time)
  - Styled Multi-Sheet Excel Reports (.xlsx)

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

# Ensure UTF-8 output encoding to avoid Windows charmap encoding errors
try:
    if sys.stdout.encoding != 'utf-8':
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Import requests and curl_cffi for resilient multi-tier requests
import requests
_SESSION = requests.Session()

try:
    from curl_cffi import requests as cffi_requests
    _USE_CURL_CFFI = True
except ImportError:
    _USE_CURL_CFFI = False
    cffi_requests = None

# Realistic Chrome headers for fallback HTTP requests
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

# Configuration for all Indian Express Group publications
PROPERTIES = {
    'indianexpress': {
        'name': 'The Indian Express',
        'url': 'https://indianexpress.com/',
        'sitemap_type': 'query',  # /sitemap.xml?yyyy=YYYY&mm=MM&dd=DD
        'sitemap_base': 'https://indianexpress.com/sitemap.xml',
    },
    'financialexpress': {
        'name': 'Financial Express',
        'url': 'https://www.financialexpress.com/',
        'sitemap_type': 'query',  # /sitemap.xml?yyyy=YYYY&mm=MM&dd=DD
        'sitemap_base': 'https://www.financialexpress.com/sitemap.xml',
    },
    'loksatta': {
        'name': 'Loksatta',
        'url': 'https://www.loksatta.com/',
        'sitemap_type': 'query',  # /sitemap.xml?yyyy=YYYY&mm=MM&dd=DD
        'sitemap_base': 'https://www.loksatta.com/sitemap.xml',
    },
    'jansatta': {
        'name': 'Jansatta',
        'url': 'https://www.jansatta.com/',
        'sitemap_type': 'query',  # /sitemap.xml?yyyy=YYYY&mm=MM&dd=DD
        'sitemap_base': 'https://www.jansatta.com/sitemap.xml',
    },
    'ietamil': {
        'name': 'IE Tamil',
        'url': 'https://tamil.indianexpress.com/',
        'sitemap_type': 'date_xml',  # /sitemap_YYYY-MM-DD.xml
        'sitemap_base': 'https://tamil.indianexpress.com/sitemap_',
    },
    'iebengali': {
        'name': 'IE Bengali',
        'url': 'https://bengali.indianexpress.com/',
        'sitemap_type': 'date_xml',  # /sitemap_YYYY-MM-DD.xml
        'sitemap_base': 'https://bengali.indianexpress.com/sitemap_',
    },
    'iegujarati': {
        'name': 'IE Gujarati',
        'url': 'https://gujarati.indianexpress.com/',
        'sitemap_type': 'date_xml',  # /sitemap_YYYY-MM-DD.xml
        'sitemap_base': 'https://gujarati.indianexpress.com/sitemap_',
    },
    'iemalayalam': {
        'name': 'IE Malayalam',
        'url': 'https://malayalam.indianexpress.com/',
        'sitemap_type': 'date_xml',  # /sitemap_YYYY-MM-DD.xml
        'sitemap_base': 'https://malayalam.indianexpress.com/sitemap_',
    }
}

# Mapping of section URL slugs to canonical display titles for English properties
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
    'news-today': 'News Today',
    'india-news': 'India News',
    'market': 'Market',
    'industry': 'Industry',
    'banking-finance': 'Banking & Finance',
    'auto': 'Auto',
    'money': 'Personal Finance / Money',
    'horoscope': 'Horoscope'
}

CITY_SLUGS = {
    'delhi', 'mumbai', 'chandigarh', 'pune', 'bangalore', 'bengaluru',
    'kolkata', 'lucknow', 'ahmedabad', 'hyderabad', 'chennai', 'jaipur',
    'patna', 'bhopal', 'shimla', 'dehradun', 'thiruvananthapuram', 'kochi',
    'nagpur', 'varanasi', 'kanpur', 'surat', 'ludhiana', 'amritsar',
    'gurgaon', 'noida', 'ghaziabad', 'faridabad'
}

WORLD_SLUGS = {'pakistan', 'world'}


def parse_flexible_date(date_str: str) -> datetime.date:
    """Parse flexible user-entered date strings into a datetime.date object."""
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
        f"Invalid date format: '{date_str}'. Please use YYYY-MM-DD or DD-MM-YYYY (e.g. 2024-12-02 or 02-12-2024)."
    )


def extract_title_from_slug(slug: str) -> str:
    """Generate a readable headline from an article's URL slug."""
    if not slug:
        return ""
    clean_slug = re.sub(r'-\d+$', '', slug)
    clean_slug = re.sub(r'/\d+$', '', clean_slug)
    words = clean_slug.replace('-', ' ').split()
    return ' '.join(word.capitalize() for word in words)


def parse_article_url(url: str, property_key: str = 'indianexpress'):
    """
    Parse section, sub-section, and headline from any Express property article URL.
    Returns: (section_clean, sub_section_clean, headline)
    """
    parsed = urlparse(url)
    parts = [p for p in parsed.path.strip('/').split('/') if p]

    if not parts:
        return 'General', '-', ''

    # Handle /article/ prefix in standard Indian Express URLs
    if parts[0] == 'article':
        parts = parts[1:]

    if not parts:
        return 'General', '-', ''

    # Check if last segment is a pure numeric article ID (e.g. /3682494/ or /7684395)
    if parts[-1].isdigit():
        slug = parts[-2] if len(parts) >= 2 else parts[-1]
        sec_parts = parts[:-2]
    else:
        slug = parts[-1]
        sec_parts = parts[:-1]

    if len(sec_parts) >= 2:
        sec = sec_parts[0].lower()
        sub = sec_parts[1].lower()
    elif len(sec_parts) == 1:
        sec = sec_parts[0].lower()
        sub = '-'
    else:
        sec = 'general'
        sub = '-'

    # Specific normalization for The Indian Express cities and world
    if property_key == 'indianexpress':
        if sec in CITY_SLUGS:
            sub = sec.capitalize()
            sec = 'cities'
        elif sec in WORLD_SLUGS and sec != 'world':
            sub = sec.capitalize()
            sec = 'world'

    section_name = SECTION_TITLE_MAP.get(
        sec,
        sec.replace('-', ' ').title()
    )
    sub_section_name = sub.replace('-', ' ').title() if sub != '-' else '-'
    headline = extract_title_from_slug(slug)

    return section_name, sub_section_name, headline


def compute_week_info(date_obj: datetime.date):
    """Calculate ISO calendar week details and format a human-readable label."""
    iso_year, iso_week, _ = date_obj.isocalendar()
    monday = date_obj - timedelta(days=date_obj.weekday())
    sunday = monday + timedelta(days=6)
    week_key = f"{iso_year}-W{iso_week:02d}"
    week_label = f"Week {iso_week:02d} ({monday.strftime('%d-%b')} to {sunday.strftime('%d-%b-%Y')})"
    return week_key, week_label, monday, sunday


def get_sitemap_url(property_key: str, target_date: datetime.date) -> str:
    """Generate the exact daily sitemap URL for the selected publication property."""
    prop = PROPERTIES.get(property_key, PROPERTIES['indianexpress'])
    stype = prop.get('sitemap_type', 'query')
    yyyy = target_date.year
    mm = f"{target_date.month:02d}"
    dd = f"{target_date.day:02d}"

    if stype == 'query':
        return f"{prop['sitemap_base']}?yyyy={yyyy}&mm={mm}&dd={dd}"
    elif stype == 'date_xml':
        return f"{prop['sitemap_base']}{yyyy}-{mm}-{dd}.xml"
    else:
        return f"{prop['sitemap_base']}?yyyy={yyyy}&mm={mm}&dd={dd}"


def fetch_raw_sitemap(sitemap_url: str, timeout: int = 20):
    """
    Fetch sitemap XML using multi-tiered bypass techniques:
      Tier 1: Googlebot search crawler headers (Cloudflare sitemap WAF explicitly whitelists for sitemaps)
      Tier 2: curl_cffi with authentic Chrome 124 TLS & HTTP/2 impersonation
      Tier 3: Bingbot search crawler headers
      Tier 4: Realistic Windows Chrome browser headers
    """
    # Tier 1: Googlebot (whitelisted by Cloudflare WAF specifically for XML sitemaps)
    try:
        headers_bot = {
            'User-Agent': 'Mozilla/5.0 (compatible; Googlebot/2.1; +http://www.google.com/bot.html)',
            'Accept': 'text/xml,application/xml,application/xhtml+xml,text/html;q=0.9,*/*;q=0.8',
            'Accept-Encoding': 'gzip, deflate',
        }
        resp = _SESSION.get(sitemap_url, headers=headers_bot, timeout=timeout)
        if resp.status_code == 200 and resp.content:
            return resp.status_code, resp.content, None
        elif resp.status_code == 404:
            return 404, None, "No sitemap found (404)"
    except Exception:
        pass

    # Tier 2: curl_cffi with authentic Chrome 124 TLS & HTTP/2 impersonation
    if _USE_CURL_CFFI and cffi_requests is not None:
        try:
            resp = cffi_requests.get(sitemap_url, impersonate="chrome124", timeout=timeout)
            if resp.status_code == 200 and resp.content:
                return resp.status_code, resp.content, None
            elif resp.status_code == 404:
                return 404, None, "No sitemap found (404)"
        except Exception:
            pass

    # Tier 3: Bingbot search crawler
    try:
        headers_bing = {
            'User-Agent': 'Mozilla/5.0 (compatible; bingbot/2.0; +http://www.bing.com/bingbot.htm)',
            'Accept': 'text/xml,application/xml,*/*',
            'Accept-Encoding': 'gzip, deflate',
        }
        resp = _SESSION.get(sitemap_url, headers=headers_bing, timeout=timeout)
        if resp.status_code == 200 and resp.content:
            return resp.status_code, resp.content, None
        elif resp.status_code == 404:
            return 404, None, "No sitemap found (404)"
    except Exception:
        pass

    # Tier 4: Realistic Chrome Desktop Session
    try:
        resp = _SESSION.get(sitemap_url, headers=HEADERS, timeout=timeout)
        if resp.status_code == 200 and resp.content:
            return resp.status_code, resp.content, None
        elif resp.status_code == 404:
            return 404, None, "No sitemap found (404)"
        elif resp.status_code == 403:
            return 403, None, "HTTP 403 Forbidden (Cloudflare bot block)"
        else:
            return resp.status_code, None, f"HTTP {resp.status_code}"
    except Exception as e:
        return 0, None, str(e)


def fetch_sitemap_for_date(target_date: datetime.date, property_key: str = 'indianexpress', max_retries: int = 3):
    """
    Fetch and parse the daily XML sitemap for a specific date and property.
    Returns: (target_date, list_of_articles, error_msg)
    """
    sitemap_url = get_sitemap_url(property_key, target_date)
    articles = []
    error_msg = None

    week_key, week_label, monday, sunday = compute_week_info(target_date)

    for attempt in range(1, max_retries + 1):
        status_code, content, err = fetch_raw_sitemap(sitemap_url, timeout=20)

        if status_code == 404:
            error_msg = f"No sitemap found (404) for {target_date}"
            break
        elif status_code != 200 or not content:
            error_msg = err or f"HTTP {status_code}"
            if attempt < max_retries:
                import time
                time.sleep(1.0 * attempt)
            continue

        try:
            # Check for brotli compressed content
            if content[:3] != b'<?x':
                try:
                    import brotli
                    content = brotli.decompress(content)
                except Exception:
                    pass

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

                    sec, sub_sec, headline = parse_article_url(url_str, property_key=property_key)
                    prop_name = PROPERTIES.get(property_key, PROPERTIES['indianexpress'])['name']
                    full_title = f"{headline} | {sec} - {prop_name}" if sub_sec == '-' else f"{headline} | {sub_sec} - {prop_name}"

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
                        'title': full_title,
                        'headline': headline,
                        'url': url_str
                    })

            error_msg = None
            break

        except Exception as e:
            error_msg = f"XML Parse Error: {e}"

        if attempt < max_retries:
            import time
            time.sleep(1.0 * attempt)

    return target_date, articles, error_msg


def fetch_all_dates(start_date: datetime.date, end_date: datetime.date,
                    property_key: str = 'indianexpress', max_workers: int = 6):
    """Fetch stories for all dates in range concurrently for the selected property."""
    prop_info = PROPERTIES.get(property_key, PROPERTIES['indianexpress'])
    total_days = (end_date - start_date).days + 1
    dates = [start_date + timedelta(days=i) for i in range(total_days)]

    print(f"\n" + "=" * 70)
    print(f"  FETCHING STORIES: {prop_info['name'].upper()}")
    print(f"  Property: {prop_info['url']}")
    print(f"  Period  : {start_date.strftime('%d-%b-%Y')} to {end_date.strftime('%d-%b-%Y')} ({total_days} days)")
    print(f"  Threads : {max_workers} concurrent network workers")
    print("=" * 70)

    all_articles = []
    date_status = {}
    completed_count = 0

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_to_date = {
            executor.submit(fetch_sitemap_for_date, d, property_key): d for d in dates
        }

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
                             property_name: str = "The Indian Express",
                             include_daily: bool = True, include_weekly: bool = True,
                             include_monthly: bool = True, include_sections: bool = True):
    """Display comprehensive terminal tables based on selected checkboxes and property."""
    if df.empty:
        print(f"\nNo articles found for {property_name} in the specified date range.")
        return

    total_articles = len(df)
    total_days = (end_date - start_date).days + 1
    daily_avg = total_articles / total_days if total_days > 0 else 0

    print("\n" + "=" * 70)
    print(f"       {property_name.upper()} — STORY COUNT ANALYSIS REPORT")
    print("=" * 70)
    print(f"  Period        : {start_date.strftime('%d-%b-%Y')} to {end_date.strftime('%d-%b-%Y')}")
    print(f"  Total Days    : {total_days}")
    print(f"  Total Stories : {total_articles:,}")
    print(f"  Daily Average : {daily_avg:.1f} stories/day")
    print("=" * 70)

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
            f"1. {property_name} — Monthly Story Count Summary",
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
            f"2. {property_name} — Weekly Story Count Summary",
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
            f"3. {property_name} — Daily Total Story Count",
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
            f"4. {property_name} — Section-Wise Story Count (Overall)",
            ["Section Name", "Story Count", "Share (%)"],
            section_rows,
            alignments=['left', 'right', 'right']
        )


def export_to_excel(df: pd.DataFrame, start_date: datetime.date, end_date: datetime.date, filepath,
                    property_name: str = "The Indian Express",
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
    wb.remove(wb.active)

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

    # SHEET 1: Monthly Summary
    if include_monthly:
        ws_m = wb.create_sheet(title="Monthly Summary")
        ws_m.views.sheetView[0].showGridLines = True

        ws_m["A1"] = f"{property_name} — Monthly Story Count Summary"
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

    # SHEET 2: Weekly Summary
    if include_weekly:
        ws_w = wb.create_sheet(title="Weekly Summary")
        ws_w.views.sheetView[0].showGridLines = True

        ws_w["A1"] = f"{property_name} — Weekly Story Count Summary"
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

    # SHEET 3: Daily Total Count
    if include_daily:
        ws_d = wb.create_sheet(title="Daily Total Count")
        ws_d.views.sheetView[0].showGridLines = True

        ws_d["A1"] = f"{property_name} — Daily Total Story Count"
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

    # SHEET 4: Section Summary
    if include_sections:
        ws_s = wb.create_sheet(title="Section Summary")
        ws_s.views.sheetView[0].showGridLines = True

        ws_s["A1"] = f"{property_name} — Section-Wise Story Count"
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

        # Daily Section Matrix
        ws_matrix = wb.create_sheet(title="Daily Section Matrix")
        ws_matrix.views.sheetView[0].showGridLines = True

        ws_matrix["A1"] = f"{property_name} — Daily Section-Wise Matrix"
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

    # SHEET 6: All Articles Metadata
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
            "Article Title",
            "Article Headline",
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
            ws_data.cell(row=r_idx, column=10, value=art.get('title', '')).alignment = Alignment(horizontal="left")
            ws_data.cell(row=r_idx, column=11, value=art.get('headline', '')).alignment = Alignment(horizontal="left")

            url_cell = ws_data.cell(row=r_idx, column=12, value=art['url'])
            url_cell.alignment = Alignment(horizontal="left")
            url_cell.font = Font(name="Calibri", size=10, color="0563C1", underline="single")

            for c_idx in range(1, 12):
                ws_data.cell(row=r_idx, column=c_idx).font = regular_font

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


def run_pipeline(start_date_str: str, end_date_str: str,
                 property_key: str = 'indianexpress',
                 output_path: str = None, max_workers: int = 6,
                 include_daily: bool = True, include_weekly: bool = True,
                 include_monthly: bool = True, include_sections: bool = True,
                 include_details: bool = True):
    """End-to-end execution flow for the chosen property and date range."""
    start_date = parse_flexible_date(start_date_str)
    end_date = parse_flexible_date(end_date_str)

    if start_date > end_date:
        print(f"\n[Notice] Start Date ({start_date}) was after End Date ({end_date}). Swapping automatically.")
        start_date, end_date = end_date, start_date

    prop_info = PROPERTIES.get(property_key, PROPERTIES['indianexpress'])
    all_articles, date_status = fetch_all_dates(start_date, end_date, property_key=property_key, max_workers=max_workers)

    if not all_articles:
        print(f"\nNo articles could be fetched for {prop_info['name']} in this date range.")
        return

    df = pd.DataFrame(all_articles)

    display_terminal_summary(
        df, start_date, end_date,
        property_name=prop_info['name'],
        include_daily=include_daily,
        include_weekly=include_weekly,
        include_monthly=include_monthly,
        include_sections=include_sections
    )

    if not output_path:
        clean_name = prop_info['name'].replace(' ', '_')
        filename = f"{clean_name}_Story_Count_{start_date.strftime('%Y%m%d')}_to_{end_date.strftime('%Y%m%d')}.xlsx"
        output_path = os.path.join(os.getcwd(), filename)

    export_to_excel(
        df, start_date, end_date, output_path,
        property_name=prop_info['name'],
        include_daily=include_daily,
        include_weekly=include_weekly,
        include_monthly=include_monthly,
        include_sections=include_sections,
        include_details=include_details
    )


def main():
    parser = argparse.ArgumentParser(
        description="Indian Express Group — Multi-Property Story Count & Article Analyzer"
    )
    parser.add_argument("--property", "-p", type=str, default="indianexpress",
                        choices=list(PROPERTIES.keys()),
                        help="Publication property (indianexpress, financialexpress, loksatta, jansatta, ietamil, iebengali, iegujarati, iemalayalam)")
    parser.add_argument("--start", "-s", type=str, help="Start Date (e.g. 2024-12-01 or 01-12-2024)")
    parser.add_argument("--end", "-e", type=str, help="End Date (e.g. 2024-12-15 or 15-12-2024)")
    parser.add_argument("--output", "-o", type=str, help="Custom output Excel filename / path (.xlsx)")
    parser.add_argument("--workers", "-w", type=int, default=6, help="Number of concurrent worker threads (default: 6)")
    parser.add_argument("--no-daily", action="store_true", help="Exclude Daily count")
    parser.add_argument("--no-weekly", action="store_true", help="Exclude Weekly count")
    parser.add_argument("--no-monthly", action="store_true", help="Exclude Monthly count")
    parser.add_argument("--no-sections", action="store_true", help="Exclude Section-wise breakdown")
    parser.add_argument("--no-details", action="store_true", help="Exclude Article metadata & URLs sheet")

    args = parser.parse_args()

    if not args.start or not args.end:
        print("=" * 70)
        print("    INDIAN EXPRESS GROUP — MULTI-PROPERTY STORY COUNT SOFTWARE")
        print("=" * 70)
        print("\nAvailable Properties:")
        prop_keys = list(PROPERTIES.keys())
        for idx, k in enumerate(prop_keys, 1):
            print(f"  [{idx}] {PROPERTIES[k]['name']} ({PROPERTIES[k]['url']})")

        p_choice = input(f"\nSelect Property [1-{len(prop_keys)}, Default: 1 - The Indian Express]: ").strip()
        selected_prop = 'indianexpress'
        if p_choice.isdigit() and 1 <= int(p_choice) <= len(prop_keys):
            selected_prop = prop_keys[int(p_choice) - 1]

        print(f"\nSelected: {PROPERTIES[selected_prop]['name']}")
        print("Supported date formats: YYYY-MM-DD or DD-MM-YYYY (e.g. 2024-12-01 or 01-12-2024)\n")

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
            inc_daily = inc_weekly = inc_monthly = inc_sections = inc_details = True
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
            property_key=selected_prop,
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
            property_key=args.property,
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
