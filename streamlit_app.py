#!/usr/bin/env python3
"""
The Indian Express — Story Count Streamlit Cloud Web Application
=================================================================
Interactive Streamlit application featuring:
  - Checkbox selectors for Daily, Weekly, Monthly, Section-wise, and Full Metadata
  - Start Date and End Date range pickers with quick shortcuts
  - Real-time tabular metrics, charts, and pivot matrices
  - One-click Excel (.xlsx) report generation and download
  - Full compatibility with Streamlit Community Cloud and local environments

Author: Antigravity Assistant
Environment: Python 3.8+
"""

import sys
import io
from datetime import datetime, timedelta
import pandas as pd
import streamlit as st

# Import core extraction logic
from count_articles import (
    fetch_all_dates,
    parse_flexible_date,
    export_to_excel,
    PROPERTIES
)

st.set_page_config(
    page_title="Indian Express Group Story Counter",
    page_icon="📰",
    layout="wide"
)

# Custom Styling
st.markdown("""
<style>
    .main-title {
        font-size: 2.2rem;
        font-weight: 700;
        color: #1F4E78;
        margin-bottom: 0.2rem;
    }
    .sub-title {
        color: #555;
        font-size: 1rem;
        margin-bottom: 1.5rem;
    }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">📰 Indian Express Group — Multi-Property Story Counter</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Daily, Weekly, Monthly & Section-wise Story Count Software (Live XML Sitemaps)</div>', unsafe_allow_html=True)

today = datetime.now().date()

# Property Selection Mapping
PROPERTY_CHOICES = {
    "📰 The Indian Express (indianexpress.com)": "indianexpress",
    "📈 Financial Express (financialexpress.com)": "financialexpress",
    "🗞️ Loksatta (loksatta.com)": "loksatta",
    "🇮🇳 Jansatta (jansatta.com)": "jansatta",
    "🏛️ IE Tamil (tamil.indianexpress.com)": "ietamil",
    "✒️ IE Bengali (bengali.indianexpress.com)": "iebengali",
    "🦁 IE Gujarati (gujarati.indianexpress.com)": "iegujarati",
    "🌴 IE Malayalam (malayalam.indianexpress.com)": "iemalayalam",
}

query_params = st.query_params

default_prop_idx = 0
if "property" in query_params:
    p_val = query_params["property"].lower()
    for idx, (lbl, k) in enumerate(PROPERTY_CHOICES.items()):
        if k == p_val:
            default_prop_idx = idx
            break

# Property Dropdown Selector
selected_prop_label = st.selectbox(
    "🌐 Select Indian Express Publication / Property:",
    options=list(PROPERTY_CHOICES.keys()),
    index=default_prop_idx
)
selected_prop_key = PROPERTY_CHOICES[selected_prop_label]
prop_info = PROPERTIES[selected_prop_key]

# Initialize session state for date range
if "start_val" not in st.session_state:
    st.session_state["start_val"] = today - timedelta(days=6)
if "end_val" not in st.session_state:
    st.session_state["end_val"] = today

# Handle URL query parameters if present
query_params = st.query_params
if "start" in query_params:
    try:
        st.session_state["start_val"] = parse_flexible_date(query_params["start"])
    except Exception:
        pass
if "end" in query_params:
    try:
        st.session_state["end_val"] = parse_flexible_date(query_params["end"])
    except Exception:
        pass

# Quick shortcuts handlers
def set_quick_dates(start_offset, end_offset):
    st.session_state["start_val"] = today - timedelta(days=start_offset)
    st.session_state["end_val"] = today - timedelta(days=end_offset)

def set_month_dates():
    st.session_state["start_val"] = today.replace(day=1)
    st.session_state["end_val"] = today

# Top Checkbox Selectors Panel
st.markdown("##### ⚙️ Select Story Count Modules to Include:")
chk_col1, chk_col2, chk_col3, chk_col4, chk_col5 = st.columns(5)
with chk_col1:
    inc_daily = st.checkbox("Daily Count (Date-wise)", value=True)
with chk_col2:
    inc_weekly = st.checkbox("Weekly Story Count", value=True)
with chk_col3:
    inc_monthly = st.checkbox("Monthly Story Count", value=True)
with chk_col4:
    inc_sections = st.checkbox("Section Breakdown", value=True)
with chk_col5:
    inc_details = st.checkbox("All Article Metadata", value=True)

st.divider()

# Quick select shortcuts
st.caption("Quick Range Select:")
q1, q2, q3, q4, q5 = st.columns(5)
with q1:
    if st.button("Today", use_container_width=True):
        set_quick_dates(0, 0)
        st.rerun()
with q2:
    if st.button("Yesterday", use_container_width=True):
        set_quick_dates(1, 1)
        st.rerun()
with q3:
    if st.button("Last 7 Days", use_container_width=True):
        set_quick_dates(6, 0)
        st.rerun()
with q4:
    if st.button("Last 14 Days", use_container_width=True):
        set_quick_dates(13, 0)
        st.rerun()
with q5:
    if st.button("This Month", use_container_width=True):
        set_month_dates()
        st.rerun()

# Date Pickers & Fetch Button
with st.container():
    col1, col2, col3 = st.columns([2, 2, 2])
    with col1:
        start_date = st.date_input("Start Date", value=st.session_state["start_val"], key="input_start")
    with col2:
        end_date = st.date_input("End Date", value=st.session_state["end_val"], key="input_end")
    with col3:
        st.write("")
        st.write("")
        fetch_btn = st.button("📊 Fetch & Calculate", type="primary", use_container_width=True)

# Main Execution Trigger
if fetch_btn or ("start" in query_params and "end" in query_params):
    if start_date > end_date:
        start_date, end_date = end_date, start_date

    total_days = (end_date - start_date).days + 1

    with st.spinner(f"Fetching stories from {prop_info['name']} sitemaps for {total_days} day(s)..."):
        all_articles, date_status = fetch_all_dates(start_date, end_date, property_key=selected_prop_key, max_workers=6)

    if not all_articles:
        st.error(f"No articles could be fetched for {prop_info['name']} during {start_date} to {end_date}.")
        with st.expander("🔍 Click to view Connection / Date-wise Status", expanded=True):
            st.write("Date-wise Network Request Status:")
            for d, status in date_status.items():
                if "Error" in str(status) or "Exception" in str(status):
                    st.error(f"**{d}**: {status}")
                else:
                    st.info(f"**{d}**: {status}")
            st.warning("💡 **Pro-Tip:** If cloud datacenter IP blocks occur, run the local software via `launch_software.bat` or `start_public_live_url.bat` on your PC, which routes through your local ISP connection and is 100% immune to Cloudflare blocks!")
    else:
        df = pd.DataFrame(all_articles)
        total_articles = len(df)
        daily_avg = total_articles / total_days if total_days > 0 else 0
        top_section = df['section'].value_counts().index[0] if not df.empty else "N/A"
        top_section_count = df['section'].value_counts().iloc[0] if not df.empty else 0

        # KPI Metrics
        st.divider()
        m1, m2, m3, m4 = st.columns(4)
        m1.metric(f"Total Stories ({prop_info['name']})", f"{total_articles:,}")
        m2.metric("Daily Average", f"{daily_avg:.1f} / day")
        m3.metric("Days Analyzed", f"{total_days} Days")
        m4.metric("Top Section", f"{top_section}", f"{top_section_count:,} stories")

        # Excel Download Button
        buf = io.BytesIO()
        export_to_excel(
            df, start_date, end_date, buf,
            property_name=prop_info['name'],
            include_daily=inc_daily,
            include_weekly=inc_weekly,
            include_monthly=inc_monthly,
            include_sections=inc_sections,
            include_details=inc_details
        )
        excel_data = buf.getvalue()
        clean_name = prop_info['name'].replace(' ', '_')
        filename = f"{clean_name}_Story_Count_{start_date.strftime('%Y%m%d')}_to_{end_date.strftime('%Y%m%d')}.xlsx"

        st.download_button(
            label="📥 Download Excel Report (.xlsx)",
            data=excel_data,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

        st.divider()

        # Build Dynamic Tabs
        tab_names = []
        if inc_daily: tab_names.append("📅 Daily Story Count")
        if inc_weekly: tab_names.append("🗓️ Weekly Story Count")
        if inc_monthly: tab_names.append("📆 Monthly Summary")
        if inc_sections: tab_names.append("📊 Section-Wise Breakdown")
        if inc_sections: tab_names.append("🔲 Daily Section Matrix")
        if inc_details: tab_names.append("📄 All Articles Metadata")

        if tab_names:
            tabs = st.tabs(tab_names)
            tab_idx = 0

            # Daily Count Tab
            if inc_daily:
                with tabs[tab_idx]:
                    daily_groups = df.groupby(['date', 'day_of_week']).size().reset_index(name='Total Stories')
                    daily_groups = daily_groups.sort_values('date')
                    c_d1, c_d2 = st.columns([1, 1])
                    with c_d1:
                        st.dataframe(daily_groups, use_container_width=True, hide_index=True)
                    with c_d2:
                        st.line_chart(daily_groups.set_index('date')['Total Stories'], color="#2F855A")
                tab_idx += 1

            # Weekly Count Tab
            if inc_weekly:
                with tabs[tab_idx]:
                    weekly_rows = []
                    for (w_key, w_label), w_df in df.groupby(['week_key', 'week_label']):
                        w_count = len(w_df)
                        w_days = w_df['date'].nunique()
                        top_sec = w_df['section'].value_counts().index[0] if not w_df.empty else "-"
                        weekly_rows.append({
                            "Calendar Week": w_label,
                            "Days Active": w_days,
                            "Total Stories": w_count,
                            "Daily Average": round(w_count / w_days, 1) if w_days > 0 else 0,
                            "Top Section": top_sec
                        })
                    w_df_display = pd.DataFrame(weekly_rows)
                    c_w1, c_w2 = st.columns([1, 1])
                    with c_w1:
                        st.dataframe(w_df_display, use_container_width=True, hide_index=True)
                    with c_w2:
                        if not w_df_display.empty:
                            st.bar_chart(w_df_display.set_index('Calendar Week')['Total Stories'], color="#1F4E78")
                tab_idx += 1

            # Monthly Summary Tab
            if inc_monthly:
                with tabs[tab_idx]:
                    monthly_rows = []
                    for (m_code, m_name), m_df in df.groupby(['month', 'month_name']):
                        m_count = len(m_df)
                        m_days = m_df['date'].nunique()
                        monthly_rows.append({
                            "Month": m_name,
                            "Days Active": m_days,
                            "Total Stories": m_count,
                            "Daily Average": round(m_count / m_days, 1) if m_days > 0 else 0
                        })
                    st.dataframe(pd.DataFrame(monthly_rows), use_container_width=True, hide_index=True)
                tab_idx += 1

            # Section-Wise Breakdown Tab
            if inc_sections:
                with tabs[tab_idx]:
                    sec_counts = df['section'].value_counts().reset_index()
                    sec_counts.columns = ['Section Name', 'Story Count']
                    sec_counts['Share (%)'] = (sec_counts['Story Count'] / total_articles * 100).round(1)

                    c_s1, c_s2 = st.columns([1, 1])
                    with c_s1:
                        st.dataframe(sec_counts, use_container_width=True, hide_index=True)
                    with c_s2:
                        st.bar_chart(sec_counts.head(10).set_index('Section Name')['Story Count'], color="#1F4E78")
                tab_idx += 1

                # Daily Section Matrix Tab
                with tabs[tab_idx]:
                    pivot = pd.crosstab(index=df['date'], columns=df['section'], margins=True, margins_name="Grand Total")
                    st.dataframe(pivot, use_container_width=True)
                tab_idx += 1

            # All Articles Metadata Tab
            if inc_details:
                with tabs[tab_idx]:
                    st.dataframe(
                        df[['publish_date', 'day_of_week', 'first_publish_time', 'last_modified_time', 'section', 'sub_section', 'headline', 'url']],
                        use_container_width=True,
                        hide_index=True
                    )
                tab_idx += 1
