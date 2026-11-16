import sys
import io
from datetime import datetime, timedelta
import pandas as pd
import streamlit as st

# Import core extraction logic
from count_articles import (
    fetch_all_dates,
    parse_flexible_date,
    export_to_excel
)

st.set_page_config(
    page_title="Indian Express Article Counter",
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
    .kpi-card {
        background-color: #F8F9FA;
        border-radius: 8px;
        padding: 16px;
        border-left: 4px solid #1F4E78;
    }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">📰 The Indian Express — Article Counter</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Daily Section-Wise, Daily Total & Monthly Summary Analyzer (Live Sitemaps)</div>', unsafe_allow_html=True)

# Parse URL parameters if present
query_params = st.query_params
default_start = None
default_end = None

if "start" in query_params:
    try:
        default_start = parse_flexible_date(query_params["start"])
    except Exception:
        pass

if "end" in query_params:
    try:
        default_end = parse_flexible_date(query_params["end"])
    except Exception:
        pass

today = datetime.now().date()
if not default_start:
    default_start = today - timedelta(days=6)
if not default_end:
    default_end = today

# Filter Section
with st.container():
    col1, col2, col3 = st.columns([2, 2, 2])
    with col1:
        start_date = st.date_input("Start Date", value=default_start)
    with col2:
        end_date = st.date_input("End Date", value=default_end)
    with col3:
        st.write("")
        st.write("")
        fetch_btn = st.button("📊 Fetch & Calculate", type="primary", use_container_width=True)

# Quick shortcuts
st.caption("Quick Select:")
quick_cols = st.columns(5)
if quick_cols[0].button("Today", use_container_width=True):
    start_date = today
    end_date = today
    fetch_btn = True
if quick_cols[1].button("Yesterday", use_container_width=True):
    start_date = today - timedelta(days=1)
    end_date = today - timedelta(days=1)
    fetch_btn = True
if quick_cols[2].button("Last 7 Days", use_container_width=True):
    start_date = today - timedelta(days=6)
    end_date = today
    fetch_btn = True
if quick_cols[3].button("Last 14 Days", use_container_width=True):
    start_date = today - timedelta(days=13)
    end_date = today
    fetch_btn = True
if quick_cols[4].button("This Month", use_container_width=True):
    start_date = today.replace(day=1)
    end_date = today
    fetch_btn = True

# Main Execution
if fetch_btn or ("start" in query_params and "end" in query_params):
    if start_date > end_date:
        start_date, end_date = end_date, start_date

    total_days = (end_date - start_date).days + 1

    with st.spinner(f"Fetching articles from Indian Express for {total_days} day(s)..."):
        all_articles, date_status = fetch_all_dates(start_date, end_date, max_workers=6)

    if not all_articles:
        st.warning(f"No articles found for the period {start_date} to {end_date}.")
    else:
        df = pd.DataFrame(all_articles)
        total_articles = len(df)
        daily_avg = total_articles / total_days if total_days > 0 else 0
        top_section = df['section'].value_counts().index[0] if not df.empty else "N/A"
        top_section_count = df['section'].value_counts().iloc[0] if not df.empty else 0

        # KPI Metrics
        st.divider()
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total Articles", f"{total_articles:,}")
        m2.metric("Daily Average", f"{daily_avg:.1f} / day")
        m3.metric("Days Analyzed", f"{total_days} Days")
        m4.metric("Top Section", f"{top_section}", f"{top_section_count:,} articles")

        # Excel Download Button
        buf = io.BytesIO()
        export_to_excel(df, start_date, end_date, buf)
        excel_data = buf.getvalue()
        filename = f"IE_Article_Count_{start_date.strftime('%Y%m%d')}_to_{end_date.strftime('%Y%m%d')}.xlsx"

        st.download_button(
            label="📥 Download Complete Excel Report (.xlsx)",
            data=excel_data,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

        st.divider()

        # Tabs for details
        tab_sec, tab_daily, tab_monthly, tab_matrix, tab_raw = st.tabs([
            "📊 Section-Wise Count",
            "📅 Daily Total Count",
            "🗓️ Monthly Summary",
            "🔲 Daily Section Matrix",
            "📄 All Articles Data"
        ])

        # TAB 1: Section-Wise
        with tab_sec:
            col_t1, col_t2 = st.columns([1, 1])
            sec_counts = df['section'].value_counts().reset_index()
            sec_counts.columns = ['Section', 'Article Count']
            sec_counts['Share (%)'] = (sec_counts['Article Count'] / total_articles * 100).round(1)

            with col_t1:
                st.dataframe(
                    sec_counts,
                    use_container_width=True,
                    hide_index=True
                )
            with col_t2:
                st.bar_chart(
                    data=sec_counts.head(10).set_index('Section')['Article Count'],
                    color="#1F4E78"
                )

        # TAB 2: Daily Count
        with tab_daily:
            daily_groups = df.groupby(['date', 'day_of_week']).size().reset_index(name='Total Articles')
            daily_groups = daily_groups.sort_values('date')
            
            c_d1, c_d2 = st.columns([1, 1])
            with c_d1:
                st.dataframe(
                    daily_groups,
                    use_container_width=True,
                    hide_index=True
                )
            with c_d2:
                st.line_chart(
                    data=daily_groups.set_index('date')['Total Articles'],
                    color="#2F855A"
                )

        # TAB 3: Monthly Summary
        with tab_monthly:
            monthly_rows = []
            for (m_code, m_name), m_df in df.groupby(['month', 'month_name']):
                m_count = len(m_df)
                m_days = m_df['date'].nunique()
                monthly_rows.append({
                    "Month": m_name,
                    "Days Active": m_days,
                    "Total Articles": m_count,
                    "Daily Average": round(m_count / m_days, 1) if m_days > 0 else 0
                })
            st.dataframe(
                pd.DataFrame(monthly_rows),
                use_container_width=True,
                hide_index=True
            )

        # TAB 4: Daily Section Matrix
        with tab_matrix:
            pivot = pd.crosstab(index=df['date'], columns=df['section'], margins=True, margins_name="Grand Total")
            st.dataframe(
                pivot,
                use_container_width=True
            )

        # TAB 5: Raw Data
        with tab_raw:
            st.dataframe(
                df[['date', 'day_of_week', 'time', 'section', 'sub_section', 'title', 'url']],
                use_container_width=True,
                hide_index=True
            )
