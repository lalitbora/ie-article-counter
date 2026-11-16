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
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">📰 The Indian Express — Article Counter</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Daily Section-Wise, Daily Total & Monthly Summary Analyzer (Live Sitemaps)</div>', unsafe_allow_html=True)

today = datetime.now().date()

# Initialize session state for dates
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

# Quick select buttons
st.caption("Quick Select:")
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

    with st.spinner(f"Fetching articles from Indian Express sitemaps for {total_days} day(s)..."):
        all_articles, date_status = fetch_all_dates(start_date, end_date, max_workers=6)

    if not all_articles:
        st.error(f"Koi articles nahi mile ya fetch fail ho gaya ({start_date} to {end_date}).")
        
        # Show Detailed Error breakdown
        with st.expander("🔍 Click to view Connection / Date-wise Status", expanded=True):
            st.write("Har date ka network request status:")
            for d, status in date_status.items():
                if "Error" in str(status) or "Exception" in str(status):
                    st.error(f"**{d}**: {status}")
                else:
                    st.info(f"**{d}**: {status}")
            
            st.warning("""
            **Troubleshooting Note:**
            - Agar yahan `HTTP 403 Forbidden` dikh raha hai, to iska matlab Indian Express ka Cloudflare Firewall cloud server ke IP ko block kar raha hai.
            - Local PC par `start_web_app.bat` chalane se ye 100% smoothly kaam karta hai kyunki local internet IP block nahi hoti.
            """)
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

        # Detail Tabs
        tab_sec, tab_daily, tab_monthly, tab_matrix, tab_raw = st.tabs([
            "📊 Section-Wise Count",
            "📅 Daily Total Count",
            "🗓️ Monthly Summary",
            "🔲 Daily Section Matrix",
            "📄 All Articles Data"
        ])

        with tab_sec:
            col_t1, col_t2 = st.columns([1, 1])
            sec_counts = df['section'].value_counts().reset_index()
            sec_counts.columns = ['Section', 'Article Count']
            sec_counts['Share (%)'] = (sec_counts['Article Count'] / total_articles * 100).round(1)

            with col_t1:
                st.dataframe(sec_counts, use_container_width=True, hide_index=True)
            with col_t2:
                st.bar_chart(
                    data=sec_counts.head(10).set_index('Section')['Article Count'],
                    color="#1F4E78"
                )

        with tab_daily:
            daily_groups = df.groupby(['date', 'day_of_week']).size().reset_index(name='Total Articles')
            daily_groups = daily_groups.sort_values('date')
            
            c_d1, c_d2 = st.columns([1, 1])
            with c_d1:
                st.dataframe(daily_groups, use_container_width=True, hide_index=True)
            with c_d2:
                st.line_chart(
                    data=daily_groups.set_index('date')['Total Articles'],
                    color="#2F855A"
                )

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
            st.dataframe(pd.DataFrame(monthly_rows), use_container_width=True, hide_index=True)

        with tab_matrix:
            pivot = pd.crosstab(index=df['date'], columns=df['section'], margins=True, margins_name="Grand Total")
            st.dataframe(pivot, use_container_width=True)

        with tab_raw:
            st.dataframe(
                df[['date', 'day_of_week', 'time', 'section', 'sub_section', 'title', 'url']],
                use_container_width=True,
                hide_index=True
            )
