"""Interactive Streamlit Web Dashboard for PIB UPSC News Automation."""

import streamlit as st
from datetime import date, timedelta
from pathlib import Path
import os

from config import UPSC_PAPERS, OUTPUT_DIR
from pib_scraper import PIBScraper
from upsc_classifier import UPSCClassifier
from doc_builder import DocumentBuilder

# Page configuration
st.set_page_config(
    page_title="PIB UPSC Daily Intelligence",
    page_icon="🇮🇳",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Styling
st.markdown("""
<style>
    .main-title {
        font-size: 2.2rem;
        font-weight: 700;
        color: #1E3A8A;
        margin-bottom: 0.2rem;
    }
    .sub-title {
        color: #64748B;
        font-size: 1.05rem;
        margin-bottom: 1.5rem;
    }
    .metric-card {
        background-color: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 8px;
        padding: 15px;
        text-align: center;
    }
    .article-card {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-left: 5px solid #2563EB;
        border-radius: 6px;
        padding: 14px 18px;
        margin-bottom: 12px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    .tag-badge {
        display: inline-block;
        background-color: #EFF6FF;
        color: #1D4ED8;
        font-size: 0.8rem;
        font-weight: 600;
        padding: 3px 8px;
        border-radius: 4px;
        margin-right: 6px;
    }
    .ministry-label {
        font-size: 0.85rem;
        color: #64748B;
        font-style: italic;
    }
</style>
""", unsafe_allow_html=True)

# Header
st.markdown('<div class="main-title">🇮🇳 PIB UPSC Daily Intelligence & Headline Picker</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Automated extraction, syllabus alignment (GS 1-4), and daily Word & Excel generation from Press Information Bureau.</div>', unsafe_allow_html=True)

# Sidebar Controls
with st.sidebar:
    st.header("⚙️ Controls & Date")
    selected_date = st.date_input(
        "Select Date for PIB Releases",
        value=date.today(),
        max_value=date.today(),
        min_value=date.today() - timedelta(days=365)
    )
    
    fetch_btn = st.button("🚀 Fetch & Process PIB News", type="primary", use_container_width=True)
    
    st.divider()
    st.markdown("### 💡 Quick Tips")
    st.info(
        "• **Morning runs** before 11:00 AM may have few releases published yet.\n"
        "• **Evening runs** (8:30 PM+ IST) contain the complete day's bulletin.\n"
        "• You can also select **yesterday's date** to see a full day's releases."
    )

# Session state initialization
if "enriched_releases" not in st.session_state or fetch_btn:
    with st.spinner(f"Fetching releases from PIB for {selected_date.strftime('%d %B %Y')}..."):
        scraper = PIBScraper()
        releases, formatted_date = scraper.fetch_releases(target_date=selected_date)
        
        classifier = UPSCClassifier()
        enriched = classifier.process_all(releases)
        
        # Build files
        date_code = selected_date.strftime("%Y-%m-%d")
        excel_path = str(OUTPUT_DIR / f"PIB_UPSC_Daily_{date_code}.xlsx")
        docx_path = str(OUTPUT_DIR / f"PIB_UPSC_Daily_{date_code}.docx")
        
        DocumentBuilder.generate_excel(enriched, excel_path, formatted_date)
        DocumentBuilder.generate_docx(enriched, docx_path, formatted_date)
        
        st.session_state["enriched_releases"] = enriched
        st.session_state["formatted_date"] = formatted_date
        st.session_state["excel_path"] = excel_path
        st.session_state["docx_path"] = docx_path

enriched = st.session_state.get("enriched_releases", [])
formatted_date = st.session_state.get("formatted_date", selected_date.strftime("%d %B %Y"))
excel_path = st.session_state.get("excel_path", "")
docx_path = st.session_state.get("docx_path", "")

if not enriched:
    st.warning(f"No releases found for {formatted_date}.")
    st.stop()

# Metrics Row
relevant_items = [r for r in enriched if r.get("is_relevant")]
col1, col2, col3, col4, col5 = st.columns(5)

gs_counts = {"GS-1": 0, "GS-2": 0, "GS-3": 0, "GS-4": 0}
for r in relevant_items:
    p = r.get("paper", "")
    if p in gs_counts:
        gs_counts[p] += 1

with col1:
    st.metric("Total Scraped", len(enriched))
with col2:
    st.metric("UPSC Relevant", len(relevant_items), delta=f"{len(relevant_items)/max(1,len(enriched)):.0%}")
with col3:
    st.metric("GS-1 & GS-2", f"{gs_counts['GS-1']} / {gs_counts['GS-2']}")
with col4:
    st.metric("GS-3 (Eco/S&T/Env)", gs_counts['GS-3'])
with col5:
    st.metric("Routine Filtered", len(enriched) - len(relevant_items))

st.markdown("---")

# Download Action Bar
d_col1, d_col2, d_col3 = st.columns([2, 1, 1])
with d_col1:
    st.subheader(f"Curated Briefing ({formatted_date})")

with d_col2:
    if os.path.exists(docx_path):
        with open(docx_path, "rb") as f:
            st.download_button(
                label="📄 Download Word (.docx)",
                data=f.read(),
                file_name=Path(docx_path).name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )

with d_col3:
    if os.path.exists(excel_path):
        with open(excel_path, "rb") as f:
            st.download_button(
                label="📊 Download Excel (.xlsx)",
                data=f.read(),
                file_name=Path(excel_path).name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

# Tabbed view by GS Paper
tab_all, tab_gs1, tab_gs2, tab_gs3, tab_gs4, tab_audit = st.tabs([
    f"🌟 All Relevant ({len(relevant_items)})",
    f"GS-1 ({gs_counts['GS-1']})",
    f"GS-2 ({gs_counts['GS-2']})",
    f"GS-3 ({gs_counts['GS-3']})",
    f"GS-4 ({gs_counts['GS-4']})",
    f"📋 Full Audit Log ({len(enriched)})"
])

def render_article_card(item):
    paper = item.get("paper", "GS")
    tags = item.get("tags", [])
    badge_html = f'<span class="tag-badge">{paper}</span>' + "".join([f'<span class="tag-badge" style="background:#F1F5F9; color:#475569;">#{t}</span>' for t in tags[:3]])
    
    st.markdown(f"""
    <div class="article-card">
        <div style="margin-bottom: 6px;">{badge_html}</div>
        <div style="font-size: 1.1rem; font-weight: 600; color: #0F172A; margin-bottom: 4px;">
            {item['headline']}
        </div>
        <div class="ministry-label">🏛️ {item['ministry']}</div>
        <div style="margin-top: 8px; font-size: 0.92rem; color: #1E40AF;">
            🎯 <b>UPSC Alignment:</b> {item.get('relevance_reason', '')}
        </div>
        <div style="margin-top: 8px;">
            <a href="{item['url']}" target="_blank" style="color: #2563EB; font-size: 0.88rem; text-decoration: none; font-weight: 500;">
                🔗 Read Original PIB Release ↗
            </a>
        </div>
    </div>
    """, unsafe_allow_html=True)

with tab_all:
    if not relevant_items:
        st.info("No UPSC-relevant articles found for this date yet.")
    for item in relevant_items:
        render_article_card(item)

with tab_gs1:
    items_gs1 = [r for r in relevant_items if r.get("paper") == "GS-1"]
    if not items_gs1:
        st.info("No GS-1 releases for this date.")
    for item in items_gs1:
        render_article_card(item)

with tab_gs2:
    items_gs2 = [r for r in relevant_items if r.get("paper") == "GS-2"]
    if not items_gs2:
        st.info("No GS-2 releases for this date.")
    for item in items_gs2:
        render_article_card(item)

with tab_gs3:
    items_gs3 = [r for r in relevant_items if r.get("paper") == "GS-3"]
    if not items_gs3:
        st.info("No GS-3 releases for this date.")
    for item in items_gs3:
        render_article_card(item)

with tab_gs4:
    items_gs4 = [r for r in relevant_items if r.get("paper") == "GS-4"]
    if not items_gs4:
        st.info("No GS-4 releases for this date.")
    for item in items_gs4:
        render_article_card(item)

with tab_audit:
    st.caption("Complete chronological record of all releases published by PIB for this date.")
    audit_data = []
    for r in enriched:
        audit_data.append({
            "Status": "✅ Relevant" if r.get("is_relevant") else "❌ Routine",
            "Paper": r.get("paper", ""),
            "Ministry": r.get("ministry", ""),
            "Headline": r.get("headline", ""),
            "URL": r.get("url", "")
        })
    st.dataframe(audit_data, use_container_width=True)
