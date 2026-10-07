"""Interactive Streamlit Web Dashboard for PIB UPSC News Automation.
Enhanced with rich aesthetics, real-time filtering, visual analytics, and email dispatch.
"""

import streamlit as st
from datetime import date, timedelta, datetime
from pathlib import Path
import os
import io
import pandas as pd

import importlib
import config
import pib_scraper
import upsc_classifier
import doc_builder
import send_email

# Dynamic reloads for development agility
importlib.reload(config)
importlib.reload(pib_scraper)
importlib.reload(upsc_classifier)
importlib.reload(doc_builder)
importlib.reload(send_email)

from config import UPSC_PAPERS, OUTPUT_DIR
from pib_scraper import PIBScraper
from upsc_classifier import UPSCClassifier
from doc_builder import DocumentBuilder
from send_email import send_daily_brief_email, load_env_file

# Set Page Config
st.set_page_config(
    page_title="PIB UPSC Daily Intelligence Hub",
    page_icon="🇮🇳",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom High-Aesthetic Styling
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;600&display=swap');

    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif;
    }

    /* Top Hero Banner */
    .hero-container {
        background: linear-gradient(135deg, #0F172A 0%, #1E3A8A 50%, #172554 100%);
        border: 1px solid rgba(255, 255, 255, 0.1);
        border-radius: 16px;
        padding: 28px 32px;
        margin-bottom: 24px;
        color: white;
        box-shadow: 0 10px 30px -5px rgba(15, 23, 42, 0.3);
        position: relative;
        overflow: hidden;
    }
    
    .hero-container::before {
        content: '';
        position: absolute;
        top: -50%;
        right: -10%;
        width: 320px;
        height: 320px;
        background: radial-gradient(circle, rgba(59, 130, 246, 0.25) 0%, rgba(0,0,0,0) 70%);
        border-radius: 50%;
        pointer-events: none;
    }

    .hero-badge {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        background: rgba(255, 255, 255, 0.12);
        border: 1px solid rgba(255, 255, 255, 0.2);
        backdrop-filter: blur(8px);
        padding: 5px 12px;
        border-radius: 9999px;
        font-size: 0.78rem;
        font-weight: 600;
        letter-spacing: 0.05em;
        text-transform: uppercase;
        color: #93C5FD;
        margin-bottom: 12px;
    }

    .pulse-dot {
        width: 8px;
        height: 8px;
        border-radius: 50%;
        background-color: #10B981;
        box-shadow: 0 0 8px #10B981;
        display: inline-block;
    }

    .hero-title {
        font-size: 2.1rem;
        font-weight: 800;
        color: #FFFFFF;
        margin: 0 0 8px 0;
        letter-spacing: -0.02em;
        line-height: 1.2;
    }

    .hero-subtitle {
        font-size: 1.0rem;
        color: #CBD5E1;
        margin: 0;
        max-width: 820px;
        line-height: 1.5;
    }

    /* Metric Stat Card */
    .stat-card {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 12px;
        padding: 18px 20px;
        box-shadow: 0 2px 8px -2px rgba(15, 23, 42, 0.05);
        transition: transform 0.2s ease, box-shadow 0.2s ease;
    }
    .stat-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 8px 16px -4px rgba(15, 23, 42, 0.08);
    }
    .stat-label {
        font-size: 0.8rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.04em;
        color: #64748B;
        margin-bottom: 4px;
    }
    .stat-value {
        font-size: 1.85rem;
        font-weight: 800;
        color: #0F172A;
        line-height: 1.1;
    }
    .stat-sub {
        font-size: 0.82rem;
        font-weight: 500;
        color: #059669;
        margin-top: 4px;
    }

    /* Paper specific article cards */
    .article-box {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 12px;
        padding: 20px 22px;
        margin-bottom: 16px;
        box-shadow: 0 2px 6px rgba(15, 23, 42, 0.04);
        transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1);
        position: relative;
    }
    .article-box:hover {
        transform: translateY(-2px);
        box-shadow: 0 10px 20px -3px rgba(15, 23, 42, 0.08);
        border-color: #CBD5E1;
    }

    .box-gs1 { border-left: 6px solid #D97706; }
    .box-gs2 { border-left: 6px solid #4F46E5; }
    .box-gs3 { border-left: 6px solid #059669; }
    .box-gs4 { border-left: 6px solid #7C3AED; }

    .paper-pill {
        display: inline-block;
        padding: 4px 10px;
        border-radius: 6px;
        font-size: 0.76rem;
        font-weight: 700;
        letter-spacing: 0.03em;
        margin-right: 8px;
    }
    .pill-gs1 { background: #FEF3C7; color: #92400E; }
    .pill-gs2 { background: #EEF2FF; color: #3730A3; }
    .pill-gs3 { background: #ECFDF5; color: #065F46; }
    .pill-gs4 { background: #F5F3FF; color: #5B21B6; }

    .tag-chip {
        display: inline-block;
        background: #F1F5F9;
        color: #475569;
        font-size: 0.74rem;
        font-weight: 600;
        padding: 3px 8px;
        border-radius: 4px;
        margin-right: 6px;
        border: 1px solid #E2E8F0;
    }

    .article-headline {
        font-size: 1.12rem;
        font-weight: 700;
        color: #0F172A;
        line-height: 1.45;
        margin: 10px 0 6px 0;
    }

    .ministry-badge {
        font-size: 0.84rem;
        color: #64748B;
        font-weight: 500;
        display: inline-flex;
        align-items: center;
        gap: 4px;
    }

    .relevance-box {
        background: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 8px;
        padding: 10px 14px;
        margin-top: 12px;
        font-size: 0.9rem;
        color: #1E3A8A;
        line-height: 1.45;
    }

    .source-btn {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        margin-top: 12px;
        color: #2563EB;
        font-size: 0.88rem;
        font-weight: 600;
        text-decoration: none;
        transition: color 0.15s ease;
    }
    .source-btn:hover {
        color: #1D4ED8;
        text-decoration: underline;
    }

    /* Distribution Progress Bar */
    .dist-wrapper {
        background: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 12px;
        padding: 16px 20px;
        margin-bottom: 24px;
    }

    /* Action bar container */
    .action-bar {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 12px;
        padding: 16px 20px;
        margin-bottom: 20px;
        box-shadow: 0 2px 6px rgba(0,0,0,0.03);
    }
</style>
""", unsafe_allow_html=True)

# -------------------------------------------------------------
# HERO BANNER
# -------------------------------------------------------------
st.markdown("""
<div class="hero-container">
    <div class="hero-badge">
        <span class="pulse-dot"></span> Official PIB • UPSC CSE Syllabus Intelligence
    </div>
    <div class="hero-title">🇮🇳 PIB UPSC Daily Intelligence & Headline Curator</div>
    <p class="hero-subtitle">
        Automated scraping from Press Information Bureau, high-precision regex classification into <b>GS 1, 2, 3 & 4</b>, routine noise elimination, and instant Word & Excel reporting.
    </p>
</div>
""", unsafe_allow_html=True)

# -------------------------------------------------------------
# SIDEBAR CONTROLS
# -------------------------------------------------------------
with st.sidebar:
    st.markdown("### ⚙️ Intelligence Controls")
    
    # Date shortcut buttons
    today_val = date.today()
    c_btn1, c_btn2 = st.columns(2)
    with c_btn1:
        if st.button("📅 Today", use_container_width=True):
            st.session_state["target_date_picker"] = today_val
    with c_btn2:
        if st.button("⏪ Yesterday", use_container_width=True):
            st.session_state["target_date_picker"] = today_val - timedelta(days=1)
            
    default_date = st.session_state.get("target_date_picker", today_val)

    selected_date = st.date_input(
        "Select Release Date:",
        value=default_date,
        max_value=today_val,
        min_value=today_val - timedelta(days=365)
    )

    fetch_btn = st.button("🚀 Fetch & Classify Releases", type="primary", use_container_width=True)

    st.markdown("---")
    
    # Automation Health Status
    st.markdown("### 📡 Cloud & Local Automation")
    st.markdown("""
    <div style="background:#F8FAFC; border:1px solid #E2E8F0; border-radius:8px; padding:12px; font-size:0.83rem; line-height:1.5;">
        <div style="font-weight:700; color:#1E3A8A; margin-bottom:4px;">⏱️ Daily Automation Timetable:</div>
        <div>• <b>9:00 PM IST</b>: Full Comprehensive Daily Brief & Email Dispatch</div>
        <div style="margin-top:6px; color:#059669; font-weight:600;">✓ Active on GitHub Actions & Windows Task Scheduler</div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("### 💡 Quick Guidance")
    st.caption("• PIB releases are published gradually throughout the day. A morning run (before 11:00 AM) may only contain early releases, whereas the 9:00 PM run captures the complete daily bulletin.")

# -------------------------------------------------------------
# SCRAPING & CLASSIFICATION LOGIC
# -------------------------------------------------------------
should_fetch = fetch_btn or ("enriched_releases" not in st.session_state) or (st.session_state.get("current_date") != selected_date)

if should_fetch:
    with st.spinner(f"Connecting to PIB portal for {selected_date.strftime('%d %B %Y')}..."):
        scraper = PIBScraper()
        releases, formatted_date = scraper.fetch_releases(target_date=selected_date)
        
        # Smart Morning Fallback: if checking today and none published yet, load yesterday's complete wrap-up
        if not releases and selected_date == today_val:
            yesterday_date = today_val - timedelta(days=1)
            st.toast(f"ℹ️ PIB hasn't published today's releases yet. Loading yesterday's briefing ({yesterday_date.strftime('%d %B %Y')}).", icon="☕")
            releases, formatted_date = scraper.fetch_releases(target_date=yesterday_date)
            selected_date = yesterday_date

        classifier = UPSCClassifier()
        enriched = classifier.process_all(releases)
        
        # Build in-memory bytes for immediate, locked-file-free downloads
        docx_bytes = DocumentBuilder.generate_docx_bytes(enriched, formatted_date)
        excel_bytes = DocumentBuilder.generate_excel_bytes(enriched, formatted_date)

        # Save to disk
        date_code = selected_date.strftime("%Y-%m-%d")
        excel_path = str(OUTPUT_DIR / f"PIB_UPSC_Daily_{date_code}.xlsx")
        docx_path = str(OUTPUT_DIR / f"PIB_UPSC_Daily_{date_code}.docx")
        
        try:
            excel_path = DocumentBuilder.generate_excel(enriched, excel_path, formatted_date)
            docx_path = DocumentBuilder.generate_docx(enriched, docx_path, formatted_date)
        except Exception:
            pass

        st.session_state["enriched_releases"] = enriched
        st.session_state["formatted_date"] = formatted_date
        st.session_state["docx_bytes"] = docx_bytes
        st.session_state["excel_bytes"] = excel_bytes
        st.session_state["current_date"] = selected_date

enriched = st.session_state.get("enriched_releases", [])
formatted_date = st.session_state.get("formatted_date", selected_date.strftime("%d %B %Y"))
docx_bytes = st.session_state.get("docx_bytes", b"")
excel_bytes = st.session_state.get("excel_bytes", b"")
date_code = selected_date.strftime("%Y-%m-%d")

if not enriched:
    st.info(f"No releases found for {formatted_date}. Try selecting another date or yesterday.")
    st.stop()

# -------------------------------------------------------------
# METRICS & STAT CARDS
# -------------------------------------------------------------
relevant_items = [r for r in enriched if r.get("is_relevant")]
routine_count = len(enriched) - len(relevant_items)
unique_ministries = len(set(r.get("ministry", "") for r in enriched if r.get("ministry")))

gs_counts = {"GS-1": 0, "GS-2": 0, "GS-3": 0, "GS-4": 0}
for r in relevant_items:
    p = r.get("paper", "")
    if p in gs_counts:
        gs_counts[p] += 1

col1, col2, col3, col4 = st.columns(4)

with col1:
    st.markdown(f"""
    <div class="stat-card">
        <div class="stat-label">📥 Total Scraped</div>
        <div class="stat-value">{len(enriched)}</div>
        <div class="stat-sub" style="color:#64748B;">Releases on {formatted_date}</div>
    </div>
    """, unsafe_allow_html=True)

with col2:
    yield_rate = (len(relevant_items) / max(1, len(enriched))) * 100
    st.markdown(f"""
    <div class="stat-card" style="border-left: 4px solid #2563EB;">
        <div class="stat-label">🎯 UPSC High-Yield</div>
        <div class="stat-value" style="color:#1D4ED8;">{len(relevant_items)}</div>
        <div class="stat-sub" style="color:#1E40AF;">Yield Ratio: {yield_rate:.1f}%</div>
    </div>
    """, unsafe_allow_html=True)

with col3:
    st.markdown(f"""
    <div class="stat-card">
        <div class="stat-label">🏛️ Ministries Represented</div>
        <div class="stat-value">{unique_ministries}</div>
        <div class="stat-sub" style="color:#64748B;">Distinct Govt Bodies</div>
    </div>
    """, unsafe_allow_html=True)

with col4:
    filtered_pct = (routine_count / max(1, len(enriched))) * 100
    st.markdown(f"""
    <div class="stat-card" style="border-left: 4px solid #10B981;">
        <div class="stat-label">🛡️ Noise Excluded</div>
        <div class="stat-value" style="color:#059669;">{routine_count}</div>
        <div class="stat-sub" style="color:#047857;">{filtered_pct:.0f}% routine noise filtered</div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<div style='height: 16px;'></div>", unsafe_allow_html=True)

# -------------------------------------------------------------
# SYLLABUS DISTRIBUTION METER
# -------------------------------------------------------------
st.markdown(f"""
<div class="dist-wrapper">
    <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:10px;">
        <span style="font-weight:700; color:#0F172A; font-size:0.92rem;">📊 Daily Syllabus Weightage Distribution</span>
        <span style="font-size:0.8rem; color:#64748B;">Civil Services Mains GS 1-4</span>
    </div>
    <div style="display:grid; grid-template-columns: repeat(4, 1fr); gap:12px;">
        <div style="background:#FFFBEB; border:1px solid #FDE68A; padding:10px 14px; border-radius:8px;">
            <div style="font-size:0.75rem; font-weight:700; color:#B45309;">📚 GS-1 (Heritage & Soc)</div>
            <div style="font-size:1.3rem; font-weight:800; color:#92400E;">{gs_counts['GS-1']} <span style="font-size:0.75rem; font-weight:500;">articles</span></div>
        </div>
        <div style="background:#EEF2FF; border:1px solid #C7D2FE; padding:10px 14px; border-radius:8px;">
            <div style="font-size:0.75rem; font-weight:700; color:#4338CA;">⚖️ GS-2 (Polity, Gov & IR)</div>
            <div style="font-size:1.3rem; font-weight:800; color:#3730A3;">{gs_counts['GS-2']} <span style="font-size:0.75rem; font-weight:500;">articles</span></div>
        </div>
        <div style="background:#ECFDF5; border:1px solid #A7F3D0; padding:10px 14px; border-radius:8px;">
            <div style="font-size:0.75rem; font-weight:700; color:#047857;">🚀 GS-3 (Economy, S&T, Env)</div>
            <div style="font-size:1.3rem; font-weight:800; color:#065F46;">{gs_counts['GS-3']} <span style="font-size:0.75rem; font-weight:500;">articles</span></div>
        </div>
        <div style="background:#F5F3FF; border:1px solid #DDD6FE; padding:10px 14px; border-radius:8px;">
            <div style="font-size:0.75rem; font-weight:700; color:#6D28D9;">🧭 GS-4 (Ethics & Values)</div>
            <div style="font-size:1.3rem; font-weight:800; color:#5B21B6;">{gs_counts['GS-4']} <span style="font-size:0.75rem; font-weight:500;">articles</span></div>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# -------------------------------------------------------------
# DOWNLOAD & EMAIL DISPATCH ACTION BAR
# -------------------------------------------------------------
act_col1, act_col2, act_col3 = st.columns([2, 1, 1])

with act_col1:
    st.markdown(f"<h3 style='margin:0; font-size:1.35rem; color:#0F172A;'>Curated Intelligence Brief — {formatted_date}</h3>", unsafe_allow_html=True)
    st.caption("Download formatted reports or dispatch directly to candidate emails.")

with act_col2:
    if docx_bytes:
        st.download_button(
            label="📄 Download Word (.docx)",
            data=docx_bytes,
            file_name=f"PIB_UPSC_Daily_{date_code}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )

with act_col3:
    if excel_bytes:
        st.download_button(
            label="📊 Download Excel (.xlsx)",
            data=excel_bytes,
            file_name=f"PIB_UPSC_Daily_{date_code}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

# Direct Email Dispatch Expander
with st.expander("📧 Dispatch Brief via Email Directly from Dashboard", expanded=False):
    load_env_file()
    default_recipient = os.environ.get("MAIL_TO", "abhijeetsachan17@gmail.com")
    
    e_col1, e_col2 = st.columns([3, 1])
    with e_col1:
        recipient_input = st.text_input(
            "Recipient Email Address(es) (comma-separated):",
            value=default_recipient,
            help="You can enter one or multiple comma-separated emails."
        )
    with e_col2:
        st.markdown("<div style='height:28px;'></div>", unsafe_allow_html=True)
        send_click = st.button("🚀 Send Email Now", type="primary", use_container_width=True)
        
    if send_click:
        with st.spinner("Connecting to Gmail SMTP and sending brief..."):
            try:
                success = send_daily_brief_email(target_date=date_code, to_override=recipient_input)
                if success:
                    st.success(f"✓ Briefing email successfully delivered to: **{recipient_input}**!")
                else:
                    st.error("Could not deliver email. Please check your credentials in .env.")
            except Exception as e:
                st.error(f"Failed to send email: {e}")

st.markdown("---")

# -------------------------------------------------------------
# LIVE SEARCH & REAL-TIME FILTERS
# -------------------------------------------------------------
s_col1, s_col2 = st.columns([2, 1])

with s_col1:
    search_query = st.text_input(
        "🔍 Real-time Search:",
        placeholder="Search keywords, topics, or ministries (e.g. 'Semiconductor', 'Defense', 'ISRO')..."
    ).strip().lower()

with s_col2:
    all_ministries = sorted(list(set(r.get("ministry", "") for r in relevant_items if r.get("ministry"))))
    selected_ministry = st.selectbox(
        "Filter by Ministry:",
        options=["All Ministries"] + all_ministries
    )

# Filter items dynamically
def matches_filter(item):
    if selected_ministry != "All Ministries" and item.get("ministry") != selected_ministry:
        return False
    if search_query:
        query_in_headline = search_query in item.get("headline", "").lower()
        query_in_ministry = search_query in item.get("ministry", "").lower()
        query_in_reason = search_query in item.get("relevance_reason", "").lower()
        query_in_tags = any(search_query in t.lower() for t in item.get("tags", []))
        return query_in_headline or query_in_ministry or query_in_reason or query_in_tags
    return True

filtered_relevant = [r for r in relevant_items if matches_filter(r)]

# -------------------------------------------------------------
# TABBED CLASSIFICATION VIEW
# -------------------------------------------------------------
tab_all, tab_gs1, tab_gs2, tab_gs3, tab_gs4, tab_audit = st.tabs([
    f"🌟 All Relevant ({len(filtered_relevant)})",
    f"📚 GS-1 ({sum(1 for r in filtered_relevant if r.get('paper') == 'GS-1')})",
    f"⚖️ GS-2 ({sum(1 for r in filtered_relevant if r.get('paper') == 'GS-2')})",
    f"🚀 GS-3 ({sum(1 for r in filtered_relevant if r.get('paper') == 'GS-3')})",
    f"🧭 GS-4 ({sum(1 for r in filtered_relevant if r.get('paper') == 'GS-4')})",
    f"📋 Full Audit Log ({len(enriched)})"
])

def render_article_card(item):
    paper = item.get("paper", "GS")
    box_class = {
        "GS-1": "box-gs1",
        "GS-2": "box-gs2",
        "GS-3": "box-gs3",
        "GS-4": "box-gs4"
    }.get(paper, "box-gs2")

    pill_class = {
        "GS-1": "pill-gs1",
        "GS-2": "pill-gs2",
        "GS-3": "pill-gs3",
        "GS-4": "pill-gs4"
    }.get(paper, "pill-gs2")

    tags = item.get("tags", [])
    tag_html = "".join([f'<span class="tag-chip">#{t}</span>' for t in tags[:4]])

    st.markdown(f"""
    <div class="article-box {box_class}">
        <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:8px;">
            <div>
                <span class="paper-pill {pill_class}">{paper}</span>
                {tag_html}
            </div>
            <div class="ministry-badge">
                🏛️ {item.get('ministry', 'Government of India')}
            </div>
        </div>
        <div class="article-headline">{item.get('headline', '')}</div>
        <div class="relevance-box">
            🎯 <b>UPSC Alignment:</b> {item.get('relevance_reason', 'Relevant for CSE preparation.')}
        </div>
        <div style="margin-top:12px; display:flex; justify-content:space-between; align-items:center;">
            <a href="{item.get('url', '#')}" target="_blank" class="source-btn">
                🔗 Read Official PIB Release ↗
            </a>
            <span style="font-size:0.75rem; color:#94A3B8; font-family:'JetBrains Mono', monospace;">
                PRID: {item.get('prid', 'N/A')}
            </span>
        </div>
    </div>
    """, unsafe_allow_html=True)

with tab_all:
    if not filtered_relevant:
        st.info("No matching UPSC-relevant articles found with current filters.")
    for item in filtered_relevant:
        render_article_card(item)

with tab_gs1:
    items_gs1 = [r for r in filtered_relevant if r.get("paper") == "GS-1"]
    if not items_gs1:
        st.info("No GS-1 releases found for this filter.")
    for item in items_gs1:
        render_article_card(item)

with tab_gs2:
    items_gs2 = [r for r in filtered_relevant if r.get("paper") == "GS-2"]
    if not items_gs2:
        st.info("No GS-2 releases found for this filter.")
    for item in items_gs2:
        render_article_card(item)

with tab_gs3:
    items_gs3 = [r for r in filtered_relevant if r.get("paper") == "GS-3"]
    if not items_gs3:
        st.info("No GS-3 releases found for this filter.")
    for item in items_gs3:
        render_article_card(item)

with tab_gs4:
    items_gs4 = [r for r in filtered_relevant if r.get("paper") == "GS-4"]
    if not items_gs4:
        st.info("No GS-4 releases found for this filter.")
    for item in items_gs4:
        render_article_card(item)

with tab_audit:
    st.markdown("#### Complete Daily PIB Audit Log")
    st.caption("Chronological ledger of all releases published on this date with classification decisions.")
    
    audit_data = []
    for r in enriched:
        audit_data.append({
            "Classification": "🎯 Relevant" if r.get("is_relevant") else "🚫 Routine Filtered",
            "Paper": r.get("paper", "—"),
            "Ministry": r.get("ministry", "N/A"),
            "Headline": r.get("headline", ""),
            "Reason": r.get("relevance_reason", "Routine event / not syllabus-aligned"),
            "URL": r.get("url", "")
        })
        
    df_audit = pd.DataFrame(audit_data)
    
    # Filter search in table
    audit_search = st.text_input("Filter audit table:", placeholder="Type to filter audit log...")
    if audit_search:
        df_audit = df_audit[df_audit.apply(lambda row: audit_search.lower() in row.astype(str).str.lower().values, axis=1)]
        
    st.dataframe(
        df_audit,
        use_container_width=True,
        column_config={
            "URL": st.column_config.LinkColumn("PIB Link", display_text="Open Release ↗"),
            "Classification": st.column_config.TextColumn("Status", width="medium"),
            "Headline": st.column_config.TextColumn("Headline", width="large")
        },
        hide_index=True
    )
