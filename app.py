"""
Apples to Apples WebUI（Streamlit）
"""

import os
import sys
import tempfile

import pandas as pd
import streamlit as st

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import ata_bid
import ata_compare
import ata_delta
import ata_drive
import ata_extract
import ata_multi_compare
from config import (
    SPREADSHEET_ID, CONSTRUCTION_ITEM_FLAGS, CONSTRUCTION_ITEM_UNIT_PRICES,
    BID_RESULT_OPTIONS, OUTDOOR_UNIT_FLOOR_OPTIONS, CONSTRUCTION_AREA_OPTIONS,
    CONSTRUCTION_DAYS_OPTIONS,
)

SPREADSHEET_URL = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}"

st.set_page_config(
    page_title="Apple to Apple",
    page_icon="A",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ════════════════════════════════════════════════════════════
# CSS — Dark sidebar + Light main
# ════════════════════════════════════════════════════════════

CUSTOM_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=Noto+Sans+JP:wght@300;400;500;600;700&display=swap');

:root {
    --bg:          #F7F8FA;
    --bg-card:     #ffffff;
    --text:        #111827;
    --text-sub:    #6b7280;
    --text-muted:  #9ca3af;
    --accent:      #C0392B;
    --accent-hover:#A93226;
    --accent-light:#fef2f2;
    --red:         #C0392B;
    --red-bg:      #fef2f2;
    --green:       #10b981;
    --green-bg:    #ecfdf5;
    --amber:       #f59e0b;
    --amber-bg:    #fffbeb;
    --border:      #e5e7eb;
    --shadow:      0 1px 3px rgba(0,0,0,0.08);
    --shadow-md:   0 4px 12px rgba(0,0,0,0.06);
    --radius:      8px;
    --transition:  all 0.15s ease;
    --font:        ui-sans-serif,-apple-system,BlinkMacSystemFont,'Segoe UI',Helvetica,Arial,sans-serif;
    /* sidebar dark navy */
    --sb-bg:       #0F172A;
    --sb-surface:  #1E293B;
    --sb-border:   #334155;
    --sb-text:     #CBD5E1;
    --sb-text-m:   #94A3B8;
    --sb-active:   #1E293B;
}

/* ── Reset ── */
html,body,.stApp,[data-testid="stAppViewContainer"]{
    font-family:var(--font)!important;
    background:var(--bg)!important;
    color:var(--text)!important;
    -webkit-font-smoothing:antialiased;
}
.stApp>header,[data-testid="stHeader"]{display:none!important;}
[data-testid="stDecoration"]{display:none!important;}
#MainMenu,footer,[data-testid="stToolbar"]{display:none!important;}
.main .block-container{padding:2rem 3rem 4rem 3rem!important;max-width:1100px;}

/* ── サイドバー開閉ボタンを非表示（常にサイドバーを表示） ── */
[data-testid="collapsedControl"]{
    display:none!important;
}
[data-testid="stSidebarCollapseButton"]{
    display:none!important;
}

/* ── Sidebar ── */
section[data-testid="stSidebar"]{
    background:var(--sb-bg)!important;
    border-right:none!important;
    width:240px!important;
}
section[data-testid="stSidebar"]>div:first-child{padding:0 0.6rem!important;}
section[data-testid="stSidebar"] [data-testid="stSidebarContent"]{padding-top:0!important;}
section[data-testid="stSidebar"] [data-testid="stSidebarUserContent"]{padding-top:0!important;margin-top:0!important;}
section[data-testid="stSidebar"] .stMarkdown p,
section[data-testid="stSidebar"] .stMarkdown li,
section[data-testid="stSidebar"] .stMarkdown span{color:var(--sb-text)!important;font-size:0.84rem!important;}
section[data-testid="stSidebar"] a{color:var(--sb-text-m)!important;text-decoration:none!important;}
section[data-testid="stSidebar"] a:hover{color:var(--sb-text)!important;}
section[data-testid="stSidebar"] hr{border-color:var(--sb-border)!important;margin:0.3rem 0!important;}

/* Sidebar nav buttons */
/* nav-btn / nav-btn-active はボタンの上に重ねて表示するオーバーレイ用 */
.nav-btn, .nav-btn-active { display:none!important; }

/* サイドバー内のボタンをナビゲーションリンク風にスタイリング */
section[data-testid="stSidebar"] .stButton > button {
    background: transparent !important;
    color: var(--sb-text-m) !important;
    border: none !important;
    border-left: 3px solid transparent !important;
    border-radius: 0 6px 6px 0 !important;
    padding: 6px 12px !important;
    font-weight: 400 !important;
    font-size: 0.82rem !important;
    text-align: left !important;
    justify-content: flex-start !important;
    width: 100% !important;
    box-shadow: none !important;
    transition: var(--transition) !important;
    margin: 0px 0 !important;
    min-height: 32px !important;
    height: auto !important;
    line-height: 1.3 !important;
}
section[data-testid="stSidebar"] .stButton > button:hover {
    background: var(--sb-surface) !important;
    color: var(--sb-text) !important;
    border-left: 3px solid transparent !important;
}
section[data-testid="stSidebar"] .stButton > button p {
    color: inherit !important;
    text-align: left !important;
    font-size: 0.82rem !important;
    margin: 0 !important;
    line-height: 1.3 !important;
}

/* サイドバー内の「更新」ボタンは別クラスでスタイリング */
section[data-testid="stSidebar"] .stButton:last-of-type > button {
    background: var(--sb-surface) !important;
    color: var(--sb-text) !important;
    border: 1px solid var(--sb-border) !important;
    border-radius: 6px !important;
    font-size: 0.75rem !important;
    padding: 0.25rem 0.6rem !important;
    box-shadow: none !important;
    justify-content: center !important;
    text-align: center !important;
    border-left: 1px solid var(--sb-border) !important;
    min-height: 28px !important;
}
section[data-testid="stSidebar"] .stButton:last-of-type > button:hover {
    background: #334155 !important;
    color: #ffffff !important;
}

/* ── Page Header ── */
.page-header{margin-bottom:4px;}
.page-header-title{font-size:1.5rem;font-weight:700;color:var(--text);letter-spacing:-0.025em;line-height:1.3;}
.page-subtitle{color:var(--text-sub);font-size:0.88rem;margin-top:4px;margin-bottom:1.8rem;font-weight:400;line-height:1.5;}

/* ── Section ── */
.section-header{font-size:0.95rem;font-weight:600;color:var(--text);margin:2rem 0 0.75rem 0;display:block;}
.stMarkdown h3,.stSubheader{color:var(--text)!important;font-weight:600!important;font-size:0.95rem!important;border:none!important;padding-bottom:0!important;}

/* ── Card ── */
.card{background:var(--bg-card);border:1px solid var(--border);border-radius:var(--radius);padding:1.2rem;margin-bottom:0.75rem;box-shadow:var(--shadow);}

/* ── Buttons ── */
.stButton>button,.stDownloadButton>button{
    background:var(--accent)!important;color:#fff!important;border:none!important;
    border-radius:var(--radius)!important;font-weight:500!important;font-size:0.84rem!important;
    padding:0.5rem 1.2rem!important;transition:var(--transition)!important;box-shadow:none!important;
}
.stButton>button:hover,.stDownloadButton>button:hover{background:var(--accent-hover)!important;transform:none!important;}
.stButton>button:active{background:#922B21!important;}

.stLinkButton>a{
    background:transparent!important;color:var(--accent)!important;
    border:1px solid var(--border)!important;border-radius:var(--radius)!important;
    font-weight:500!important;font-size:0.84rem!important;padding:0.5rem 1.2rem!important;
    box-shadow:none!important;transition:var(--transition)!important;
}
.stLinkButton>a:hover{background:var(--accent-light)!important;}

/* ── Inputs — force white bg / dark text ── */
.stTextInput>div>div,.stNumberInput>div>div,.stSelectbox>div>div{
    background:#ffffff!important;border:1px solid var(--border)!important;
    border-radius:var(--radius)!important;color:var(--text)!important;box-shadow:none!important;
    transition:var(--transition)!important;
}
.stTextInput>div>div:focus-within,.stNumberInput>div>div:focus-within,.stSelectbox>div>div:focus-within{
    border-color:var(--accent)!important;box-shadow:0 0 0 3px rgba(192,57,43,0.12)!important;
}
.stTextInput label,.stNumberInput label,.stSelectbox label,.stFileUploader label,.stCheckbox label{
    color:var(--text-sub)!important;font-weight:500!important;font-size:0.8rem!important;
    text-transform:uppercase!important;letter-spacing:0.04em!important;
}
/* All input/textarea/select elements */
input,textarea,select,
[data-testid="stNumberInput"] input,
[data-testid="stTextInput"] input,
[data-baseweb="select"] div,
[data-baseweb="input"] input,
[data-baseweb="textarea"] textarea{
    background-color:#ffffff!important;color:#111827!important;
    caret-color:var(--accent)!important;font-family:var(--font)!important;
}
/* Selectbox dropdown list */
[data-baseweb="popover"] li,
[data-baseweb="menu"] li,
[role="option"]{background-color:#ffffff!important;color:#111827!important;}
[data-baseweb="popover"] li:hover,
[data-baseweb="menu"] li:hover,
[role="option"]:hover{background-color:#f3f4f6!important;}
/* Selectbox selected value text */
[data-baseweb="select"] span,
[data-baseweb="select"] [data-testid="stMarkdownContainer"]{color:#111827!important;}
.stCheckbox label span{color:var(--text-sub)!important;font-size:0.84rem!important;text-transform:none!important;letter-spacing:0!important;}

/* ── File Uploader ── */
[data-testid="stFileUploaderDropzoneInstructions"]{display:none!important;}
section[data-testid="stFileUploaderDropzone"]>button{display:none!important;}
[data-testid="stFileUploader"]{background:#f9fafb!important;border:2px dashed #d1d5db!important;border-radius:var(--radius)!important;padding:1.2rem!important;}
[data-testid="stFileUploaderDropzone"]{background-color:#f9fafb!important;border:2px dashed #d1d5db!important;}
[data-testid="stFileUploader"]:hover,[data-testid="stFileUploaderDropzone"]:hover{border-color:var(--accent)!important;background:#eef2ff!important;}
[data-testid="stFileUploader"] section{color:var(--text-sub)!important;}
[data-testid="stFileUploader"] small{color:var(--text-muted)!important;}
[data-testid="stFileUploader"] button{background:#ffffff!important;color:var(--text)!important;border:1px solid var(--border)!important;border-radius:6px!important;font-size:0.8rem!important;}

/* ── Metrics ── */
div[data-testid="stMetric"]{background:var(--bg-card)!important;border:1px solid var(--border)!important;border-radius:var(--radius)!important;padding:1rem 1.2rem!important;box-shadow:var(--shadow)!important;}
div[data-testid="stMetric"] label{color:var(--text-sub)!important;font-size:0.7rem!important;font-weight:500!important;text-transform:uppercase!important;letter-spacing:0.05em!important;}
div[data-testid="stMetric"] [data-testid="stMetricValue"]{color:var(--text)!important;font-size:1.4rem!important;font-weight:600!important;}

/* ── Dataframes ── */
[data-testid="stDataFrame"]{border-radius:var(--radius)!important;overflow:hidden!important;border:1px solid var(--border)!important;box-shadow:var(--shadow)!important;}

/* ── Alerts ── */
.stAlert{border-radius:var(--radius)!important;}
div[data-testid="stAlert"][data-type="info"]{background:var(--accent-light)!important;border:none!important;border-left:3px solid var(--accent)!important;}
div[data-testid="stAlert"][data-type="success"]{background:var(--green-bg)!important;border:none!important;border-left:3px solid var(--green)!important;}
div[data-testid="stAlert"][data-type="warning"]{background:var(--amber-bg)!important;border:none!important;border-left:3px solid var(--amber)!important;}
div[data-testid="stAlert"][data-type="error"]{background:var(--red-bg)!important;border:none!important;border-left:3px solid var(--red)!important;}

.stSpinner>div{color:var(--text-sub)!important;}
[data-testid="stStatusWidget"]{background:var(--bg-card)!important;border:1px solid var(--border)!important;border-radius:var(--radius)!important;}
.stCaption,.stMarkdown small{color:var(--text-muted)!important;}
hr{border-color:var(--border)!important;}
.stMarkdown p,.stMarkdown li,.stMarkdown span{color:var(--text)!important;}
.stMarkdown strong{color:var(--text)!important;font-weight:600!important;}

/* ── Dropdown ── */
[data-baseweb="popover"]{background:var(--bg-card)!important;border:1px solid var(--border)!important;border-radius:var(--radius)!important;box-shadow:var(--shadow-md)!important;}
[data-baseweb="popover"] li{color:var(--text)!important;font-size:0.84rem!important;}
[data-baseweb="popover"] li:hover{background:#f3f4f6!important;}

/* ── Chart ── */
[data-testid="stVegaLiteChart"]{background:var(--bg-card)!important;border:1px solid var(--border)!important;border-radius:var(--radius)!important;padding:0.5rem!important;box-shadow:var(--shadow)!important;}

/* ── Custom Metric Grid ── */
.metric-row{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:12px;margin:0.75rem 0;}
.metric-item{background:var(--bg-card);border:1px solid var(--border);border-radius:var(--radius);padding:1rem 1.2rem;box-shadow:var(--shadow);}
.metric-label{font-size:0.7rem;font-weight:500;color:var(--text-sub);text-transform:uppercase;letter-spacing:0.05em;margin-bottom:4px;}
.metric-value{font-size:1.4rem;font-weight:600;color:var(--text);letter-spacing:-0.02em;line-height:1.3;}
.metric-value-accent{font-size:1.4rem;font-weight:600;color:var(--accent);letter-spacing:-0.02em;line-height:1.3;}
.metric-sub{font-size:0.78rem;color:var(--text-sub);margin-top:4px;}

/* ── Info Bar ── */
.info-bar{display:flex;flex-wrap:wrap;gap:20px;padding:0.7rem 1rem;background:var(--bg-card);border:1px solid var(--border);border-radius:var(--radius);margin:0.75rem 0;box-shadow:var(--shadow);}
.info-bar-item{display:flex;align-items:center;gap:6px;font-size:0.84rem;}
.info-bar-label{color:var(--text-muted);font-weight:400;}
.info-bar-value{color:var(--text);font-weight:500;}

/* ── Tags ── */
.tag{display:inline-flex;align-items:center;padding:2px 8px;border-radius:10px;font-size:0.75rem;font-weight:500;line-height:1.5;}
.tag-accent{background:var(--accent-light);color:var(--accent);}
.tag-green{background:var(--green-bg);color:var(--green);}
.tag-red{background:var(--red-bg);color:var(--red);}
.tag-amber{background:var(--amber-bg);color:var(--amber);}
.tag-gray{background:#f3f4f6;color:var(--text-sub);}

/* ── Compare Cards ── */
.compare-grid{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin:0.75rem 0;}
.compare-card{background:var(--bg-card);border:1px solid var(--border);border-radius:var(--radius);padding:1rem 1.2rem;box-shadow:var(--shadow);}
.compare-card-label{font-size:0.68rem;font-weight:500;text-transform:uppercase;letter-spacing:0.06em;margin-bottom:6px;color:var(--text-muted);}
.compare-card-new .compare-card-label{color:var(--accent);}
.compare-card-name{font-size:0.92rem;font-weight:600;color:var(--text);margin-bottom:3px;}
.compare-card-detail{font-size:0.8rem;color:var(--text-sub);}

/* ── Total Card ── */
.total-card{background:var(--bg-card);border:1px solid var(--border);border-radius:var(--radius);padding:1.6rem 2rem;margin:1rem 0;text-align:center;box-shadow:var(--shadow);}
.total-card-label{font-size:0.75rem;font-weight:500;color:var(--text-sub);text-transform:uppercase;letter-spacing:0.05em;margin-bottom:6px;}
.total-card-value{font-size:2.2rem;font-weight:700;color:var(--text);letter-spacing:-0.03em;}
.total-card-note{font-size:0.78rem;color:var(--text-muted);margin-top:8px;}

/* ── Sidebar brand ── */
.sb-brand{padding:4px 12px 8px 12px;border-bottom:1px solid var(--sb-border);margin-bottom:4px;margin-top:0;}
.sb-brand-name{font-size:0.95rem;font-weight:700;color:#ffffff;letter-spacing:-0.01em;display:block;}
.sb-brand-sub{font-size:0.65rem;color:var(--sb-text-m);margin-top:1px;letter-spacing:0.04em;text-transform:uppercase;display:block;}

/* ── Sidebar status card ── */
.sb-status{background:var(--sb-surface);border:1px solid var(--sb-border);border-radius:6px;padding:7px 10px;margin:0 4px 4px 4px;}
.sb-status-title{font-size:0.65rem;font-weight:500;color:var(--sb-text-m);text-transform:uppercase;letter-spacing:0.06em;margin-bottom:5px;display:flex;align-items:center;gap:6px;}
.sb-status-dot{width:6px;height:6px;border-radius:50%;display:inline-block;}
.sb-status-dot-on{background:var(--green);}
.sb-status-dot-off{background:var(--sb-text-m);}
.sb-status-row{display:flex;justify-content:space-between;align-items:center;font-size:0.75rem;margin-bottom:2px;}
.sb-status-row:last-child{margin-bottom:0;}
.sb-status-label{color:var(--sb-text-m);}
.sb-status-val{color:var(--sb-text);font-weight:500;}
.sb-status-badge{display:inline-flex;padding:1px 7px;border-radius:10px;font-size:0.72rem;font-weight:500;}
.sb-status-badge-ok{background:rgba(16,185,129,0.15);color:var(--green);}
.sb-status-badge-warn{background:rgba(192,57,43,0.2);color:#F1948A;}

/* Sidebar link */
.sb-link{display:flex;align-items:center;gap:6px;padding:4px 10px;border-radius:6px;color:var(--sb-text-m)!important;text-decoration:none!important;font-size:0.75rem;font-weight:400;transition:var(--transition);}
.sb-link:hover{background:var(--sb-surface);color:var(--sb-text)!important;}
.sb-footer{font-size:0.65rem;color:var(--sb-text-m);padding:0.3rem 10px;margin-top:0.4rem;border-top:1px solid var(--sb-border);opacity:0.7;}

/* ── Step Guide ── */
.step-guide{display:flex;align-items:center;gap:6px;margin-bottom:1.5rem;padding:10px 14px;background:#ffffff;border:1px solid var(--border);border-radius:var(--radius);box-shadow:var(--shadow);}
.step-item{display:flex;align-items:center;gap:5px;font-size:0.8rem;color:var(--text-muted);font-weight:400;}
.step-item.active{color:var(--accent);font-weight:600;}
.step-num{width:20px;height:20px;border-radius:50%;display:inline-flex;align-items:center;justify-content:center;font-size:0.68rem;font-weight:600;background:#f3f4f6;color:var(--text-muted);flex-shrink:0;}
.step-item.active .step-num{background:var(--accent);color:#fff;}
.step-arrow{color:var(--border);font-size:0.75rem;}

/* ── Upload hint ── */
.upload-hint{text-align:center;padding:0.5rem;color:var(--text-sub);font-size:0.82rem;margin-top:4px;}
.upload-note{font-size:0.75rem;color:var(--text-muted);margin-top:2px;}

/* ── File count badge ── */
.file-count{display:inline-flex;align-items:center;gap:6px;padding:6px 14px;background:#ffffff;border:1px solid var(--border);border-radius:var(--radius);font-size:0.84rem;color:var(--text);font-weight:500;margin:8px 0;box-shadow:var(--shadow);}
.file-count-num{background:var(--accent);color:#fff;width:22px;height:22px;border-radius:50%;display:inline-flex;align-items:center;justify-content:center;font-size:0.72rem;font-weight:600;}

/* ── Empty State ── */
.empty-state{text-align:center;padding:3rem 2rem;border:2px dashed var(--border);border-radius:var(--radius);margin:1rem 0;}
.empty-state-icon{font-size:1.8rem;margin-bottom:0.6rem;color:var(--text-muted);}
.empty-state-text{color:var(--text-sub);font-size:0.88rem;}
.empty-state-hint{color:var(--text-muted);font-size:0.78rem;margin-top:0.3rem;}

/* Confidence highlight */
.cond-high{background:var(--bg-card);border:1px solid var(--border);border-radius:var(--radius);padding:0.5rem 0.75rem;margin-bottom:4px;}
.cond-medium{background:#FFF9C4;border:1px solid #FFF176;border-radius:var(--radius);padding:0.5rem 0.75rem;margin-bottom:4px;}
.cond-low{background:#FFE0B2;border:1px solid #FFB74D;border-radius:var(--radius);padding:0.5rem 0.75rem;margin-bottom:4px;}
.cond-label{font-size:0.72rem;font-weight:500;color:var(--text-sub);text-transform:uppercase;letter-spacing:0.04em;}
.cond-badge{font-size:0.65rem;font-weight:500;padding:1px 6px;border-radius:8px;margin-left:6px;}
.cond-badge-high{background:var(--green-bg);color:var(--green);}
.cond-badge-medium{background:#FFF9C4;color:#F57F17;}
.cond-badge-low{background:#FFE0B2;color:#E65100;}

/* Delta */
.delta-up{color:var(--red);font-weight:500;}
.delta-down{color:var(--accent);font-weight:500;}

/* Scrollbar */
::-webkit-scrollbar{width:6px;height:6px;}
::-webkit-scrollbar-track{background:transparent;}
::-webkit-scrollbar-thumb{background:#d1d5db;border-radius:3px;}
::-webkit-scrollbar-thumb:hover{background:#9ca3af;}
</style>
"""

st.markdown(CUSTOM_CSS, unsafe_allow_html=True)




# ── Helpers ──

def fmt_yen(val: int | float) -> str:
    return f"¥{int(val):,}"

def render_page_header(title: str, subtitle: str):
    st.markdown(f"""
    <div class="page-header">
        <div class="page-header-title">{title}</div>
    </div>
    <div class="page-subtitle">{subtitle}</div>
    """, unsafe_allow_html=True)

def render_section(title: str):
    st.markdown(f'<div class="section-header">{title}</div>', unsafe_allow_html=True)

def render_steps(steps: list[str], active: int = 0):
    items = []
    for i, s in enumerate(steps):
        cls = "step-item active" if i == active else "step-item"
        items.append(f'<span class="{cls}"><span class="step-num">{i+1}</span>{s}</span>')
        if i < len(steps) - 1:
            items.append('<span class="step-arrow">&rarr;</span>')
    st.markdown(f'<div class="step-guide">{"".join(items)}</div>', unsafe_allow_html=True)

def render_metric_card(label: str, value: str, accent: bool = False, sub: str = ""):
    cls = "metric-value-accent" if accent else "metric-value"
    sub_html = f'<div class="metric-sub">{sub}</div>' if sub else ""
    return f'<div class="metric-item"><div class="metric-label">{label}</div><div class="{cls}">{value}</div>{sub_html}</div>'

def render_metric_row(items: list[dict]):
    cards = "".join(render_metric_card(**item) for item in items)
    st.markdown(f'<div class="metric-row">{cards}</div>', unsafe_allow_html=True)

@st.cache_data(ttl=120)
def load_spreadsheet_data():
    try:
        sh = ata_compare.get_spreadsheet()
        return ata_compare.load_master(sh), ata_compare.load_trades(sh)
    except Exception:
        return [], []


# ── Sidebar ──
st.sidebar.markdown("""
<div class="sb-brand">
    <div class="sb-brand-name">Apple to Apple</div>
    <div class="sb-brand-sub">内装工事見積 比較ツール</div>
</div>
""", unsafe_allow_html=True)

# 未入力案件数を事前計算（サイドバーバッジ用）
@st.cache_data(ttl=120)
def _count_incomplete():
    try:
        sh = ata_compare.get_spreadsheet()
        ws = sh.worksheet("案件マスタ")
        data = ws.get_all_values()
        if len(data) < 2:
            return 0
        headers = data[0]
        check_cols = ["室外機設置階数", "施工エリア", "工事種別"]
        col_indices = [headers.index(c) for c in check_cols if c in headers]
        if not col_indices:
            return 0
        count = 0
        for row in data[1:]:
            if not row[0]:  # 案件IDが空ならスキップ
                continue
            for ci in col_indices:
                if ci < len(row) and (not row[ci] or row[ci].strip() == ""):
                    count += 1
                    break
        return count
    except Exception:
        return 0

_incomplete_count = _count_incomplete()
_incomplete_badge = f" ({_incomplete_count}件)" if _incomplete_count > 0 else ""

# ── ナビゲーション（ボタン方式）──
if "page" not in st.session_state:
    st.session_state.page = "PDF抽出"

nav_items = [
    ("PDF抽出",     "PDF 抽出"),
    ("類似案件検索", "類似案件検索"),
    ("増減表作成",   "増減表作成"),
    ("複数社比較",   "複数社比較"),
    ("入札分析",     "入札分析"),
    ("未入力補完",   f"未入力案件の補完{_incomplete_badge}"),
]

for key, label in nav_items:
    is_active = st.session_state.page == key
    active_class = "nav-btn-active" if is_active else "nav-btn"
    st.sidebar.markdown(f'<div class="{active_class}" id="nav-{key}">{label}</div>', unsafe_allow_html=True)
    if st.sidebar.button(label, key=f"nav_{key}", use_container_width=True):
        st.session_state.page = key
        st.rerun()

page = st.session_state.page

st.sidebar.markdown("---")

# DRIVE監視状態を session_state で直接管理
if "drive_last_check" not in st.session_state:
    st.session_state.drive_last_check = None
if "drive_pending" not in st.session_state:
    st.session_state.drive_pending = 0
if "drive_check_msg" not in st.session_state:
    st.session_state.drive_check_msg = None
if "drive_check_error" not in st.session_state:
    st.session_state.drive_check_error = None


def _format_last_check(iso_str):
    if not iso_str:
        return "未実行"
    try:
        from datetime import datetime, timezone
        dt = datetime.fromisoformat(iso_str)
        now = datetime.now(timezone.utc)
        delta = now - dt
        minutes = int(delta.total_seconds() / 60)
        if minutes < 1:
            return "たった今"
        if minutes < 60:
            return f"{minutes}分前"
        hours = minutes // 60
        if hours < 24:
            return f"{hours}時間前"
        days = hours // 24
        return f"{days}日前"
    except Exception:
        return "-"


last_check_ago = _format_last_check(st.session_state.drive_last_check)
pending = st.session_state.drive_pending
badge_cls = "sb-status-badge-warn" if pending > 0 else "sb-status-badge-ok"
badge_text = f"{pending} 件" if pending > 0 else "0"

st.sidebar.markdown(f"""
<div class="sb-status">
    <div class="sb-status-title"><span class="sb-status-dot sb-status-dot-on"></span>DRIVE 監視</div>
    <div class="sb-status-row">
        <span class="sb-status-label">最終チェック</span>
        <span class="sb-status-val">{last_check_ago}</span>
    </div>
    <div class="sb-status-row">
        <span class="sb-status-label">未処理</span>
        <span class="sb-status-badge {badge_cls}">{badge_text}</span>
    </div>
</div>
""", unsafe_allow_html=True)

# メッセージ表示（前回のチェック結果）
if st.session_state.drive_check_msg:
    st.sidebar.success(st.session_state.drive_check_msg)
    st.session_state.drive_check_msg = None
if st.session_state.drive_check_error:
    st.sidebar.error(st.session_state.drive_check_error)
    st.session_state.drive_check_error = None

if st.sidebar.button("更新", use_container_width=True):
    try:
        with st.sidebar:
            with st.spinner("Drive を確認中..."):
                result = ata_drive.check_inbox_once()
        from datetime import datetime, timezone
        st.session_state.drive_last_check = datetime.now(timezone.utc).isoformat()
        found = result.get("found", 0)
        processed = result.get("processed", 0)
        pending_new = result.get("pending", 0)
        pairs = result.get("pairs", 0)
        drawings = result.get("drawings", 0)
        estimates = result.get("estimates", 0)
        unknowns = result.get("unknowns", 0)
        errors = result.get("errors", [])
        st.session_state.drive_pending = pending_new
        if found == 0:
            st.session_state.drive_check_msg = "新規ファイルなし"
        elif processed > 0:
            st.session_state.drive_check_msg = f"検出: {found}件 / 処理: {processed}件 / 残: {pending_new}件"
        else:
            # 処理失敗の詳細を表示
            detail = f"図面: {drawings}件, 見積書: {estimates}件, 不明: {unknowns}件, ペア: {pairs}組"
            if errors:
                err_text = " / ".join(errors[:2])  # 最大、2件表示
                st.session_state.drive_check_error = f"処理失敗: {detail}\nエラー: {err_text}"
            elif pairs == 0:
                st.session_state.drive_check_error = f"ペアが見つかりませんでした。{detail}"
            else:
                st.session_state.drive_check_error = f"処理失敗。{detail}"
    except Exception as e:
        import traceback
        st.session_state.drive_check_error = f"エラー: {e}\n\n{traceback.format_exc()}"
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.markdown(f'<a class="sb-link" href="{SPREADSHEET_URL}" target="_blank">スプレッドシートを開く &#8599;</a>', unsafe_allow_html=True)
st.sidebar.markdown('<div class="sb-footer">GPT-4o + Google Sheets で構築</div>', unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════
# Page 1: PDF抽出
# ════════════════════════════════════════════════════════════

if page == "PDF抽出":
    render_page_header("PDF アップロード・データ抽出", "図面PDFと見積書PDFをアップロードして、坪数・金額・工種別金額と工事条件を抽出します。")

    current_step = 0
    if st.session_state.get("extract_done"):
        current_step = 1
    render_steps(["PDFをアップロード", "内容確認・条件修正", "確定して保存"], active=current_step)

    st.markdown("""
    <style>
    [data-testid="stFileUploader"] section {
        min-height: 300px !important;
        display: flex;
        align-items: center;
        justify-content: center;
    }
    [data-testid="stFileUploadDropzone"] {
        min-height: 280px !important;
    }
    </style>
    """, unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        drawing_file = st.file_uploader("図面PDF（平面図・設計図）", type=["pdf"], key="drawing")
    with col2:
        estimate_file = st.file_uploader("見積書PDF", type=["pdf"], key="estimate")

    # --- Step 1: 抽出実行 ---
    if drawing_file and estimate_file:
        if st.button("抽出を開始", type="primary", use_container_width=True):
            with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as f_d:
                f_d.write(drawing_file.read())
                drawing_path = f_d.name
            with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as f_e:
                f_e.write(estimate_file.read())
                estimate_path = f_e.name

            try:
                with st.status("処理中...", expanded=True) as status:
                    st.write("PDFからテキストを抽出中...")
                    drawing_text = ata_extract.extract_text_from_pdf(drawing_path, max_pages=5)
                    estimate_text = ata_extract.extract_text_from_pdf(estimate_path, max_pages=3)
                    st.write("GPT-4oでデータを解析中...")
                    from openai import OpenAI
                    client = OpenAI()
                    data = ata_extract.extract_with_gpt4o(client, drawing_text, estimate_text)
                    st.write("スプレッドシートに書き込み中...")
                    ata_extract.write_to_spreadsheet(data)
                    status.update(label="抽出完了 — 工事条件を確認してください", state="complete", expanded=True)

                # session_stateにデータを保存
                st.session_state["extract_data"] = data
                st.session_state["extract_done"] = True
            finally:
                os.unlink(drawing_path)
                os.unlink(estimate_path)

    # --- Step 2: 抽出結果表示 + 工事条件確認フォーム ---
    if st.session_state.get("extract_done"):
        data = st.session_state["extract_data"]
        tsubo = data["tsubo"]
        total = data["total_amount"]

        st.success("データ抽出・スプレッドシート書き込みが完了しました。以下の工事条件を確認・修正してください。")

        render_section("案件情報")
        render_metric_row([
            {"label": "案件名", "value": data["project_name"]},
            {"label": "坪数", "value": f"{tsubo} 坪", "accent": True},
            {"label": "合計金額（税抜）", "value": fmt_yen(total), "accent": True},
            {"label": "全体坪単価", "value": f"{fmt_yen(total / tsubo)}/坪"},
        ])

        st.markdown(f"""
        <div class="info-bar">
            <div class="info-bar-item"><span class="info-bar-label">ブランド</span><span class="tag tag-accent">{data['brand']}</span></div>
            <div class="info-bar-item"><span class="info-bar-label">業態</span><span class="tag tag-green">{data['category']}</span></div>
            <div class="info-bar-item"><span class="info-bar-label">所在地</span><span class="info-bar-value">{data['prefecture']}（{data['station']}）</span></div>
        </div>
        """, unsafe_allow_html=True)

        render_section("工種別内訳")
        trade_rows = []
        for t in data["trades"]:
            amt = t["amount"]
            trade_rows.append({"会社名": t["company"], "工種": t["trade"], "カテゴリ": t.get("mapped_category") or "-", "金額": amt, "坪単価": round(amt / tsubo)})
        df = pd.DataFrame(trade_rows)
        st.dataframe(df.style.format({"金額": "{:,.0f}", "坪単価": "{:,.0f}"}), use_container_width=True, hide_index=True)

        # --- 工事条件確認フォーム ---
        render_section("工事条件の確認・修正")

        conditions_raw = data.get("conditions", {})

        # confidence別の説明
        st.markdown("""
        <div style="display:flex;gap:12px;margin-bottom:12px;font-size:0.78rem;">
            <span><span class="cond-badge cond-badge-high">HIGH</span> PDFに明記あり</span>
            <span><span class="cond-badge cond-badge-medium">MEDIUM</span> 推測（要確認）</span>
            <span><span class="cond-badge cond-badge-low">LOW</span> 情報なし（要入力）</span>
        </div>
        """, unsafe_allow_html=True)

        edited_conditions = {}

        # 4カラムでフォーム表示
        cols_per_row = 4
        field_list = ata_extract.CONDITION_FIELDS
        for row_start in range(0, len(field_list), cols_per_row):
            cols = st.columns(cols_per_row)
            for i, col in enumerate(cols):
                idx = row_start + i
                if idx >= len(field_list):
                    break
                key, label, choices = field_list[idx]

                # config参照の選択肢を解決
                import config as _cfg
                _config_choices_map = {
                    "biz_category": _cfg.GYOTAI_OPTIONS,
                    "OUTDOOR_UNIT_FLOOR_OPTIONS": _cfg.OUTDOOR_UNIT_FLOOR_OPTIONS,
                    "CONSTRUCTION_DAYS_OPTIONS": _cfg.CONSTRUCTION_DAYS_OPTIONS,
                    "CONSTRUCTION_AREA_OPTIONS": _cfg.CONSTRUCTION_AREA_OPTIONS,
                }
                if isinstance(choices, str) and choices in _config_choices_map:
                    choices = _config_choices_map[choices]
                elif key == "biz_category":
                    choices = _cfg.GYOTAI_OPTIONS

                cond = conditions_raw.get(key, {})
                val = cond.get("value", "")
                conf = cond.get("confidence", "low")

                # confidence=low の場合は value を空に
                if conf == "low":
                    val = ""

                badge_cls = f"cond-badge-{conf}"
                bg_cls = f"cond-{conf}"

                with col:
                    st.markdown(f'<div class="{bg_cls}"><span class="cond-label">{label}</span><span class="cond-badge {badge_cls}">{conf.upper()}</span></div>', unsafe_allow_html=True)
                    if choices:
                        options = ["（未選択）"] + list(choices)
                        default_idx = 0
                        if val in choices:
                            default_idx = list(choices).index(val) + 1
                        selected = st.selectbox(label, options, index=default_idx, key=f"cond_{key}", label_visibility="collapsed")
                        edited_conditions[key] = "" if selected == "（未選択）" else selected
                    else:
                        edited_conditions[key] = st.text_input(label, value=val, key=f"cond_{key}", label_visibility="collapsed")

        # ── 入札結果入力 ──
        render_section("入札結果")
        bid_col1, bid_col2 = st.columns(2)
        with bid_col1:
            bid_result = st.selectbox("入札結果", ["（未選択）"] + BID_RESULT_OPTIONS, key="bid_result")
        with bid_col2:
            competitor_amount = 0
            if bid_result == "他社受注":
                competitor_amount = st.number_input("他社受注金額（税抜・円）", min_value=0, value=0, step=100000, key="competitor_amount")
            else:
                st.markdown('<div style="padding-top:1.6rem;font-size:0.82rem;color:var(--text-muted);">「他社受注」選択時のみ入力可</div>', unsafe_allow_html=True)

        # 確定ボタン
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("確定して保存", type="primary", use_container_width=True, key="save_conditions"):
            with st.spinner("工事条件をスプレッドシートに保存中..."):
                # 最新の案件IDを取得
                gc = ata_extract.get_gspread_client()
                sh = gc.open_by_key(ata_extract.SPREADSHEET_ID)
                ws = sh.worksheet(ata_extract.SHEET_MASTER)
                all_ids = [v for v in ws.col_values(1)[1:] if v]
                project_id = all_ids[-1] if all_ids else "P-001"

                ata_extract.write_conditions_to_spreadsheet(project_id, data["project_name"], edited_conditions)

                # 追加項目の保存（工事条件フォームの値を案件マスタに書き込み）
                _add_headers = ws.row_values(1)
                _add_fields = {
                    "室外機設置階数": edited_conditions.get("outdoor_unit_floor", ""),
                    "工期（日数）": edited_conditions.get("construction_days", ""),
                    "施工エリア": edited_conditions.get("construction_area", ""),
                    "その他備考": edited_conditions.get("remarks", ""),
                }
                _add_row_idx = len([v for v in ws.col_values(1) if v])  # データ行数（ヘッダー含む）
                for col_name, col_val in _add_fields.items():
                    if col_name in _add_headers and col_val:
                        _col_idx = _add_headers.index(col_name) + 1
                        ws.update_cell(_add_row_idx, _col_idx, col_val)

                # 入札結果の保存
                if bid_result and bid_result != "（未選択）":
                    ata_bid.write_bid_result(project_id, bid_result, competitor_amount)

            st.success(f"工事条件を「案件サマリー」シートに保存しました（{project_id}）。")
            st.link_button("スプレッドシートを確認", SPREADSHEET_URL)

            # 状態をリセット
            del st.session_state["extract_done"]
            del st.session_state["extract_data"]


# ════════════════════════════════════════════════════════════
# Page 2: 類似案件検索
# ════════════════════════════════════════════════════════════

elif page == "類似案件検索":
    render_page_header("類似案件検索・ベース金額算出", "坪数・ブランド・業態・工事条件を入力して、過去の類似案件からベース金額を算出します。")
    render_steps(["条件を入力", "類似案件を確認", "ベース金額を算出"], active=0)

    # 基本条件
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        new_tsubo = st.number_input("新規案件の坪数", min_value=1.0, max_value=500.0, value=20.0, step=0.5, help="例: 20.5")
    with col2:
        brand = st.text_input("ブランド名", placeholder="例: CRISP SALAD WORKS")
    with col3:
        from config import GYOTAI_OPTIONS
        category = st.selectbox("業態", [""] + GYOTAI_OPTIONS, key="search_cat", format_func=lambda x: x if x else "選択してください")
    with col4:
        kojishu = st.selectbox("工事種別", ["すべて", "新装のみ", "改装のみ"], key="search_kojishu")

    # 工事条件フィルター
    with st.expander("工事条件フィルター（任意）", expanded=False):
        st.caption("必須フィルター")
        fc1, fc2, fc3 = st.columns(3)
        with fc1:
            f_location = st.selectbox("立地タイプ", ["指定なし", "路面店", "ロードサイド", "商業施設", "その他"], key="f_loc")
        with fc2:
            f_state = st.selectbox("工事状態", ["指定なし", "スケルトン", "居抜き"], key="f_state")
        with fc3:
            from config import GYOTAI_OPTIONS as _GF
            f_biz = st.selectbox("業態（詳細）", ["指定なし"] + _GF, key="f_biz")

        st.caption("オプションフィルター")
        oc1, oc2, oc3, oc4 = st.columns(4)
        with oc1:
            f_hours = st.selectbox("工事時間帯", ["指定なし", "日中", "夜間", "混在"], key="f_hours")
        with oc2:
            f_floor = st.selectbox("フロア", ["指定なし", "地下", "1F", "2F以上"], key="f_floor")
        with oc3:
            f_kitchen = st.selectbox("厨房区画", ["指定なし", "新規", "既存", "なし"], key="f_kitchen")
        with oc4:
            f_designated = st.selectbox("指定業者", ["指定なし", "なし", "電気のみ", "空調のみ", "給排水のみ", "電気＋空調", "複数工種", "その他"], key="f_desig")

        st.caption("追加フィルター")
        ec1, ec2 = st.columns(2)
        with ec1:
            f_outdoor = st.selectbox("室外機設置階数", ["指定なし"] + OUTDOOR_UNIT_FLOOR_OPTIONS, key="f_outdoor")
        with ec2:
            f_area = st.selectbox("施工エリア", ["指定なし"] + CONSTRUCTION_AREA_OPTIONS, key="f_area")

    if st.button("検索する", type="primary", use_container_width=True):
        with st.spinner("スプレッドシートからデータを取得中..."):
            master, trades = load_spreadsheet_data()
        if not master:
            st.error("案件マスタにデータがありません。")
        else:
            # 工事条件でフィルタリング（案件サマリーシートがあれば）
            cond_data = {}
            try:
                sh_r = ata_compare.get_spreadsheet()
                ws_cond = sh_r.worksheet("案件サマリー")
                _cond_data = ws_cond.get_all_values()
                _cond_headers = _cond_data[0] if _cond_data else []
                cond_records = [dict(zip(_cond_headers, row)) for row in _cond_data[1:]] if len(_cond_data) > 1 else []
                cond_data = {r.get("案件ID", ""): r for r in cond_records}
            except Exception:
                pass

            # フィルターマッピング: (フィルター値, 案件サマリーの列名)
            active_filters = []
            for fval, col_name in [
                (f_location, "立地タイプ"), (f_state, "工事状態"), (f_biz, "業態"),
                (f_hours, "工事時間帯"), (f_floor, "フロア"),
                (f_kitchen, "厨房区画"), (f_designated, "指定業者範囲"),
            ]:
                if fval != "指定なし":
                    active_filters.append((fval, col_name))

            # 案件マスタ列に対するフィルター（室外機設置階数・施工エリア）
            master_filters = []
            if f_outdoor != "指定なし":
                master_filters.append((f_outdoor, "室外機設置階数"))
            if f_area != "指定なし":
                master_filters.append((f_area, "施工エリア"))

            # フィルター適用
            filtered_master = []
            for p in master:
                pid = p.get("案件ID", "")
                cond = cond_data.get(pid, {})
                match = True
                # 案件サマリーフィルター
                for fval, col_name in active_filters:
                    if cond.get(col_name, "") != fval:
                        match = False
                        break
                # 案件マスタフィルター
                if match:
                    for fval, col_name in master_filters:
                        if p.get(col_name, "") != fval:
                            match = False
                            break
                if match:
                    filtered_master.append(p)

            # ブランド名フィルター（入力がある場合、案件名に含まれる案件のみ）
            if brand and brand.strip():
                brand_filter = brand.strip().lower()
                filtered_master = [
                    p for p in filtered_master
                    if brand_filter in p.get("案件名", "").lower()
                    or brand_filter in p.get("ブランド名", "").lower()
                ]

            # 工事種別フィルター
            if kojishu == "新装のみ":
                filtered_master = [p for p in filtered_master if p.get("工事種別", "") == "新装"]
            elif kojishu == "改装のみ":
                filtered_master = [p for p in filtered_master if p.get("工事種別", "") == "改装"]

            filter_note = ""
            if (active_filters or master_filters or brand or kojishu != "すべて") and len(filtered_master) < len(master):
                filter_note = f"（フィルター適用: {len(master)} → {len(filtered_master)} 件）"

            filter_html = f'<div class="info-bar-item"><span class="info-bar-value" style="font-size:0.8rem">{filter_note}</span></div>' if filter_note else ""
            st.markdown(f'<div class="info-bar"><div class="info-bar-item"><span class="info-bar-label">取得件数</span><span class="tag tag-accent">{len(filtered_master)} 件</span></div>{filter_html}</div>', unsafe_allow_html=True)

            if not filtered_master:
                st.warning("条件に一致する案件が見つかりませんでした。条件を緩めて再検索してください。")

            # 坪数が近い順にソート
            def _tsubo_diff(p):
                try:
                    return abs(float(p.get("坪数", 0)) - new_tsubo)
                except (ValueError, TypeError):
                    return 999999
            sorted_master = sorted(filtered_master, key=_tsubo_diff)

            render_section("過去案件一覧（坪数が近い順）")
            table_rows = []
            for p in sorted_master:
                total_val = p.get("合計金額（税抜）", 0)
                try: total_val = int(total_val)
                except (ValueError, TypeError): total_val = 0
                tsubo_val = p.get("坪数", 0)
                try: tsubo_val = float(tsubo_val)
                except (ValueError, TypeError): tsubo_val = 0
                table_rows.append({
                    "案件ID": p.get("案件ID", ""),
                    "案件名": p.get("案件名", ""),
                    "業態": p.get("業態", ""),
                    "工事種別": p.get("工事種別", ""),
                    "坪数": tsubo_val,
                    "施工エリア": p.get("施工エリア", ""),
                    "室外機設置階数": p.get("室外機設置階数", ""),
                    "工期": p.get("工期（日数）", ""),
                    "合計金額": total_val,
                    "坪単価": round(total_val / tsubo_val) if tsubo_val else 0,
                    "その他備考": p.get("その他備考", ""),
                })
            df_projects = pd.DataFrame(table_rows)
            st.dataframe(df_projects.style.format({"坪数": "{:.1f}", "合計金額": "{:,.0f}", "坪単価": "{:,.0f}"}), use_container_width=True, hide_index=True)

            render_section("ベース金額算出")
            project_ids = [p.get("案件ID", "") for p in sorted_master]
            ref_id = st.selectbox("参照する案件を選択", project_ids, format_func=lambda pid: f"{pid} - {next((p.get('案件名','') for p in sorted_master if p.get('案件ID')==pid), '')}")

            if ref_id:
                ref_project = next((p for p in sorted_master if p.get("案件ID") == ref_id), None)
                if ref_project:
                    calc = ata_compare.calc_base_amount(ref_project, trades, new_tsubo)
                    if calc:
                        ref_tsubo = float(ref_project.get("坪数", 0))
                        st.markdown(f"""
                        <div class="compare-grid">
                            <div class="compare-card"><div class="compare-card-label">参照案件</div><div class="compare-card-name">[{ref_id}] {ref_project.get('案件名','')}</div><div class="compare-card-detail">{ref_tsubo} 坪</div></div>
                            <div class="compare-card compare-card-new"><div class="compare-card-label">新規案件</div><div class="compare-card-name">新規見積</div><div class="compare-card-detail">{new_tsubo} 坪</div></div>
                        </div>
                        """, unsafe_allow_html=True)

                        base_rows = []
                        for cat in ata_compare.TRADE_CATEGORIES:
                            v = calc.get(cat, {})
                            if v: base_rows.append({"工種": cat, "参照金額": v["ref_amount"], "坪単価": v["unit_price"], "ベース金額": v["base_amount"]})
                        v = calc["合計"]
                        base_rows.append({"工種": "合計", "参照金額": v["ref_amount"], "坪単価": v["unit_price"], "ベース金額": v["base_amount"]})
                        df_base = pd.DataFrame(base_rows)
                        st.dataframe(df_base.style.format({"参照金額": "{:,.0f}", "坪単価": "{:,.0f}", "ベース金額": "{:,.0f}"}), use_container_width=True, hide_index=True)

                        render_section("概算サマリー")
                        base_total = v["base_amount"]
                        tax = round(base_total * 0.1)
                        direction = round(base_total * 0.05)
                        render_metric_row([
                            {"label": "ベース金額（税抜）", "value": fmt_yen(base_total), "accent": True},
                            {"label": "消費税（10%）", "value": fmt_yen(tax)},
                            {"label": "ベース金額（税込）", "value": fmt_yen(base_total + tax)},
                            {"label": "ディレクション費（5%）", "value": fmt_yen(direction)},
                        ])
                        total_amount = base_total + tax + round(direction * 1.1)
                        st.markdown(f"""
                        <div class="total-card">
                            <div class="total-card-label">請求総額（税込）</div>
                            <div class="total-card-value">{fmt_yen(total_amount)}</div>
                            <div class="total-card-note">※ 増減項目は含まれていません。個別の追加・削除項目を考慮して最終金額を調整してください。</div>
                        </div>
                        """, unsafe_allow_html=True)

                        # ── 工事項目差分によるベース金額補正 ──
                        render_section("工事項目差分によるベース金額補正")
                        st.caption("新規案件に含まれる工事項目をチェックしてください。類似案件との差分から補正後ベース金額を算出します。")

                        # 新規案件の工事項目チェックボックス
                        new_item_flags = {}
                        flag_cols = st.columns(len(CONSTRUCTION_ITEM_FLAGS))
                        for i, (key, label) in enumerate(CONSTRUCTION_ITEM_FLAGS):
                            with flag_cols[i]:
                                new_item_flags[key] = st.checkbox(label, value=False, key=f"new_flag_{key}")

                        # 補正計算
                        corrections = ata_compare.calc_item_correction(
                            ref_project, new_item_flags, new_tsubo, trades
                        )

                        if corrections:
                            correction_total = 0
                            correction_html_items = ""
                            for c in corrections:
                                if c["direction"] == "subtract":
                                    sign = "ー"
                                    amt_display = f"-{fmt_yen(c['amount'])}"
                                    color = "var(--accent)"
                                    note = f"類似案件あり→新規なし"
                                    correction_total -= c["amount"]
                                else:
                                    sign = "＋"
                                    amt_display = f"+{fmt_yen(c['amount'])}"
                                    color = "var(--red)"
                                    note = f"類似案件なし→新規あり"
                                    correction_total += c["amount"]
                                correction_html_items += f'<div style="display:flex;justify-content:space-between;padding:4px 0;font-size:0.88rem;"><span>{sign} {c["label"]}（{note}）</span><span style="color:{color};font-weight:600;">{amt_display}</span></div>'

                            corrected_base = base_total + correction_total
                            corr_sign = "+" if correction_total >= 0 else ""

                            st.markdown(f"""
                            <div class="card">
                                <div style="display:flex;justify-content:space-between;padding:6px 0;border-bottom:1px solid var(--border);margin-bottom:8px;">
                                    <span style="font-weight:500;">類似案件ベース金額</span>
                                    <span style="font-weight:600;">{fmt_yen(base_total)}</span>
                                </div>
                                <div style="padding:4px 0;font-size:0.82rem;color:var(--text-sub);font-weight:500;">補正内容：</div>
                                {correction_html_items}
                                <div style="display:flex;justify-content:space-between;padding:8px 0 4px 0;border-top:1px solid var(--border);margin-top:8px;">
                                    <span style="font-weight:500;color:var(--text-sub);">補正合計</span>
                                    <span style="font-weight:600;">{corr_sign}{fmt_yen(correction_total)}</span>
                                </div>
                                <div style="display:flex;justify-content:space-between;padding:10px 0 4px 0;border-top:2px solid var(--accent);margin-top:4px;">
                                    <span style="font-weight:700;font-size:1.05rem;">補正後ベース金額 ★</span>
                                    <span style="font-weight:700;font-size:1.15rem;color:var(--accent);">{fmt_yen(corrected_base)}</span>
                                </div>
                            </div>
                            """, unsafe_allow_html=True)

                            # 補正後サマリー
                            corr_tax = round(corrected_base * 0.1)
                            corr_dir = round(corrected_base * 0.05)
                            corr_total = corrected_base + corr_tax + round(corr_dir * 1.1)
                            render_metric_row([
                                {"label": "補正後ベース金額（税抜）", "value": fmt_yen(corrected_base), "accent": True},
                                {"label": "消費税（10%）", "value": fmt_yen(corr_tax)},
                                {"label": "補正後（税込）", "value": fmt_yen(corrected_base + corr_tax)},
                                {"label": "補正後請求総額", "value": fmt_yen(corr_total), "accent": True},
                            ])
                        else:
                            st.info("類似案件と新規案件の工事項目に差分がありません。補正は不要です。")


# ════════════════════════════════════════════════════════════
# Page 3: 増減表作成
# ════════════════════════════════════════════════════════════

elif page == "増減表作成":
    render_page_header("増減表作成", "新規案件と参照案件を工種別に比較し、ベース金額との増減を可視化します。")
    render_steps(["案件を選択", "増減表を確認", "スプレッドシートに保存"], active=0)

    with st.spinner("データを取得中..."):
        master, trades = load_spreadsheet_data()

    if not master:
        st.error("案件マスタにデータがありません。先にPDF抽出を実行してください。")
    else:
        project_options = {p.get("案件ID", ""): f"{p.get('案件ID','')} - {p.get('案件名','')}（{p.get('坪数','')} 坪）" for p in master}
        col1, col2 = st.columns(2)
        with col1:
            new_id = st.selectbox("新規案件（比較したい案件を選択）", list(project_options.keys()), format_func=lambda x: project_options[x], key="delta_new")
        with col2:
            ref_id = st.selectbox("参照案件（ベースとなる過去案件を選択）", list(project_options.keys()), format_func=lambda x: project_options[x], key="delta_ref")

        if new_id == ref_id:
            st.warning("新規案件と参照案件は異なる案件を選択してください。同じ案件同士の比較はできません。")
        else:
            write_to_sheet = st.checkbox("スプレッドシートにも書き込む", value=True)
            if st.button("増減表を作成", type="primary", use_container_width=True):
                new_project = ata_delta.find_project(master, new_id)
                ref_project = ata_delta.find_project(master, ref_id)
                if not new_project or not ref_project:
                    st.error("案件が見つかりません。")
                else:
                    delta_rows = ata_delta.calc_delta(ref_project, new_project, trades)
                    new_tsubo = float(new_project.get("坪数", 0))
                    ref_tsubo = float(ref_project.get("坪数", 0))

                    render_section("比較対象")
                    st.markdown(f"""
                    <div class="compare-grid">
                        <div class="compare-card compare-card-new"><div class="compare-card-label">新規案件</div><div class="compare-card-name">[{new_id}] {new_project.get('案件名','')}</div><div class="compare-card-detail">{new_tsubo} 坪 | {fmt_yen(float(new_project.get('合計金額（税抜）',0)))}</div></div>
                        <div class="compare-card"><div class="compare-card-label">参照案件</div><div class="compare-card-name">[{ref_id}] {ref_project.get('案件名','')}</div><div class="compare-card-detail">{ref_tsubo} 坪 | {fmt_yen(float(ref_project.get('合計金額（税抜）',0)))}</div></div>
                    </div>
                    """, unsafe_allow_html=True)

                    render_section("工種別増減表")
                    display_rows = []
                    for r in delta_rows:
                        rate = r["rate"]
                        if abs(rate) <= 5: j = "妥当"
                        elif abs(rate) <= 15: j = "軽微増" if rate > 0 else "軽微減"
                        elif abs(rate) <= 30: j = "要確認" if rate > 0 else "大幅減"
                        else: j = "要精査" if rate > 0 else "大幅減"
                        display_rows.append({"工種": r["trade"], "参照坪単価": r["ref_unit_price"], "ベース金額": r["base_amount"], "新規金額": r["new_amount"], "新規坪単価": r["new_unit_price"], "増減額": r["delta"], "増減率(%)": r["rate"], "判定": j})

                    df_delta = pd.DataFrame(display_rows)
                    def hl(row):
                        if row["工種"] == "合計": return ["font-weight:600;background:#f3f4f6"] * len(row)
                        if row["増減額"] > 0: return ["background:#fef2f2"] * len(row)
                        if row["増減額"] < 0: return ["background:#eef2ff"] * len(row)
                        return [""] * len(row)
                    st.dataframe(df_delta.style.format({"参照坪単価":"{:,.0f}","ベース金額":"{:,.0f}","新規金額":"{:,.0f}","新規坪単価":"{:,.0f}","増減額":"{:+,.0f}","増減率(%)":"{:+.1f}%"}).apply(hl, axis=1), use_container_width=True, hide_index=True)

                    render_section("サマリー")
                    tr = delta_rows[-1]
                    ds = "+" if tr["delta"] > 0 else ""
                    dc = "delta-up" if tr["delta"] > 0 else "delta-down"
                    render_metric_row([
                        {"label": "ベース金額", "value": fmt_yen(tr["base_amount"])},
                        {"label": "新規金額", "value": fmt_yen(tr["new_amount"]), "accent": True},
                        {"label": "増減額", "value": f"{ds}{fmt_yen(tr['delta'])}", "sub": f'<span class="{dc}">{ds}{tr["rate"]:.1f}%</span>'},
                        {"label": "新規坪単価", "value": f"{fmt_yen(tr['new_unit_price'])}/坪"},
                    ])

                    render_section("工種別増減グラフ")
                    chart_data = df_delta[df_delta["工種"] != "合計"][["工種", "増減額"]].set_index("工種")
                    st.bar_chart(chart_data, color=["#C0392B"])

                    if write_to_sheet:
                        with st.spinner("スプレッドシートに書き込み中..."):
                            sh_write = ata_delta.get_spreadsheet(readonly=False)
                            ata_delta.write_delta_sheet(sh_write, ref_project, new_project, delta_rows)
                        st.success("スプレッドシートに増減表を書き込みました。")
                        st.link_button("スプレッドシートを確認", SPREADSHEET_URL)


# ════════════════════════════════════════════════════════════
# Page 4: 複数社比較
# ════════════════════════════════════════════════════════════

elif page == "複数社比較":
    render_page_header("複数社見積 横並び比較", "同一案件に対する複数社の見積書PDFをアップロードし、工種別に横並びで比較します。")
    render_steps(["案件情報を入力", "見積書PDFをアップロード", "比較結果を確認"], active=0)

    col1, col2 = st.columns(2)
    with col1:
        mc_project_name = st.text_input("案件名", placeholder="例: CRISP SALAD WORKS 新日本橋 新装工事", key="mc_project")
    with col2:
        mc_tsubo = st.number_input("坪数", min_value=1.0, max_value=500.0, value=20.0, step=0.5, key="mc_tsubo", help="例: 20.5")

    st.markdown('<div style="font-size:0.84rem;color:#6b7280;margin-bottom:4px;">2社以上の見積書PDFをアップロードしてください（複数選択可）</div>', unsafe_allow_html=True)
    uploaded_files = st.file_uploader("見積書PDF（複数選択可）", type=["pdf"], accept_multiple_files=True, key="mc_pdfs")

    if uploaded_files:
        n = len(uploaded_files)
        st.markdown(f'<div class="file-count"><span class="file-count-num">{n}</span>{n}社分のPDFを選択中</div>', unsafe_allow_html=True)

    if uploaded_files and len(uploaded_files) >= 2 and mc_project_name:
        write_to_sheet = st.checkbox("スプレッドシートの「比較表」にも書き込む", value=True, key="mc_write")
        btn_label = f"{len(uploaded_files)}社を比較する"
        if st.button(btn_label, type="primary", use_container_width=True, key="mc_run"):
            from openai import OpenAI
            openai_client = OpenAI()
            estimates, temp_paths = [], []
            with st.status("見積書を解析中...", expanded=True) as status:
                for i, uf in enumerate(uploaded_files):
                    st.write(f"({i+1}/{len(uploaded_files)}) {uf.name} を解析中...")
                    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
                        tmp.write(uf.read())
                        temp_paths.append(tmp.name)
                    text = ata_multi_compare.extract_text_from_pdf(tmp.name, max_pages=5)
                    est = ata_multi_compare.extract_estimate_with_gpt(openai_client, text, uf.name)
                    estimates.append(est)
                    st.write(f"  {est['company']}: {fmt_yen(est['total_amount'])}")
                status.update(label=f"{len(estimates)} 社の解析完了", state="complete")
            for p in temp_paths: os.unlink(p)

            comparison = ata_multi_compare.build_comparison(estimates)
            companies = comparison["companies"]
            trade_data = comparison["trade_data"]
            totals = comparison["totals"]
            cheapest = comparison["cheapest"]

            render_section("比較対象")
            st.markdown(f"""
            <div class="info-bar">
                <div class="info-bar-item"><span class="info-bar-label">案件名</span><span class="info-bar-value">{mc_project_name}</span></div>
                <div class="info-bar-item"><span class="info-bar-label">坪数</span><span class="tag tag-accent">{mc_tsubo} 坪</span></div>
                <div class="info-bar-item"><span class="info-bar-label">比較社数</span><span class="tag tag-green">{len(companies)} 社</span></div>
            </div>
            """, unsafe_allow_html=True)

            render_section("工種別比較表")
            compare_rows = []
            for cat in ata_multi_compare.TRADE_CATEGORIES:
                row = {"工種": cat}
                for c in companies: row[c] = trade_data[cat].get(c, 0)
                compare_rows.append(row)
            tr_data = {"工種": "合計"}
            for c in companies: tr_data[c] = totals.get(c, 0)
            compare_rows.append(tr_data)
            ur_data = {"工種": "坪単価"}
            for c in companies: ur_data[c] = round(totals.get(c, 0) / mc_tsubo) if mc_tsubo else 0
            compare_rows.append(ur_data)
            df_compare = pd.DataFrame(compare_rows)

            def hl_cheap(row):
                styles = [""] * len(row)
                if row["工種"] in ("合計", "坪単価"):
                    styles = ["font-weight:600;background:#f3f4f6"] * len(row)
                    if row["工種"] == "合計":
                        amts = {c: row[c] for c in companies if row[c] > 0}
                        if amts:
                            mc = min(amts, key=amts.get)
                            styles[list(row.index).index(mc)] = "font-weight:700;background:#ecfdf5;color:#065f46"
                    return styles
                amts = {c: row[c] for c in companies if row[c] > 0}
                if amts:
                    mc = min(amts, key=amts.get)
                    styles[list(row.index).index(mc)] = "font-weight:600;background:#ecfdf5;color:#065f46"
                return styles

            st.dataframe(df_compare.style.format({c: "{:,.0f}" for c in companies}).apply(hl_cheap, axis=1), use_container_width=True, hide_index=True)

            render_section("最安値サマリー")
            summary_rows = []
            for cat in ata_multi_compare.TRADE_CATEGORIES:
                ch = cheapest.get(cat)
                if ch: summary_rows.append({"工種": cat, "最安値業者": ch, "金額": trade_data[cat].get(ch, 0), "坪単価": round(trade_data[cat].get(ch, 0) / mc_tsubo) if mc_tsubo else 0})
            ch_total = cheapest.get("合計")
            if ch_total: summary_rows.append({"工種": "合計", "最安値業者": ch_total, "金額": totals.get(ch_total, 0), "坪単価": round(totals.get(ch_total, 0) / mc_tsubo) if mc_tsubo else 0})
            st.dataframe(pd.DataFrame(summary_rows).style.format({"金額": "{:,.0f}", "坪単価": "{:,.0f}"}), use_container_width=True, hide_index=True)

            render_section("各社合計金額比較")
            st.bar_chart(pd.DataFrame({"業者": companies, "合計金額": [totals.get(c, 0) for c in companies]}).set_index("業者"), color=["#C0392B"])

            if write_to_sheet:
                with st.spinner("スプレッドシートに書き込み中..."):
                    ata_multi_compare.write_comparison_sheet(mc_project_name, mc_tsubo, comparison)
                st.success("スプレッドシートの「比較表」シートに書き込みました。")
                st.link_button("スプレッドシートを確認", SPREADSHEET_URL)

    elif uploaded_files and len(uploaded_files) < 2:
        st.warning("比較には最低2社分の見積書PDFが必要です。")
    elif not mc_project_name and uploaded_files:
        st.warning("案件名を入力してください。")
    else:
        st.markdown("""
        <div class="empty-state">
            <div class="empty-state-icon">&#8860;</div>
            <div class="empty-state-text">案件名を入力し、2社以上の見積書PDFをアップロードしてください</div>
            <div class="empty-state-hint">各PDFからGPT-4oが自動で業者名・工種別金額を抽出します</div>
        </div>
        """, unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════
# Page 5: 入札分析
# ════════════════════════════════════════════════════════════

elif page == "入札分析":
    render_page_header("入札結果分析", "入札結果データを分析し、受注できた坪単価の傾向を可視化します。")

    # フィルター
    render_section("フィルター")
    fc1, fc2, fc3 = st.columns(3)
    with fc1:
        bid_brand = st.text_input("ブランド名", placeholder="例: CRISP SALAD WORKS", key="bid_brand")
    with fc2:
        from config import GYOTAI_OPTIONS as _BG
        bid_category = st.selectbox("業態", [""] + _BG, key="bid_cat", format_func=lambda x: x if x else "すべて")
    with fc3:
        bid_koji = st.selectbox("工事種別", ["すべて", "新装のみ", "改装のみ"], key="bid_koji")

    with st.spinner("データを取得中..."):
        master, _ = load_spreadsheet_data()

    bid_data = ata_bid.load_bid_data(master)
    if not bid_data:
        st.markdown("""
        <div class="empty-state">
            <div class="empty-state-icon">&#9878;</div>
            <div class="empty-state-text">入札結果データがありません</div>
            <div class="empty-state-hint">PDF抽出画面で入札結果を入力するか、スプレッドシートの「入札結果」列を直接入力してください</div>
        </div>
        """, unsafe_allow_html=True)
    else:
        filtered_bid = ata_bid.filter_bid_data(bid_data, bid_brand, bid_category, bid_koji)

        if not filtered_bid:
            st.warning("条件に一致する入札データがありません。フィルターを緩めてください。")
        else:
            st.markdown(f'<div class="info-bar"><div class="info-bar-item"><span class="info-bar-label">対象件数</span><span class="tag tag-accent">{len(filtered_bid)} 件</span></div></div>', unsafe_allow_html=True)

            # 集計サマリー
            render_section("集計サマリー")
            summary = ata_bid.calc_bid_summary(filtered_bid)

            render_metric_row([
                {"label": "自社受注案件の平均坪単価", "value": fmt_yen(summary["自社受注平均坪単価"]) + "/坪" if summary["自社受注平均坪単価"] > 0 else "N/A", "accent": True, "sub": f'{summary["自社受注件数"]} 件'},
                {"label": "他社受注案件の平均受注坪単価", "value": fmt_yen(summary["他社受注平均坪単価"]) + "/坪" if summary["他社受注平均坪単価"] > 0 else "N/A", "sub": f'{summary["他社受注件数"]} 件'},
                {"label": "差額（自社が高かった平均）", "value": f"+{fmt_yen(summary['差額'])}/坪" if summary["差額"] > 0 else (fmt_yen(summary["差額"]) + "/坪" if summary["差額"] != 0 else "N/A"), "sub": "敗退時の自社提出坪単価 − 他社受注坪単価"},
            ])

            # 案件一覧テーブル
            render_section("案件一覧")
            table_rows = []
            for r in filtered_bid:
                table_rows.append({
                    "案件名": r["案件名"],
                    "坪数": r["坪数"],
                    "自社提出坪単価": r["自社提出坪単価"],
                    "他社受注坪単価": r["他社受注坪単価"] if r["他社受注坪単価"] > 0 else None,
                    "差額": r["差額"] if r["差額"] != 0 else None,
                    "入札結果": r["入札結果"],
                })
            df_bid = pd.DataFrame(table_rows)

            def hl_bid(row):
                if row["入札結果"] == "自社受注":
                    return ["background:#eef2ff"] * len(row)
                elif row["入札結果"] == "他社受注":
                    return ["background:#fef2f2"] * len(row)
                return [""] * len(row)

            fmt_dict = {"坪数": "{:.1f}", "自社提出坪単価": "{:,.0f}", "他社受注坪単価": "{:,.0f}", "差額": "{:+,.0f}"}
            st.dataframe(df_bid.style.format(fmt_dict, na_rep="-").apply(hl_bid, axis=1), use_container_width=True, hide_index=True)

            # 散布図
            render_section("坪数 × 坪単価 散布図")
            st.caption("X軸=坪数、Y軸=坪単価、色=入札結果（青=自社受注、赤=他社受注）")

            import altair as alt

            # 散布図用データ構築
            scatter_rows = []
            for r in filtered_bid:
                if r["自社提出坪単価"] > 0:
                    scatter_rows.append({
                        "坪数": r["坪数"],
                        "坪単価": r["自社提出坪単価"],
                        "種別": "自社提出",
                        "入札結果": r["入札結果"],
                        "案件名": r["案件名"],
                    })
                if r["他社受注坪単価"] > 0:
                    scatter_rows.append({
                        "坪数": r["坪数"],
                        "坪単価": r["他社受注坪単価"],
                        "種別": "他社受注",
                        "入札結果": r["入札結果"],
                        "案件名": r["案件名"],
                    })

            if scatter_rows:
                df_scatter = pd.DataFrame(scatter_rows)

                color_scale = alt.Scale(
                    domain=["自社受注", "他社受注", "随意契約"],
                    range=["#C0392B", "#E74C3C", "#f59e0b"]
                )

                chart = alt.Chart(df_scatter).mark_circle(size=80, opacity=0.7).encode(
                    x=alt.X("坪数:Q", title="坪数", scale=alt.Scale(zero=False)),
                    y=alt.Y("坪単価:Q", title="坪単価（円/坪）", scale=alt.Scale(zero=False)),
                    color=alt.Color("入札結果:N", scale=color_scale, title="入札結果"),
                    shape=alt.Shape("種別:N", title="種別"),
                    tooltip=["案件名", "坪数", "坪単価", "入札結果", "種別"],
                ).properties(
                    width="container",
                    height=400,
                ).interactive()

                st.altair_chart(chart, use_container_width=True)
            else:
                st.info("散布図に表示できるデータがありません。")


# ════════════════════════════════════════════════════════════
# Page 6: 未入力案件の補完
# ════════════════════════════════════════════════════════════

elif page == "未入力補完":
    render_page_header("未入力案件の補完", "自動読み込みされた案件の未入力項目を補完してください。")

    with st.spinner("データを取得中..."):
        master, _ = load_spreadsheet_data()

    # 未入力案件の判定
    CHECK_FIELDS = ["室外機設置階数", "施工エリア", "工事種別"]
    incomplete = []
    for p in master:
        pid = p.get("案件ID", "")
        if not pid:
            continue
        missing = []
        for f in CHECK_FIELDS:
            val = str(p.get(f, "")).strip()
            if not val:
                missing.append(f)
        if missing:
            incomplete.append((p, missing))

    if not incomplete:
        st.markdown("""
        <div class="empty-state">
            <div class="empty-state-icon">&#10003;</div>
            <div class="empty-state-text">未入力案件はありません</div>
            <div class="empty-state-hint">すべての案件の必須項目が入力済みです</div>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.markdown(f'<div class="info-bar"><div class="info-bar-item"><span class="info-bar-label">未入力案件</span><span class="tag tag-red">{len(incomplete)} 件</span></div></div>', unsafe_allow_html=True)

        # 各案件の入力フォーム
        all_inputs = {}

        for p, missing_fields in incomplete:
            pid = p.get("案件ID", "")
            pname = p.get("案件名", "")
            tsubo_val = p.get("坪数", "")
            category_val = p.get("業態", "")
            total_val = p.get("合計金額（税抜）", 0)
            try:
                total_val = int(float(total_val))
            except (ValueError, TypeError):
                total_val = 0

            missing_labels = ", ".join(missing_fields)
            with st.expander(f"[{pid}] {pname}", expanded=False):
                # 基本情報（読み取り専用）
                render_metric_row([
                    {"label": "坪数", "value": f"{tsubo_val} 坪"},
                    {"label": "業態", "value": category_val or "-"},
                    {"label": "合計金額（税抜）", "value": fmt_yen(total_val) if total_val else "-"},
                ])
                st.caption(f"未入力項目: {missing_labels}")

                col1, col2 = st.columns(2)

                # 工事種別
                cur_koji = str(p.get("工事種別", "")).strip()
                koji_opts = ["", "新装", "改装"]
                koji_idx = koji_opts.index(cur_koji) if cur_koji in koji_opts else 0
                with col1:
                    inp_koji = st.selectbox("工事種別", koji_opts, index=koji_idx, key=f"koji_{pid}",
                                            format_func=lambda x: x if x else "（未選択）")

                # 室外機設置階数
                cur_outdoor = str(p.get("室外機設置階数", "")).strip()
                outdoor_opts = [""] + OUTDOOR_UNIT_FLOOR_OPTIONS
                outdoor_idx = outdoor_opts.index(cur_outdoor) if cur_outdoor in outdoor_opts else 0
                with col1:
                    inp_outdoor = st.selectbox("室外機設置階数", outdoor_opts, index=outdoor_idx, key=f"outdoor_{pid}",
                                               format_func=lambda x: x if x else "（未選択）")

                # 工期
                cur_days_str = str(p.get("工期（日数）", "")).strip()
                days_opts = [""] + CONSTRUCTION_DAYS_OPTIONS
                days_idx = days_opts.index(cur_days_str) if cur_days_str in days_opts else 0
                with col1:
                    inp_days = st.selectbox("工期", days_opts, index=days_idx, key=f"days_{pid}",
                                            format_func=lambda x: x if x else "（未選択）")

                # 施工エリア
                cur_area = str(p.get("施工エリア", "")).strip()
                area_opts = [""] + CONSTRUCTION_AREA_OPTIONS
                area_idx = area_opts.index(cur_area) if cur_area in area_opts else 0
                with col2:
                    inp_area = st.selectbox("施工エリア", area_opts, index=area_idx, key=f"area_{pid}",
                                            format_func=lambda x: x if x else "（未選択）")

                # 入札結果
                cur_bid = str(p.get("入札結果", "")).strip()
                bid_opts = [""] + BID_RESULT_OPTIONS
                bid_idx = bid_opts.index(cur_bid) if cur_bid in bid_opts else 0
                with col2:
                    inp_bid = st.selectbox("入札結果", bid_opts, index=bid_idx, key=f"bid_{pid}",
                                           format_func=lambda x: x if x else "（未選択）")

                # その他備考
                cur_remarks = str(p.get("その他備考", "")).strip()
                with col2:
                    inp_remarks = st.text_area("その他備考", value=cur_remarks, max_chars=200,
                                               placeholder="夜間工事・居抜き・スケルトン等の特殊条件を記入",
                                               key=f"remarks_{pid}")

                # 入力値を記録
                all_inputs[pid] = {
                    "工事種別": inp_koji,
                    "室外機設置階数": inp_outdoor,
                    "工期（日数）": inp_days,
                    "施工エリア": inp_area,
                    "入札結果": inp_bid,
                    "その他備考": inp_remarks,
                }

                # 個別保存ボタン
                if st.button("この案件を保存", key=f"save_{pid}", type="primary"):
                    fields = {k: v for k, v in all_inputs[pid].items() if v != "" and v is not None}
                    if fields:
                        ok = ata_extract.update_project_fields(pid, fields)
                        if ok:
                            st.success(f"保存しました（{pid}）")
                            load_spreadsheet_data.clear()
                            _count_incomplete.clear()
                        else:
                            st.error(f"保存に失敗しました（{pid}）")
                    else:
                        st.warning("入力された値がありません。")

        # 一括保存ボタン
        st.markdown("---")
        if st.button("すべての案件をまとめて保存", type="primary", use_container_width=True, key="save_all_incomplete"):
            saved = 0
            failed = 0
            with st.spinner("保存中..."):
                for pid, fields_raw in all_inputs.items():
                    fields = {k: v for k, v in fields_raw.items() if v != "" and v is not None}
                    if not fields:
                        continue
                    ok = ata_extract.update_project_fields(pid, fields)
                    if ok:
                        saved += 1
                    else:
                        failed += 1
            if saved > 0:
                load_spreadsheet_data.clear()
                _count_incomplete.clear()
            if failed == 0 and saved > 0:
                st.success(f"{saved} 件の案件を保存しました。")
            elif failed > 0:
                st.warning(f"{saved} 件保存、{failed} 件失敗しました。")
            else:
                st.info("保存対象の入力値がありませんでした。")
