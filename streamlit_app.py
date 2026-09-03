import re
import base64
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from streamlit_folium import st_folium
import folium


# =========================================================
# 1. PAGE CONFIG
# =========================================================

st.set_page_config(
    page_title="Penanganan Rehabilitasi dan Rekonstruksi Pasca Bencana",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded"
)


# =========================================================
# HELPER FUNCTIONS
# =========================================================

def html(content: str):
    cleaned = re.sub(r"(?m)^[ \t]+", "", content.strip())
    st.markdown(cleaned, unsafe_allow_html=True)


def to_roman(n: int) -> str:
    vals = [
        (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"),
        (100, "C"), (90, "XC"), (50, "L"), (40, "XL"),
        (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I"),
    ]
    result = ""
    for v, sym in vals:
        while n >= v:
            result += sym
            n -= v
    return result


# =========================================================
# SINGKATAN UNIT ORGANISASI (BM / CK / SDA / PS)
# =========================================================

UNOR_ABBR_RULES = [
    ("bina marga", "BM"),
    ("cipta karya", "CK"),
    ("sumber daya air", "SDA"),
    ("prasarana strategis", "PS"),
]

UNOR_FULLNAME = {
    "BM": "Bina Marga",
    "CK": "Cipta Karya",
    "SDA": "Sumber Daya Air",
    "PS": "Prasarana Strategis",
}


def abbreviate_unor(name) -> str:
    if not isinstance(name, str):
        return str(name)
    low = name.lower()
    for keyword, abbr in UNOR_ABBR_RULES:
        if keyword in low:
            return abbr
    return name


def styled_dataframe(df, format_map, key=None):
    """Render dataframe dengan st.dataframe (native Streamlit table) + formatting."""
    st.dataframe(
        df.style.format(format_map),
        use_container_width=True,
        hide_index=True,
        key=key,
    )


@st.cache_data
def get_logo_base64(path: str):
    try:
        with open(path, "rb") as f:
            return base64.b64encode(f.read()).decode()
    except Exception:
        return None


LOGO_PATH = "logo-pu.png"
LOGO_B64 = get_logo_base64(LOGO_PATH)


# =========================================================
# ATURAN MAPPING KATEGORI (KEYWORD MAPPER)
# =========================================================

MAPPING_RULES = {
    'Rehabilitasi Jaringan Irigasi': ['irigasi', 'd.i.', 'd.i '],
    'Pengendalian Banjir dan Normalisasi Sungai': ['banjir', 'sungai', 'normalisasi'],
    'Pembangunan Cekdam': ['cekdam'],
    'Air Baku (Sumur air tanah)': ['air baku', 'sumur air'],
    'Pembangunan/Rehabilitasi Jalan': ['jalan'],
    'Pembangunan/Rehabilitasi Jembatan': ['jembatan'],
    'Rehabilitasi SPAM': ['spam'],
    'Sumur Bor': ['sumur bor'],
    'Rehabilitasi TPA': ['tpa'],
    'Rehabilitasi IPLT': ['iplt'],
    'Pembangunan Huntara': ['huntara'],
    'Rehabilitasi Sarpras Kesehatan': ['kesehatan', 'rsud', 'puskesmas'],
    'Rehabilitasi Sarpras Peribadatan': ['peribadatan', 'masjid', 'gereja'],
    'Rehabilitasi Sarpras Ponpes': ['ponpes', 'pesantren'],
    'Rehabilitasi Sarpras Madrasah': ['madrasah', 'man', 'mts', 'min']
}


def auto_categorize(nama_paket):
    if not isinstance(nama_paket, str):
        return "Lainnya / Cek Manual"
    
    nama_lower = nama_paket.lower()
    for kategori, keywords in MAPPING_RULES.items():
        if any(kw in nama_lower for kw in keywords):
            return kategori
    return "Lainnya / Cek Manual"


# =========================================================
# MAPPING CLUSTER (6 KELOMPOK BESAR UNTUK DONUT CHART & INFOGRAFIS)
# =========================================================

CLUSTER_MAP = {
    'Rehabilitasi Jaringan Irigasi': 'Irigasi, Rawa, & Sungai',
    'Pengendalian Banjir dan Normalisasi Sungai': 'Irigasi, Rawa, & Sungai',
    'Pembangunan Cekdam': 'Irigasi, Rawa, & Sungai',
    'Air Baku (Sumur air tanah)': 'Air Baku & Air Bersih',
    'Rehabilitasi SPAM': 'Air Baku & Air Bersih',
    'Sumur Bor': 'Air Baku & Air Bersih',
    'Pembangunan/Rehabilitasi Jalan': 'Konektivitas',
    'Pembangunan/Rehabilitasi Jembatan': 'Konektivitas',
    'Rehabilitasi TPA': 'Sanitasi & Persampahan',
    'Rehabilitasi IPLT': 'Sanitasi & Persampahan',
    'Pembangunan Huntara': 'Rumah Hunian & Fasilitas Umum',
    'Rehabilitasi Sarpras Kesehatan': 'Rumah Hunian & Fasilitas Umum',
    'Rehabilitasi Sarpras Peribadatan': 'Rumah Hunian & Fasilitas Umum',
    'Rehabilitasi Sarpras Ponpes': 'Rumah Hunian & Fasilitas Umum',
    'Rehabilitasi Sarpras Madrasah': 'Rumah Hunian & Fasilitas Umum',
}

CLUSTER_COLOR_MAP = {
    'Air Baku & Air Bersih': '#22b8c8',
    'Konektivitas': '#e5383b',
    'Sanitasi & Persampahan': '#f2b705',
    'Irigasi, Rawa, & Sungai': '#1e40af',
    'Rumah Hunian & Fasilitas Umum': '#b8860b',
    'Lainnya': '#94a3b8',
}

CLUSTER_ORDER = list(CLUSTER_COLOR_MAP.keys())


def map_to_cluster(kategori):
    return CLUSTER_MAP.get(kategori, "Lainnya")


# =========================================================
# DATA LOADING
# =========================================================

DEFAULT_FILE_PATH = "Paket Bencana RR status 02-09-2026.xlsx"
REQUIRED_SHEETS = ["Daftar RO", "Daftar Paket"]


@st.cache_data(show_spinner="Memproses data baru...")
def load_data(file_source):
    xls = pd.ExcelFile(file_source)
    missing = [s for s in REQUIRED_SHEETS if s not in xls.sheet_names]
    if missing:
        raise ValueError(
            f"Sheet {missing} tidak ditemukan. "
            f"Sheet yang tersedia di file: {xls.sheet_names}"
        )
    df_ro = pd.read_excel(xls, sheet_name="Daftar RO")
    df_paket = pd.read_excel(xls, sheet_name="Daftar Paket")
    
    # Cari kolom Nama Paket secara fleksibel
    nama_col = None
    for c in df_paket.columns:
        if "nama" in str(c).lower() and "paket" in str(c).lower():
            nama_col = c
            break
            
    if nama_col:
        df_paket["Kategori"] = df_paket[nama_col].apply(auto_categorize)
    elif "Nama Paket" in df_paket.columns:
        df_paket["Kategori"] = df_paket["Nama Paket"].apply(auto_categorize)

    if "Kategori" in df_paket.columns:
        df_paket["Cluster"] = df_paket["Kategori"].apply(map_to_cluster)

    return df_ro, df_paket


def find_col_by_keywords(df, keywords):
    for c in df.columns:
        cl = str(c).lower()
        if any(kw in cl for kw in keywords):
            return c
    return None


def find_lokasi_col(df):
    priority_names = [
        "Provinsi/Lokasi RO", "Lokasi RO", "Kabupaten/Kota",
        "Kab/Kota", "Kota/Kabupaten", "Lokasi",
    ]
    for name in priority_names:
        if name in df.columns:
            return name
    for c in df.columns:
        cl = c.lower()
        if "kabupaten" in cl or ("lokasi" in cl and "ro" in cl):
            return c
    return None


_LOKASI_PREFIXES = ["Kab. ", "Kabupaten "]


def normalize_lokasi(raw) -> str | None:
    if pd.isna(raw):
        return None
    s = str(raw).strip()
    if not s:
        return None
    for pref in _LOKASI_PREFIXES:
        if s.startswith(pref):
            s = s[len(pref):].strip()
            break
    return s or None


def is_specific_kabupaten(raw, provinsi_name: str) -> bool:
    if pd.isna(raw):
        return False
    s = str(raw).strip()
    if not s:
        return False
    low = s.lower()
    if "tersebar" in low:
        return False
    if "," in s or " dan " in low:
        return False
    if low.startswith("provinsi") or low.startswith("prov."):
        return False
    prov_clean = re.sub(r"(?i)^provinsi\s+", "", str(provinsi_name)).strip().lower()
    if low == prov_clean:
        return False
    norm = normalize_lokasi(s)
    if norm is None or len(norm) > 25:
        return False
    return True


def render_rincian_item(df_items: pd.DataFrame, level_label: str, level_name: str):
    if df_items.empty:
        st.info(f"Tidak ada data paket untuk {level_label} {level_name} pada filter saat ini.")
        return

    df_item = df_items.reset_index(drop=True).copy()

    df_item["Real. Keu (%)"] = (
        df_item["Realisasi (paket) (Rp ribu)"] / df_item["Pagu (paket) (Rp ribu)"] * 100
    ).fillna(0)

    if "No" in df_item.columns:
        df_item = df_item.drop(columns=["No"])

    df_item_base = df_item[
        [
            "Nama Paket",
            "Pagu (paket) (Rp ribu)",
            "Realisasi (paket) (Rp ribu)",
            "Real. Keu (%)",
            "Real. Fis (%)",
        ]
    ].rename(columns={
        "Pagu (paket) (Rp ribu)": "Pagu (Rp ribu)",
        "Realisasi (paket) (Rp ribu)": "Realisasi (Rp ribu)",
    })
    df_item_base["Real. Fis (%)"] = df_item_base["Real. Fis (%)"].fillna(0)

    if level_label == "Provinsi":
        subtitle = (
            f"Rincian item pada blok 'Provinsi {level_name}' "
            f"(paket yang tidak terikat kabupaten/kota tertentu), dalam Rp ribu"
        )
    else:
        subtitle = f"Rincian item paket yang berlokasi di {level_label} {level_name}, dalam Rp ribu"

    html(f"""
    <div style="font-size:0.95rem; font-weight:800; color:#1F4E78; margin-bottom:2px;">
        Ringkasan Paket Tingkat {level_label} - {level_name}
    </div>
    <div style="font-size:0.74rem; color:#94a3b8; font-family: 'Segoe UI', sans-serif; font-style: italic; margin-bottom:12px;">
        {subtitle}
    </div>
    """)

    sort_key = f"item_{level_label}_{level_name}"
    df_item_display = df_item_base.sort_values(
        "Pagu (Rp ribu)", ascending=False
    ).reset_index(drop=True)
    df_item_display.insert(0, "No", df_item_display.index + 1)

    styled_dataframe(
        df_item_display,
        {
            "Pagu (Rp ribu)": "{:,.0f}".format,
            "Realisasi (Rp ribu)": "{:,.0f}".format,
            "Real. Keu (%)": "{:.2f}%".format,
            "Real. Fis (%)": "{:.2f}%".format,
        },
        key=f"df_{sort_key}",
    )
    html('<div style="height:14px;"></div>')

    fig_combo = make_subplots(specs=[[{"secondary_y": True}]])

    fig_combo.add_trace(
        go.Bar(
            x=df_item_display["No"],
            y=df_item_display["Pagu (Rp ribu)"],
            name="Pagu (Rp ribu)",
            marker_color="#4682B4",
        ),
        secondary_y=False,
    )
    fig_combo.add_trace(
        go.Bar(
            x=df_item_display["No"],
            y=df_item_display["Realisasi (Rp ribu)"],
            name="Realisasi (Rp ribu)",
            marker_color="#C0392B",
        ),
        secondary_y=False,
    )
    fig_combo.add_trace(
        go.Scatter(
            x=df_item_display["No"],
            y=df_item_display["Real. Keu (%)"],
            name="Real. Keu (%)",
            mode="lines+markers",
            line=dict(color="#7CB5EC", width=3),
            marker=dict(size=6, symbol="diamond"),
        ),
        secondary_y=True,
    )
    fig_combo.add_trace(
        go.Scatter(
            x=df_item_display["No"],
            y=df_item_display["Real. Fis (%)"],
            name="Real. Fis (%)",
            mode="lines+markers",
            line=dict(color="#90ED7D", width=3),
            marker=dict(size=6, symbol="square"),
        ),
        secondary_y=True,
    )

    fig_combo.update_layout(
        title=dict(
            text=f"<b>Pagu vs Realisasi per Item - {level_name}</b>",
            font=dict(size=16, family="Segoe UI, Arial", color="#000000"),
            x=0.5,
            xanchor="center"
        ),
        barmode="group",
        height=450,
        paper_bgcolor="#ffffff",
        plot_bgcolor="#ffffff",
        font=dict(family="Segoe UI, Arial", size=11, color="#000000"),
        legend=dict(
            orientation="v",
            yanchor="top",
            y=0.7,
            xanchor="left",
            x=1.12,
            bordercolor="#d1d5db",
            borderwidth=1
        ),
        margin=dict(l=60, r=120, t=50, b=50),
        hovermode="x unified",
    )

    fig_combo.update_xaxes(
        title_text="",
        showgrid=True,
        gridcolor="#E0E0E0",
        dtick=1,
        tickfont=dict(size=11, color="#000000")
    )
    
    fig_combo.update_yaxes(
        title_text="<b>Rp ribu</b>",
        showgrid=True,
        gridcolor="#E0E0E0",
        secondary_y=False,
        tickformat=",.0f",
        zeroline=True,
        zerolinecolor="#000000"
    )

    fig_combo.update_yaxes(
        title_text="<b>Persentase (%)</b>",
        showgrid=False,
        secondary_y=True,
        ticksuffix="%",
        range=[0, max(100, df_item_display["Real. Keu (%)"].max() * 1.1)]
    )

    st.plotly_chart(fig_combo, use_container_width=True, config={"displayModeBar": False})


# =========================================================
# 2. CUSTOM CSS
# =========================================================

html("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
}

.stApp {
    background: #f5f7fb;
}

.block-container {
    padding-top: 1.5rem;
    padding-bottom: 3rem;
    max-width: 1500px;
}

/* SIDEBAR */
section[data-testid="stSidebar"] {
    background: #ffffff;
    border-right: 1px solid #e7ebf2;
}
section[data-testid="stSidebar"] > div {
    padding-top: 1.5rem;
}
section[data-testid="stSidebar"] div[data-testid="stVerticalBlock"] {
    gap: 0.6rem !important;
}

.sidebar-brand {
    padding: 6px 18px 16px 18px;
    margin-bottom: 4px;
    border-bottom: 1px solid #eef1f6;
    text-align: center;
}
.sidebar-brand img {
    width: 100%;
    max-width: 220px;
    height: auto;
    object-fit: contain;
}
.sidebar-brand-fallback {
    font-size: 1.05rem;
    font-weight: 800;
    color: #0f2747;
}

.sidebar-filters {
    padding: 4px 18px 0 18px;
    margin-top: 16px;
}
.sidebar-brand + .sidebar-filters {
    margin-top: 12px;
}

.filter-title {
    display: flex;
    align-items: center;
    gap: 6px;
    font-size: 0.72rem;
    font-weight: 700;
    color: #52627a;
    text-transform: uppercase;
    letter-spacing: 0.6px;
    margin-top: 14px;
    margin-bottom: 6px;
}
.filter-title:first-child {
    margin-top: 0;
}

section[data-testid="stSidebar"] div[data-baseweb="select"] > div {
    background-color: #f4f6fa !important;
    border: 1px solid #e2e8f0 !important;
    border-radius: 10px !important;
    min-height: 38px;
    box-shadow: none !important;
    transition: border-color 0.15s ease, background-color 0.15s ease;
}
section[data-testid="stSidebar"] div[data-baseweb="select"] > div:hover {
    background-color: #eef2f8 !important;
    border-color: #cbd5e1 !important;
}
section[data-testid="stSidebar"] div[data-baseweb="select"]:focus-within > div {
    background-color: #ffffff !important;
    border-color: #176b91 !important;
    box-shadow: 0 0 0 3px rgba(23, 107, 145, 0.12) !important;
}
section[data-testid="stSidebar"] div[data-baseweb="select"] span {
    color: #1e293b !important;
    font-weight: 600;
    font-size: 0.85rem;
}
section[data-testid="stSidebar"] div[data-baseweb="select"] svg {
    fill: #64748b !important;
}

/* FILE UPLOADER */
section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] {
    background-color: #f4f6fa !important;
    border: 1.5px dashed #cbd5e1 !important;
    border-radius: 12px !important;
    padding: 10px 8px !important;
}
section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"]:hover {
    border-color: #176b91 !important;
    background-color: #eef4f8 !important;
}
section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] small {
    color: #8290a3 !important;
}
section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] button {
    border-radius: 8px !important;
    border: 1px solid #dce3ec !important;
    background: #ffffff !important;
    color: #124d7c !important;
    font-weight: 700 !important;
}

.upload-status {
    display: flex;
    flex-direction: column;
    gap: 3px;
    font-size: 0.78rem;
    font-weight: 600;
    border-radius: 10px;
    padding: 10px 12px;
    margin-top: 8px;
}
.upload-ok {
    background: #eafaf1;
    border: 1px solid #b9ecd0;
    color: #0f7a4b;
}
.upload-default {
    background: #f4f6fa;
    border: 1px solid #e2e8f0;
    color: #64748b;
    font-weight: 500;
}
.upload-meta {
    font-size: 0.72rem;
    font-weight: 500;
    color: #4b9d78;
}

/* HEADER */
.dashboard-header {
    position: relative;
    overflow: hidden;
    padding: 30px 34px;
    border-radius: 18px;
    margin-bottom: 25px;
    background:
        radial-gradient(circle at 95% 15%, rgba(255, 196, 0, 0.28), transparent 25%),
        linear-gradient(135deg, #0d2b4f 0%, #124d7c 55%, #176b91 100%);
    box-shadow: 0 10px 30px rgba(15, 39, 71, 0.16);
}
.dashboard-header::after {
    content: "";
    position: absolute;
    right: -80px;
    bottom: -100px;
    width: 280px;
    height: 280px;
    border: 35px solid rgba(255,255,255,0.05);
    border-radius: 50%;
}
.header-label {
    color: #ffd43b;
    font-size: 0.75rem;
    font-weight: 800;
    text-transform: uppercase;
    letter-spacing: 1.3px;
    margin-bottom: 8px;
}
.dashboard-header h1 {
    color: #ffffff !important;
    font-size: 2rem;
    font-weight: 800;
    margin: 0;
    line-height: 1.2;
}
.dashboard-header p {
    color: #dceaf5;
    font-size: 0.92rem;
    margin-top: 10px;
    margin-bottom: 0;
}

/* SECTION TITLE */
.section-title {
    display: flex;
    align-items: center;
    gap: 10px;
    font-size: 1.05rem;
    font-weight: 800;
    color: #162b49;
    margin-top: 28px;
    margin-bottom: 15px;
}
.section-title::before {
    content: "";
    width: 5px;
    height: 22px;
    border-radius: 8px;
    background: #f5b700;
}

/* INFOGRAPHICS / CARD INFOGRAM STYLES */
.info-card-container {
    background: #ffffff;
    border-radius: 14px;
    border: 1px solid #e2e8f0;
    overflow: hidden;
    box-shadow: 0 4px 15px rgba(0,0,0,0.05);
    height: 100%;
}
.info-card-header {
    padding: 12px 16px;
    color: #ffffff;
    font-size: 0.95rem;
    font-weight: 800;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    display: flex;
    align-items: center;
    gap: 8px;
}
.info-card-body {
    padding: 14px 16px;
    background: #ffffff;
}
.info-item-row {
    display: flex;
    align-items: center;
    gap: 10px;
    margin-bottom: 10px;
    font-size: 0.82rem;
    color: #1e293b;
}
.info-item-row:last-child {
    margin-bottom: 0;
}
.info-item-icon {
    width: 28px;
    height: 28px;
    border-radius: 6px;
    background: #f1f5f9;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 14px;
    flex-shrink: 0;
}
.info-item-val {
    font-weight: 800;
    color: #0f172a;
}

/* METRIC CARD */
.metric-card {
    position: relative;
    overflow: hidden;
    background: #ffffff;
    border: 1px solid #e8edf3;
    border-radius: 16px;
    padding: 20px 21px;
    min-height: 135px;
    box-shadow: 0 5px 18px rgba(15, 39, 71, 0.055);
    transition: all 0.2s ease;
}
.metric-card:hover {
    transform: translateY(-3px);
    box-shadow: 0 10px 25px rgba(15, 39, 71, 0.10);
}
.metric-card::after {
    content: "";
    position: absolute;
    width: 80px;
    height: 80px;
    right: -25px;
    bottom: -30px;
    border-radius: 50%;
    background: rgba(15, 76, 129, 0.04);
}
.metric-top {
    display: flex;
    justify-content: space-between;
    align-items: center;
}
.metric-icon {
    width: 38px;
    height: 38px;
    display: flex;
    align-items: center;
    justify-content: center;
    border-radius: 10px;
    font-size: 18px;
}
.metric-title {
    color: #718096;
    font-size: 0.76rem;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}
.metric-value {
    color: #102a43;
    font-size: 1.48rem;
    font-weight: 800;
    margin-top: 13px;
}
.metric-sub {
    color: #8290a3;
    font-size: 0.76rem;
    margin-top: 4px;
}

/* CONTENT CARD */
.content-card {
    background: #ffffff;
    border: 1px solid #e8edf3;
    border-radius: 16px;
    padding: 18px 20px;
    box-shadow: 0 5px 18px rgba(15, 39, 71, 0.045);
    margin-bottom: 20px;
}

/* INFO BOX */
.info-box {
    background: #f8fafc;
    border: 1px solid #e8edf3;
    border-radius: 12px;
    padding: 14px 16px;
    color: #526173;
    font-size: 0.82rem;
}

/* TABLE */
div[data-testid="stDataFrame"] {
    border-radius: 12px;
    overflow: hidden;
}

/* TABS */
button[data-baseweb="tab"] {
    font-weight: 700;
    color: #64748b;
}
button[data-baseweb="tab"][aria-selected="true"] {
    color: #0f4c81;
}

/* BUTTON */
.stButton > button {
    border-radius: 9px;
    border: 1px solid #dce3ec;
    font-weight: 600;
}

/* FOOTER */
.dashboard-footer {
    text-align: center;
    color: #94a3b8;
    font-size: 0.72rem;
    padding-top: 25px;
    margin-top: 35px;
    border-top: 1px solid #e5eaf0;
}
</style>
""")


# =========================================================
# SIDEBAR
# =========================================================

with st.sidebar:
    if LOGO_B64:
        brand_content = f'<img src="data:image/png;base64,{LOGO_B64}" alt="Logo PU">'
    else:
        brand_content = '<div class="sidebar-brand-fallback">🏗️ Penanganan Rehabilitasi dan Rekonstruksi Pasca Bencana</div>'

    html(f"""
    <div class="sidebar-brand">{brand_content}</div>
    <div class="sidebar-filters">
        <div class="info-box">
            <b>💡 Filter Dashboard</b><br>
            Gunakan filter di atas untuk melihat kondisi paket
            berdasarkan unit organisasi, wilayah, kategori pekerjaan,
            dan jenis bencana.
        </div>
    </div>
    <div class="sidebar-filters">
    """)

    html('<div class="filter-title">📤&nbsp; Unggah Data Baru (.xlsx)</div>')
    uploaded_file = st.file_uploader(
        "Unggah data",
        type=["xlsx"],
        label_visibility="collapsed",
        help=(
            "File harus punya sheet 'Daftar RO' dan 'Daftar Paket' "
            "dengan format kolom yang sama seperti data awal."
        ),
    )

    file_source = uploaded_file if uploaded_file is not None else DEFAULT_FILE_PATH

    try:
        df_ro, df_paket = load_data(file_source)
        data_ok = True
    except Exception as e:
        data_ok = False
        st.error(f"Gagal memuat data.\n\n**Detail:** {e}")

    if uploaded_file is not None and data_ok:
        html(f"""
        <div class="upload-status upload-ok">
            ✅ <b>{uploaded_file.name}</b> berhasil dimuat &amp; dianalisis
            <span class="upload-meta">{len(df_paket):,} paket · {len(df_ro):,} baris RO</span>
        </div>
        """)
    elif uploaded_file is None:
        html("""
        <div class="upload-status upload-default">
            📁 Menggunakan data default (belum ada file diunggah)
        </div>
        """)

    html("</div>")

    if not data_ok:
        st.stop()

    LOKASI_COL = find_lokasi_col(df_paket)
    if LOKASI_COL is not None:
        df_paket["_lokasi_clean"] = df_paket.apply(
            lambda r: normalize_lokasi(r[LOKASI_COL])
            if is_specific_kabupaten(r[LOKASI_COL], r["Provinsi"]) else None,
            axis=1,
        )
    else:
        df_paket["_lokasi_clean"] = None

    html('<div class="sidebar-filters">')

    html('<div class="filter-title">🏢&nbsp; Unit Organisasi</div>')
    unor_options = ["Semua"] + sorted(
        df_paket["Unit Organisasi"].dropna().astype(str).unique().tolist()
    )
    selected_unor = st.selectbox(
        "Unit Organisasi", unor_options, label_visibility="collapsed"
    )

    html('<div class="filter-title">📍&nbsp; Provinsi</div>')
    prov_options = ["Semua"] + sorted(
        df_paket["Provinsi"].dropna().astype(str).unique().tolist()
    )
    selected_prov = st.selectbox(
        "Provinsi", prov_options, label_visibility="collapsed"
    )

    html('<div class="filter-title">🏘️&nbsp; Kabupaten/Kota</div>')
    if LOKASI_COL is not None:
        kab_source = df_paket if selected_prov == "Semua" else df_paket[df_paket["Provinsi"] == selected_prov]
        kab_options = ["Semua"] + sorted(
            kab_source["_lokasi_clean"].dropna().astype(str).unique().tolist()
        )
        selected_kab = st.selectbox(
            "Kabupaten/Kota", kab_options, label_visibility="collapsed"
        )
    else:
        selected_kab = "Semua"
        html("""
        <div class="info-box" style="font-size:0.74rem;">
            Kolom lokasi (Kabupaten/Kota) tidak ditemukan pada data 'Daftar Paket'.
        </div>
        """)

    html('<div class="filter-title">🧱&nbsp; Kategori Pekerjaan</div>')
    kategori_options = ["Semua"] + sorted(
        df_paket["Kategori"].dropna().astype(str).unique().tolist()
    )
    selected_kategori = st.selectbox(
        "Kategori", kategori_options, label_visibility="collapsed"
    )

    html('<div class="filter-title">🌪️&nbsp; Jenis Bencana</div>')
    bencana_options = ["Semua"] + sorted(
        df_paket["Jenis Bencana"].dropna().astype(str).unique().tolist()
    )
    selected_bencana = st.selectbox(
        "Jenis Bencana", bencana_options, label_visibility="collapsed"
    )

    html("</div>")


# =========================================================
# FILTER DATA
# =========================================================

df_filtered = df_paket.copy()

if selected_unor != "Semua":
    df_filtered = df_filtered[df_filtered["Unit Organisasi"] == selected_unor]

if selected_prov != "Semua":
    df_filtered = df_filtered[df_filtered["Provinsi"] == selected_prov]

if LOKASI_COL is not None and selected_kab != "Semua":
    df_filtered = df_filtered[df_filtered["_lokasi_clean"] == selected_kab]

if selected_kategori != "Semua":
    df_filtered = df_filtered[df_filtered["Kategori"] == selected_kategori]

if selected_bencana != "Semua":
    df_filtered = df_filtered[df_filtered["Jenis Bencana"] == selected_bencana]


# =========================================================
# HEADER
# =========================================================

html("""
<div class="dashboard-header">
    <div class="header-label">Dashboard Monitoring Infrastruktur</div>
    <h1>📊 Penanganan Rehabilitasi dan Rekonstruksi Pasca Bencana</h1>
    <p>Monitoring status implementasi, progres fisik, dan realisasi anggaran paket pemulihan bencana.</p>
</div>
""")


# =========================================================
# METRICS
# =========================================================

total_pagu = df_filtered["Pagu (paket) (Rp ribu)"].sum() * 1000
total_realisasi = df_filtered["Realisasi (paket) (Rp ribu)"].sum() * 1000
avg_real_keu = (total_realisasi / total_pagu * 100) if total_pagu > 0 else 0
avg_real_fisik = df_filtered["Real. Fis (%)"].mean() if not df_filtered.empty else 0
total_paket = len(df_filtered)


def rupiah_miliar(value):
    return f"Rp {value / 1e9:,.2f} M"


m1, m2, m3, m4 = st.columns(4)

with m1:
    html(f"""
    <div class="metric-card">
        <div class="metric-top">
            <div class="metric-title">Total Pagu</div>
            <div class="metric-icon" style="background:#e8f1fb;">💰</div>
        </div>
        <div class="metric-value">{rupiah_miliar(total_pagu)}</div>
        <div class="metric-sub">Anggaran yang dialokasikan</div>
    </div>
    """)

with m2:
    html(f"""
    <div class="metric-card">
        <div class="metric-top">
            <div class="metric-title">Realisasi Keuangan</div>
            <div class="metric-icon" style="background:#e8f8f1;">📈</div>
        </div>
        <div class="metric-value">{rupiah_miliar(total_realisasi)}</div>
        <div class="metric-sub" style="color:#10a66a;">▲ {avg_real_keu:.2f}% penyerapan</div>
    </div>
    """)

with m3:
    html(f"""
    <div class="metric-card">
        <div class="metric-top">
            <div class="metric-title">Progress Fisik</div>
            <div class="metric-icon" style="background:#fff5df;">🏗️</div>
        </div>
        <div class="metric-value">{avg_real_fisik:.2f}%</div>
        <div class="metric-sub">Rata-rata kemajuan pekerjaan</div>
    </div>
    """)

with m4:
    html(f"""
    <div class="metric-card">
        <div class="metric-top">
            <div class="metric-title">Total Paket</div>
            <div class="metric-icon" style="background:#eeeefe;">📦</div>
        </div>
        <div class="metric-value">{total_paket:,}</div>
        <div class="metric-sub">Paket terdaftar</div>
    </div>
    """)


# =========================================================
# CHART SECTION
# =========================================================

html('<div class="section-title">📊 Analisis Anggaran & Pekerjaan</div>')

col_left, col_right = st.columns([6, 4])

with col_left:
    with st.container(border=True):
        html('<b>Realisasi vs Pagu per Unit Organisasi</b>')

        df_unor_agg = (
            df_filtered
            .groupby("Unit Organisasi")[
                ["Pagu (paket) (Rp ribu)", "Realisasi (paket) (Rp ribu)"]
            ]
            .sum()
            .reset_index()
        )
        df_unor_agg["Unor Singkat"] = df_unor_agg["Unit Organisasi"].apply(abbreviate_unor)
        df_unor_agg["Pagu (Juta)"] = df_unor_agg["Pagu (paket) (Rp ribu)"] / 1_000_000
        df_unor_agg["Realisasi (Juta)"] = df_unor_agg["Realisasi (paket) (Rp ribu)"] / 1_000_000

        fig_bar = px.bar(
            df_unor_agg,
            x="Unor Singkat",
            y=["Pagu (Juta)", "Realisasi (Juta)"],
            barmode="group",
            labels={"value": "Nilai (Juta Rp)", "variable": "", "Unor Singkat": ""},
            text_auto=".2f",
            color_discrete_sequence=["#124d7c", "#f5b700"],
            hover_data={"Unit Organisasi": True},
        )
        fig_bar.update_traces(textposition="outside", textfont_size=10)
        fig_bar.update_layout(
            height=400,
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            margin=dict(l=10, r=10, t=30, b=20),
            font=dict(family="Inter", size=11, color="#475569"),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            hovermode="x unified",
        )
        fig_bar.update_xaxes(showgrid=False, title=None)
        fig_bar.update_yaxes(showgrid=True, gridcolor="#edf1f5", title=None)

        st.plotly_chart(fig_bar, use_container_width=True, config={"displayModeBar": False})

with col_right:
    with st.container(border=True):
        html('<b>Distribusi Cluster Pekerjaan</b>')

        df_cluster_agg = (
            df_filtered["Cluster"]
            .value_counts()
            .reindex(CLUSTER_ORDER)
            .dropna()
            .reset_index()
        )
        df_cluster_agg.columns = ["Cluster", "Jumlah"]

        fig_pie = px.pie(
            df_cluster_agg,
            names="Cluster",
            values="Jumlah",
            hole=0.58,
            color="Cluster",
            color_discrete_map=CLUSTER_COLOR_MAP,
        )
        fig_pie.update_traces(
            textposition="inside",
            textinfo="percent",
            hovertemplate="<b>%{label}</b><br>%{value} paket (%{percent})<extra></extra>",
        )
        fig_pie.update_layout(
            height=400,
            paper_bgcolor="rgba(0,0,0,0)",
            margin=dict(l=10, r=10, t=30, b=10),
            font=dict(family="Inter", size=11),
            legend=dict(orientation="h", yanchor="top", y=-0.15, xanchor="center", x=0.5),
        )

        st.plotly_chart(fig_pie, use_container_width=True, config={"displayModeBar": False})


# =========================================================
# PROGRESS SUMMARY
# =========================================================

html('<div class="section-title">📌 Ringkasan Progress</div>')

progress_col1, progress_col2 = st.columns(2)

with progress_col1:
    with st.container(border=True):
        html('<b>Penyerapan Keuangan</b>')
        st.progress(min(max(avg_real_keu / 100, 0), 1))
        html(f"""
        <div style="display:flex;justify-content:space-between;margin-top:5px;font-size:0.8rem;color:#64748b;">
            <span>Realisasi</span><b>{avg_real_keu:.2f}%</b>
        </div>
        """)

with progress_col2:
    with st.container(border=True):
        html('<b>Progress Fisik</b>')
        st.progress(min(max(avg_real_fisik / 100, 0), 1))
        html(f"""
        <div style="display:flex;justify-content:space-between;margin-top:5px;font-size:0.8rem;color:#64748b;">
            <span>Kemajuan pekerjaan</span><b>{avg_real_fisik:.2f}%</b>
        </div>
        """)


# =========================================================
# INFOGRAFIS TAMPILAN REKAPITULASI PENANGANAN BENCANA (INFOGRAPHIC CARDS)
# =========================================================

html('<div class="section-title">📋 Ringkasan Output Infrastruktur Penanganan Bencana</div>')

# Helper function kalkulasi agregasi per kategori dari dataframe terfilter
vol_col_info = find_col_by_keywords(df_filtered, ["vol", "volume", "panjang", "jumlah"])
satuan_col_info = find_col_by_keywords(df_filtered, ["satuan", "unit"])

def get_cat_summary(kategori_name):
    df_sub = df_filtered[df_filtered["Kategori"] == kategori_name]
    paket_cnt = len(df_sub)
    if vol_col_info and vol_col_info in df_sub.columns:
        vol_sum = pd.to_numeric(df_sub[vol_col_info], errors='coerce').fillna(0).sum()
        satuan = df_sub[satuan_col_info].dropna().iloc[0] if (satuan_col_info and not df_sub[satuan_col_info].dropna().empty) else "Unit"
        return f"{vol_sum:,.0f} {satuan}".strip()
    return f"{paket_cnt} Paket"

info_col1, info_col2 = st.columns(2)

with info_col1:
    # 1. AIR BAKU & AIR BERSIH
    html(f"""
    <div class="info-card-container" style="margin-bottom: 20px;">
        <div class="info-card-header" style="background: linear-gradient(90deg, #22b8c8, #0ea5e9);">
            💧 Air Baku &amp; Air Bersih
        </div>
        <div class="info-card-body">
            <div class="info-item-row">
                <div class="info-item-icon">🚰</div>
                <div><b>SPAM:</b> <span class="info-item-val">{get_cat_summary('Rehabilitasi SPAM')}</span></div>
            </div>
            <div class="info-item-row">
                <div class="info-item-icon">🏞️</div>
                <div><b>Air Baku:</b> <span class="info-item-val">{get_cat_summary('Air Baku (Sumur air tanah)')}</span></div>
            </div>
            <div class="info-item-row">
                <div class="info-item-icon">🕳️</div>
                <div><b>Sumur Bor:</b> <span class="info-item-val">{get_cat_summary('Sumur Bor')}</span></div>
            </div>
        </div>
    </div>
    """)

    # 2. KONEKTIVITAS
    html(f"""
    <div class="info-card-container" style="margin-bottom: 20px;">
        <div class="info-card-header" style="background: linear-gradient(90deg, #e5383b, #ef4444);">
            🛣️ Konektivitas
        </div>
        <div class="info-card-body">
            <div class="info-item-row">
                <div class="info-item-icon">🛣️</div>
                <div><b>Jalan:</b> <span class="info-item-val">{get_cat_summary('Pembangunan/Rehabilitasi Jalan')}</span></div>
            </div>
            <div class="info-item-row">
                <div class="info-item-icon">🌉</div>
                <div><b>Jembatan:</b> <span class="info-item-val">{get_cat_summary('Pembangunan/Rehabilitasi Jembatan')}</span></div>
            </div>
        </div>
    </div>
    """)

    # 3. SANITASI & PERSAMPAHAN
    html(f"""
    <div class="info-card-container">
        <div class="info-card-header" style="background: linear-gradient(90deg, #f2b705, #eab308);">
            🗑️ Sanitasi &amp; Persampahan
        </div>
        <div class="info-card-body">
            <div class="info-item-row">
                <div class="info-item-icon">🚛</div>
                <div><b>TPA:</b> <span class="info-item-val">{get_cat_summary('Rehabilitasi TPA')}</span></div>
            </div>
            <div class="info-item-row">
                <div class="info-item-icon">🚽</div>
                <div><b>IPLT:</b> <span class="info-item-val">{get_cat_summary('Rehabilitasi IPLT')}</span></div>
            </div>
        </div>
    </div>
    """)

with info_col2:
    # 4. IRIGASI, RAWA, & SUNGAI
    html(f"""
    <div class="info-card-container" style="margin-bottom: 20px;">
        <div class="info-card-header" style="background: linear-gradient(90deg, #1e40af, #2563eb);">
            🌾 Irigasi, Rawa, &amp; Sungai
        </div>
        <div class="info-card-body">
            <div class="info-item-row">
                <div class="info-item-icon">🌊</div>
                <div><b>Pengendalian Banjir &amp; Sungai:</b> <span class="info-item-val">{get_cat_summary('Pengendalian Banjir dan Normalisasi Sungai')}</span></div>
            </div>
            <div class="info-item-row">
                <div class="info-item-icon">🧱</div>
                <div><b>Cekdam:</b> <span class="info-item-val">{get_cat_summary('Pembangunan Cekdam')}</span></div>
            </div>
            <div class="info-item-row">
                <div class="info-item-icon">🌾</div>
                <div><b>Irigasi:</b> <span class="info-item-val">{get_cat_summary('Rehabilitasi Jaringan Irigasi')}</span></div>
            </div>
        </div>
    </div>
    """)

    # 5. RUMAH HUNIAN & FASILITAS UMUM
    html(f"""
    <div class="info-card-container">
        <div class="info-card-header" style="background: linear-gradient(90deg, #b8860b, #d97706);">
            🏠 Rumah Hunian &amp; Fasilitas Umum
        </div>
        <div class="info-card-body">
            <div class="info-item-row">
                <div class="info-item-icon">🏕️</div>
                <div><b>Huntara:</b> <span class="info-item-val">{get_cat_summary('Pembangunan Huntara')}</span></div>
            </div>
            <div class="info-item-row">
                <div class="info-item-icon">🏥</div>
                <div><b>Kesehatan:</b> <span class="info-item-val">{get_cat_summary('Rehabilitasi Sarpras Kesehatan')}</span></div>
            </div>
            <div class="info-item-row">
                <div class="info-item-icon">🕌</div>
                <div><b>Peribadatan:</b> <span class="info-item-val">{get_cat_summary('Rehabilitasi Sarpras Peribadatan')}</span></div>
            </div>
            <div class="info-item-row">
                <div class="info-item-icon">🏫</div>
                <div><b>Ponpes &amp; Madrasah:</b> <span class="info-item-val">{get_cat_summary('Rehabilitasi Sarpras Ponpes')} / {get_cat_summary('Rehabilitasi Sarpras Madrasah')}</span></div>
            </div>
        </div>
    </div>
    """)


# =========================================================
# ANALISIS & PENGELOMPOKAN KATEGORI (GROUP BY UNOR, PROVINSI, KATEGORI & SATUAN)
# =========================================================

html('<div class="section-title">🏷️ Analisis & Pengelompokan Kategori Paket</div>')
category_card = st.container(border=True)

vol_col = find_col_by_keywords(df_filtered, ["vol", "volume", "panjang", "jumlah"])
satuan_col = find_col_by_keywords(df_filtered, ["satuan", "unit"])

df_filtered_kat = df_filtered.copy()
if vol_col and satuan_col:
    df_filtered_kat[vol_col] = pd.to_numeric(df_filtered_kat[vol_col], errors='coerce').fillna(0)

category_card.markdown("""
<div style="font-size:0.95rem; font-weight:800; color:#1F4E78; margin-bottom:4px;">
    Rekapitulasi Paket & Volume Output Berdasarkan Unit Organisasi, Provinsi & Kategori
</div>
<div style="font-size:0.75rem; color:#64748b; margin-bottom:14px;">
    Tabel dikelompokkan berdasarkan <b>Unit Organisasi</b>, <b>Provinsi</b>, <b>Jenis Kegiatan (Kategori)</b>, dan dipecah terpisah per <b>Satuan Volume</b>.
</div>
""", unsafe_allow_html=True)

# Group By dengan Unit Organisasi & Provinsi di Paling Depan
custom_group = ["Unit Organisasi", "Provinsi", "Kategori"]
if satuan_col:
    custom_group.append(satuan_col)

if vol_col and satuan_col:
    df_grouped_unor = df_filtered_kat.groupby(custom_group).agg(
        **{
            "Total Volume": (vol_col, "sum"),
            "Jumlah Paket": ("Nama Paket", "count"),
            "Total Pagu (Rp ribu)": ("Pagu (paket) (Rp ribu)", "sum")
        }
    ).reset_index()
    
    df_grouped_unor["Volume Output"] = df_grouped_unor.apply(
        lambda r: f"{r['Total Volume']:,.2f} {r[satuan_col]}".rstrip('0').rstrip('.'), axis=1
    )
    
    df_display_unor = df_grouped_unor[
        ["Unit Organisasi", "Provinsi", "Kategori", "Volume Output", "Jumlah Paket", "Total Pagu (Rp ribu)"]
    ].rename(columns={"Kategori": "Jenis Kegiatan"})
else:
    df_grouped_unor = df_filtered_kat.groupby(["Unit Organisasi", "Provinsi", "Kategori"]).agg(
        **{
            "Jumlah Paket": ("Nama Paket", "count"),
            "Total Pagu (Rp ribu)": ("Pagu (paket) (Rp ribu)", "sum")
        }
    ).reset_index()
    df_display_unor = df_grouped_unor.rename(columns={"Kategori": "Jenis Kegiatan"})

df_display_unor = df_display_unor.sort_values(
    ["Unit Organisasi", "Provinsi", "Jenis Kegiatan"], 
    ascending=[True, True, True]
).reset_index(drop=True)

df_display_unor["_unor_abbr"] = df_display_unor["Unit Organisasi"].apply(abbreviate_unor)

UNOR_ORDER = ["BM", "CK", "SDA", "PS"]
UNOR_ICON = {"BM": "🛣️", "CK": "🏘️", "SDA": "💧", "PS": "🏗️"}

other_abbrs_kat = sorted(
    a for a in df_display_unor["_unor_abbr"].dropna().unique() if a not in UNOR_ORDER
)

kat_unor_tab_labels = [
    f"{UNOR_ICON.get(a, '📦')} {UNOR_FULLNAME.get(a, a)} ({a})" for a in UNOR_ORDER
]
kat_unor_tabs = category_card.tabs(kat_unor_tab_labels)

for tab, abbr in zip(kat_unor_tabs, UNOR_ORDER):
    with tab:
        df_kat_unor_sub = df_display_unor[df_display_unor["_unor_abbr"] == abbr].drop(columns=["_unor_abbr", "Unit Organisasi"])

        if df_kat_unor_sub.empty:
            st.info(f"Tidak ada data untuk unit organisasi {UNOR_FULLNAME.get(abbr, abbr)} pada filter saat ini.")
        else:
            st.dataframe(
                df_kat_unor_sub.style.format({"Total Pagu (Rp ribu)": "Rp {:,.0f}"}),
                use_container_width=True,
                hide_index=True,
            )

if other_abbrs_kat:
    with category_card.expander(f"📦 Unit organisasi lainnya di luar BM/CK/SDA/PS ({', '.join(other_abbrs_kat)})"):
        other_kat_tabs = st.tabs(other_abbrs_kat)
        for tab, abbr in zip(other_kat_tabs, other_abbrs_kat):
            with tab:
                df_kat_unor_sub = df_display_unor[df_display_unor["_unor_abbr"] == abbr].drop(columns=["_unor_abbr", "Unit Organisasi"])
                st.dataframe(
                    df_kat_unor_sub.style.format({"Total Pagu (Rp ribu)": "Rp {:,.0f}"}),
                    use_container_width=True,
                    hide_index=True,
                )

category_card.markdown('<div style="height:20px;"></div>', unsafe_allow_html=True)

kat_list = sorted(df_filtered_kat["Kategori"].dropna().unique().tolist())
if kat_list:
    selected_view_kat = category_card.selectbox(
        "🔎 Pilih Kategori Pekerjaan untuk melihat rincian paket dan volumenya:",
        kat_list
    )
    
    show_cols = ["Nama Paket", "Unit Organisasi", "Provinsi"]
    if vol_col:
        show_cols.append(vol_col)
    if satuan_col:
        show_cols.append(satuan_col)
    show_cols.extend(["Pagu (paket) (Rp ribu)", "Realisasi (paket) (Rp ribu)", "Real. Fis (%)"])
    
    df_kat_detail = df_filtered_kat[df_filtered_kat["Kategori"] == selected_view_kat][
        [c for c in show_cols if c in df_filtered_kat.columns]
    ].copy()
    
    df_kat_detail = df_kat_detail.rename(columns={
        "Pagu (paket) (Rp ribu)": "Pagu (Rp ribu)",
        "Realisasi (paket) (Rp ribu)": "Realisasi (Rp ribu)"
    })
    
    detail_fmt = {
        "Pagu (Rp ribu)": "{:,.0f}",
        "Realisasi (Rp ribu)": "{:,.0f}",
        "Real. Fis (%)": "{:.2f}%",
    }
    if vol_col and vol_col in df_kat_detail.columns:
        detail_fmt[vol_col] = "{:,.2f}"
        
    category_card.dataframe(
        df_kat_detail.style.format(detail_fmt),
        use_container_width=True,
        hide_index=True
    )


# =========================================================
# RINGKASAN REALISASI PER PROVINSI
# =========================================================

html('<div class="section-title">🗂️ Ringkasan Realisasi Paket per Provinsi</div>')

df_prov_summary = (
    df_filtered
    .groupby("Provinsi")
    .agg(**{
        "Pagu (Rp ribu)": ("Pagu (paket) (Rp ribu)", "sum"),
        "Realisasi (Rp ribu)": ("Realisasi (paket) (Rp ribu)", "sum"),
        "Real. Fis (%)": ("Real. Fis (%)", "mean"),
    })
    .reset_index()
)

df_prov_summary["Real. Keu (%)"] = (
    df_prov_summary["Realisasi (Rp ribu)"] / df_prov_summary["Pagu (Rp ribu)"] * 100
).fillna(0)

df_prov_summary = df_prov_summary[
    ["Provinsi", "Pagu (Rp ribu)", "Realisasi (Rp ribu)", "Real. Keu (%)", "Real. Fis (%)"]
].sort_values("Pagu (Rp ribu)", ascending=False)

with st.container(border=True):
    html("""
    <div style="font-size:0.95rem; font-weight:800; color:#1F4E78; margin-bottom:2px;">
        Ringkasan Realisasi Paket Rehabilitasi &amp; Rekonstruksi per Provinsi
    </div>
    <div style="font-size:0.74rem; color:#94a3b8; margin-bottom:12px;">
        Dihitung otomatis dari data yang sedang difilter (dalam Rp ribu)
    </div>
    """)

    if not df_prov_summary.empty:
        styled_dataframe(
            df_prov_summary,
            {
                "Pagu (Rp ribu)": "{:,.0f}",
                "Realisasi (Rp ribu)": "{:,.0f}",
                "Real. Keu (%)": "{:.2f}%",
                "Real. Fis (%)": "{:.2f}%",
            },
            key="df_prov_summary",
        )
    else:
        st.info("Tidak ada data provinsi untuk pilihan filter saat ini.")


# =========================================================
# RINGKASAN REALISASI PER KABUPATEN/KOTA
# =========================================================

html('<div class="section-title">🏘️ Ringkasan Realisasi Paket per Kabupaten/Kota</div>')
with st.container(border=True):
    if LOKASI_COL is None:
        st.info(
            "Kolom lokasi (Kabupaten/Kota) tidak ditemukan pada data (sheet 'Daftar Paket'). "
            "Tambahkan kolom seperti 'Provinsi/Lokasi RO' atau 'Kabupaten/Kota' pada data agar ringkasan ini bisa ditampilkan."
        )
    elif df_prov_summary.empty:
        st.info("Tidak ada data provinsi untuk pilihan filter saat ini.")
    else:
        province_list = df_prov_summary["Provinsi"].tolist()
        kab_tabs = st.tabs(province_list)

        for tab, prov in zip(kab_tabs, province_list):
            with tab:
                df_kab = df_filtered[
                    (df_filtered["Provinsi"] == prov)
                    & df_filtered["_lokasi_clean"].notna()
                ].copy()

                if df_kab.empty:
                    st.info(f"Tidak ada data kabupaten/kota untuk provinsi {prov} pada filter saat ini.")
                    continue

                df_kab_summary = (
                    df_kab
                    .groupby("_lokasi_clean")
                    .agg(**{
                        "Pagu (Rp ribu)": ("Pagu (paket) (Rp ribu)", "sum"),
                        "Realisasi (Rp ribu)": ("Realisasi (paket) (Rp ribu)", "sum"),
                        "Real. Fis (%)": ("Real. Fis (%)", "mean"),
                    })
                    .reset_index()
                    .rename(columns={"_lokasi_clean": "Kabupaten/Kota"})
                )

                df_kab_summary["Real. Keu (%)"] = (
                    df_kab_summary["Realisasi (Rp ribu)"] / df_kab_summary["Pagu (Rp ribu)"] * 100
                ).fillna(0)

                df_kab_summary = df_kab_summary.sort_values(
                    "Pagu (Rp ribu)", ascending=False
                ).reset_index(drop=True)

                df_kab_summary.insert(
                    0, "No", [to_roman(i + 1) for i in range(len(df_kab_summary))]
                )

                df_kab_summary = df_kab_summary[
                    ["No", "Kabupaten/Kota", "Pagu (Rp ribu)", "Realisasi (Rp ribu)", "Real. Keu (%)", "Real. Fis (%)"]
                ]

                styled_dataframe(
                    df_kab_summary,
                    {
                        "Pagu (Rp ribu)": "{:,.0f}",
                        "Realisasi (Rp ribu)": "{:,.0f}",
                        "Real. Keu (%)": "{:.2f}%",
                        "Real. Fis (%)": "{:.2f}%",
                    },
                    key=f"df_kab_summary_{prov}",
                )


# =========================================================
# RINCIAN PAKET PER PROVINSI / KABUPATEN
# =========================================================

html('<div class="section-title">📑 Rincian Paket per Provinsi / Kabupaten</div>')
detail_card = st.container(border=True)


def strip_provinsi_prefix(name: str) -> str:
    return re.sub(r"(?i)^provinsi\s+", "", str(name)).strip()


with detail_card:
    if LOKASI_COL is not None and selected_kab != "Semua":
        render_rincian_item(df_filtered, "Kabupaten/Kota", selected_kab)

    elif selected_prov != "Semua":
        df_item_prov = df_filtered[df_filtered["_lokasi_clean"].isna()].copy() \
            if LOKASI_COL is not None else df_filtered
        render_rincian_item(df_item_prov, "Provinsi", strip_provinsi_prefix(selected_prov))

    else:
        province_list = df_prov_summary["Provinsi"].tolist()

        if province_list:
            prov_tabs = st.tabs([strip_provinsi_prefix(p) for p in province_list])

            for tab, prov in zip(prov_tabs, province_list):
                with tab:
                    df_item_prov = df_filtered[
                        (df_filtered["Provinsi"] == prov)
                        & (df_filtered["_lokasi_clean"].isna() if LOKASI_COL is not None else True)
                    ].copy()
                    render_rincian_item(df_item_prov, "Provinsi", strip_provinsi_prefix(prov))
        else:
            st.info("Tidak ada data untuk ditampilkan pada pilihan filter saat ini.")


# =========================================================
# MAP
# =========================================================

html('<div class="section-title">🗺️ Sebaran Lokasi Paket</div>')

df_map = df_filtered.dropna(subset=["Latitude", "Longitude"])

if not df_map.empty:
    center_lat = df_map["Latitude"].mean()
    center_lon = df_map["Longitude"].mean()

    m = folium.Map(
        location=[center_lat, center_lon],
        zoom_start=6,
        tiles="CartoDB positron",
        control_scale=True,
    )

    for _, row in df_map.iterrows():
        progress = row["Real. Fis (%)"]
        if pd.isna(progress):
            progress = 0

        if progress < 30:
            marker_color, status_text = "red", "Progress Rendah"
        elif progress < 70:
            marker_color, status_text = "orange", "Dalam Proses"
        else:
            marker_color, status_text = "green", "Progress Baik"

        popup_html = f"""
        <div style="font-family:Arial;min-width:260px;padding:5px;">
            <div style="font-size:14px;font-weight:bold;color:#123;margin-bottom:8px;">
                {row['Nama Paket']}
            </div>
            <hr style="border:0;border-top:1px solid #ddd;">
            <div style="margin:6px 0;"><b>Provinsi</b><br>{row['Provinsi']}</div>
            <div style="margin:6px 0;"><b>Unit Organisasi</b><br>{row['Unit Organisasi']}</div>
            <div style="margin:6px 0;"><b>Pagu</b><br>Rp {row['Pagu (paket) (Rp ribu)'] * 1000:,.0f}</div>
            <div style="margin:6px 0;">
                <b>Progress Fisik</b><br>
                <span style="color:{marker_color};font-weight:bold;font-size:15px;">{progress:.2f}%</span>
                &nbsp;— {status_text}
            </div>
        </div>
        """

        folium.CircleMarker(
            location=[row["Latitude"], row["Longitude"]],
            radius=7,
            popup=folium.Popup(popup_html, max_width=330),
            tooltip=f"{row['Nama Paket']} — {progress:.1f}%",
            color=marker_color,
            weight=2,
            fill=True,
            fill_color=marker_color,
            fill_opacity=0.75,
        ).add_to(m)

    with st.container(border=True):
        st_folium(m, width=None, height=520, returned_objects=[])

else:
    st.info("Tidak ada koordinat lokasi yang tersedia untuk pilihan filter saat ini.")


# =========================================================
# TOP 10 PAKET BERDASARKAN REALISASI, PER UNIT ORGANISASI
# =========================================================

html('<div class="section-title">🏆 Top 10 Paket Berdasarkan Realisasi per Unit Organisasi</div>')

df_filtered_top = df_filtered.copy()
df_filtered_top["_unor_abbr"] = df_filtered_top["Unit Organisasi"].apply(abbreviate_unor)

# Kelompok lain di luar BM/CK/SDA/PS (kalau ada), supaya data tidak hilang begitu saja
other_abbrs = sorted(
    a for a in df_filtered_top["_unor_abbr"].dropna().unique() if a not in UNOR_ORDER
)

top10_card = st.container(border=True)

unor_tab_labels = [
    f"{UNOR_ICON.get(a, '📦')} {UNOR_FULLNAME.get(a, a)} ({a})" for a in UNOR_ORDER
]
unor_tabs = top10_card.tabs(unor_tab_labels)

for tab, abbr in zip(unor_tabs, UNOR_ORDER):
    with tab:
        df_unor_top = df_filtered_top[df_filtered_top["_unor_abbr"] == abbr]

        if df_unor_top.empty:
            st.info(f"Tidak ada paket untuk unit organisasi {UNOR_FULLNAME.get(abbr, abbr)} pada filter saat ini.")
        else:
            top10_unor = (
                df_unor_top
                .nlargest(10, "Realisasi (paket) (Rp ribu)")
                [["Nama Paket", "Provinsi", "Pagu (paket) (Rp ribu)", "Realisasi (paket) (Rp ribu)", "Real. Fis (%)"]]
                .copy()
            )
            top10_unor["Pagu (Rp)"] = top10_unor["Pagu (paket) (Rp ribu)"] * 1000
            top10_unor["Realisasi (Rp)"] = top10_unor["Realisasi (paket) (Rp ribu)"] * 1000
            top10_display_unor = top10_unor[
                ["Nama Paket", "Provinsi", "Pagu (Rp)", "Realisasi (Rp)", "Real. Fis (%)"]
            ]

            st.dataframe(
                top10_display_unor.style.format({
                    "Pagu (Rp)": "Rp {:,.0f}",
                    "Realisasi (Rp)": "Rp {:,.0f}",
                    "Real. Fis (%)": "{:.2f}%",
                }),
                use_container_width=True,
                hide_index=True,
            )

if other_abbrs:
    with st.expander(f"📦 Unit organisasi lainnya di luar BM/CK/SDA/PS ({', '.join(other_abbrs)})"):
        other_tabs = st.tabs(other_abbrs)
        for tab, abbr in zip(other_tabs, other_abbrs):
            with tab:
                df_unor_top = df_filtered_top[df_filtered_top["_unor_abbr"] == abbr]
                top10_unor = (
                    df_unor_top
                    .nlargest(10, "Realisasi (paket) (Rp ribu)")
                    [["Nama Paket", "Provinsi", "Pagu (paket) (Rp ribu)", "Realisasi (paket) (Rp ribu)", "Real. Fis (%)"]]
                    .copy()
                )
                top10_unor["Pagu (Rp)"] = top10_unor["Pagu (paket) (Rp ribu)"] * 1000
                top10_unor["Realisasi (Rp)"] = top10_unor["Realisasi (paket) (Rp ribu)"] * 1000
                st.dataframe(
                    top10_unor[["Nama Paket", "Provinsi", "Pagu (Rp)", "Realisasi (Rp)", "Real. Fis (%)"]]
                    .style.format({
                        "Pagu (Rp)": "Rp {:,.0f}",
                        "Realisasi (Rp)": "Rp {:,.0f}",
                        "Real. Fis (%)": "{:.2f}%",
                    }),
                    use_container_width=True,
                    hide_index=True,
                )


# =========================================================
# DETAIL DATA
# =========================================================

html('<div class="section-title">📋 Detail Data</div>')

tab1, tab2 = st.tabs(["📦 Daftar Paket", "📑 Rincian Output (RO)"])

with tab1:
    st.dataframe(df_filtered, use_container_width=True, hide_index=True)

with tab2:
    st.dataframe(df_ro, use_container_width=True, hide_index=True)


# =========================================================
# FOOTER
# =========================================================

html("""
<div class="dashboard-footer">
    Dashboard Monitoring Penanganan Rehabilitasi dan Rekonstruksi Pasca Bencana<br>
    Data diperbarui berdasarkan file status 02 September 2026
</div>
""")