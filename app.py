import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import numpy as np
from openpyxl import load_workbook
from datetime import timedelta, date
import requests
import json
import io

# ── Page Config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Monitoring PLTM Cilaki 1-B",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Share+Tech+Mono&family=Barlow:wght@300;400;600;700&display=swap');
  html, body, [class*="css"] { font-family: 'Barlow', sans-serif; }
  .stApp { background-color: #0a0e1a; }
  section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #0d1220 0%, #111827 100%);
    border-right: 1px solid #1e3a5f;
  }
  .metric-card {
    background: linear-gradient(135deg, #0f1f35 0%, #132338 100%);
    border: 1px solid #1e3a5f; border-radius: 12px;
    padding: 18px 20px; text-align: center;
    box-shadow: 0 4px 24px rgba(0,0,0,0.4);
    position: relative; overflow: hidden;
  }
  .metric-card::before {
    content:''; position:absolute; top:0; left:0; right:0; height:3px;
    background: linear-gradient(90deg, #00d4ff, #0077ff);
  }
  .metric-card.warning::before { background: linear-gradient(90deg,#ffb800,#ff6d00); }
  .metric-card.danger::before  { background: linear-gradient(90deg,#ff4444,#cc0000); }
  .metric-card.success::before { background: linear-gradient(90deg,#00e676,#00b248); }
  .metric-label { font-size:10px; font-weight:700; letter-spacing:2px; text-transform:uppercase; color:#4a7fa5; margin-bottom:6px; }
  .metric-value { font-family:'Share Tech Mono',monospace; font-size:28px; color:#e0f0ff; line-height:1; }
  .metric-sub   { font-size:11px; color:#4a7fa5; margin-top:5px; }
  .section-header {
    font-size:12px; font-weight:700; letter-spacing:3px; text-transform:uppercase;
    color:#00d4ff; border-bottom:1px solid #1e3a5f; padding-bottom:8px; margin-bottom:14px;
  }
  .dash-title    { font-family:'Share Tech Mono',monospace; font-size:26px; color:#00d4ff; letter-spacing:2px; }
  .dash-subtitle { font-size:12px; color:#4a7fa5; letter-spacing:1px; margin-top:-4px; }
  header[data-testid="stHeader"] { background:transparent; }
  .block-container { padding-top:1.5rem; }
</style>
""", unsafe_allow_html=True)


# ── Supabase REST API ─────────────────────────────────────────────────────────
SUPABASE_URL = st.secrets["SUPABASE_URL"]
SUPABASE_KEY = st.secrets["SUPABASE_KEY"]

HEADERS = {
    "apikey": SUPABASE_KEY,
    "Authorization": f"Bearer {SUPABASE_KEY}",
    "Content-Type": "application/json",
    "Prefer": "return=minimal"
}

def sb_select(table, params=""):
    url = f"{SUPABASE_URL}/rest/v1/{table}?{params}"
    r = requests.get(url, headers=HEADERS)
    if r.status_code == 200:
        return r.json()
    else:
        st.error(f"Error ambil data: {r.text}")
        return []

def sb_upsert(table, data):
    url = f"{SUPABASE_URL}/rest/v1/{table}"
    headers = {**HEADERS, "Prefer": "resolution=merge-duplicates"}
    r = requests.post(url, headers=headers, data=json.dumps(data))
    return r.status_code in [200, 201, 204]

def sb_delete(table, eq_col, eq_val):
    url = f"{SUPABASE_URL}/rest/v1/{table}?{eq_col}=eq.{eq_val}"
    r = requests.delete(url, headers=HEADERS)
    return r.status_code in [200, 204]


# ── Helpers ───────────────────────────────────────────────────────────────────
def td_to_str(val):
    if isinstance(val, timedelta):
        h = int(val.total_seconds()) // 3600
        return f"{h:02d}:00"
    return str(val) if val is not None else ""

def num(v):
    if isinstance(v, (int, float)) and not (isinstance(v, float) and np.isnan(v)):
        return float(v)
    return None

DARK_BG  = "#0d1a2d"
GRID_COL = "#1e3a5f"
FONT_COL = "#aac4e0"
LAYOUT   = dict(
    plot_bgcolor=DARK_BG, paper_bgcolor=DARK_BG,
    font=dict(color=FONT_COL, family="Barlow"),
    margin=dict(t=30, b=50, l=10, r=10)
)

def axis(title=""):
    return dict(title=title, gridcolor=GRID_COL, color="#4a7fa5", zerolinecolor=GRID_COL)

def kpi(col, label, value, sub, cls=""):
    col.markdown(
        f'<div class="metric-card {cls}">'
        f'<div class="metric-label">{label}</div>'
        f'<div class="metric-value">{value}</div>'
        f'<div class="metric-sub">{sub}</div>'
        f'</div>', unsafe_allow_html=True)


# ── Parse Excel harian ────────────────────────────────────────────────────────
def parse_excel_harian(file_bytes, tanggal_upload):
    wb = load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    target_sheet = None
    for name in wb.sheetnames:
        if name.isdigit():
            target_sheet = name
            break
        elif name.lower() not in ["rekap", "summary", "rekapitulasi"]:
            target_sheet = name
            break
    if not target_sheet:
        target_sheet = wb.sheetnames[0]

    ws   = wb[target_sheet]
    rows = list(ws.iter_rows(min_row=1, max_row=32, values_only=True))
    data = []

    for r in rows[3:27]:
        if r[0] is None:
            continue
        jam = td_to_str(r[0])
        if not jam or jam == "None":
            continue
        data.append({
            "tanggal"    : str(tanggal_upload),
            "jam"        : jam,
            "tg1_mw"     : num(r[2]),
            "tg1_pf"     : num(r[3]),
            "tg2_mw"     : num(r[30]) if len(r) > 30 else None,
            "tg3_mw"     : num(r[58]) if len(r) > 58 else None,
            "total_mw"   : num(r[86]) if len(r) > 86 else None,
            "total_pf"   : num(r[87]) if len(r) > 87 else None,
            "total_mvar" : num(r[88]) if len(r) > 88 else None,
            "volt_r"     : num(r[89]) if len(r) > 89 else None,
            "volt_s"     : num(r[90]) if len(r) > 90 else None,
            "volt_t"     : num(r[91]) if len(r) > 91 else None,
        })
    return data


# ── Load data ─────────────────────────────────────────────────────────────────
@st.cache_data(ttl=300)
def load_data():
    rows = sb_select("data_harian", "order=tanggal.asc,jam.asc&limit=10000")
    if not rows:
        return pd.DataFrame()
    df = pd.DataFrame(rows)
    df["tanggal"] = pd.to_datetime(df["tanggal"]).dt.date
    return df

def hitung_summary(df):
    if df.empty:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    cols = ["total_mw","tg1_mw","tg2_mw","tg3_mw","total_pf","total_mvar","volt_r","volt_s","volt_t"]
    df_max = df.groupby("tanggal")[cols].max().reset_index()
    df_avg = df.groupby("tanggal")[cols].mean().reset_index()
    df_min = df.groupby("tanggal")[cols].min().reset_index()
    return df_max, df_avg, df_min


# ════════════════════════════════════════════════════════════════════════════
#  SIDEBAR
# ════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("### ⚡ MONITORING PLTM CILAKI 1-B")
    st.markdown("---")
    st.markdown("**📂 Upload Data Harian**")

    tanggal_upload = st.date_input("Tanggal Data", value=date.today())
    uploaded = st.file_uploader("Pilih file Excel", type=["xlsx","xls"])

    if uploaded and st.button("💾 Simpan ke Database", type="primary", use_container_width=True):
        with st.spinner("Memproses dan menyimpan data..."):
            file_bytes = uploaded.read()
            data_list  = parse_excel_harian(file_bytes, tanggal_upload)
            if data_list:
                ok = sb_upsert("data_harian", data_list)
                if ok:
                    st.success(f"✅ {len(data_list)} data jam berhasil disimpan!")
                    st.cache_data.clear()
                    st.rerun()
                else:
                    st.error("❌ Gagal menyimpan ke database.")
            else:
                st.error("❌ Tidak ada data yang terbaca dari file.")

    st.markdown("---")
    st.markdown("**⚙️ Threshold Tegangan (kV)**")
    volt_min   = st.slider("Tegangan Minimum",  18.0, 21.0, 20.0, 0.1)
    volt_max_v = st.slider("Tegangan Maksimum", 20.0, 23.0, 21.5, 0.1)
    st.markdown("**⚙️ Batas Beban Normal**")
    beban_max  = st.slider("Beban Maks (MW)", 1.0, 5.0, 2.8, 0.1)
    st.markdown("---")
    st.markdown(
        "<div style='color:#4a7fa5;font-size:11px;line-height:1.7'>"
        "📊 Dashboard Monitoring<br>PLTM Cilaki 1-B<br>"
        "Sistem 20 kV · TG1·TG2·TG3"
        "</div>", unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════════════
#  HEADER
# ════════════════════════════════════════════════════════════════════════════
st.markdown(
    '<div class="dash-title">⚡ MONITORING PLTM CILAKI 1-B</div>'
    '<div class="dash-subtitle">DASHBOARD PROFIL TEGANGAN & BEBAN — REAL TIME</div>',
    unsafe_allow_html=True)
st.markdown("")

# ── Load data ─────────────────────────────────────────────────────────────────
df = load_data()

if df.empty:
    st.info("📭 Belum ada data. Silakan upload file Excel harian di sidebar.")
    st.markdown("""
    ### 📋 Cara Upload Data Harian
    1. Pilih **tanggal** data di sidebar
    2. Klik **Browse files** dan pilih file Excel harian
    3. Klik tombol **Simpan ke Database**
    4. Data otomatis tersimpan dan dashboard terupdate
    """)
    st.stop()

df_max, df_avg, df_min = hitung_summary(df)
df_max["tgl_str"] = df_max["tanggal"].astype(str)
df_avg["tgl_str"] = df_avg["tanggal"].astype(str)
df_min["tgl_str"] = df_min["tanggal"].astype(str)


# ── Tabs ──────────────────────────────────────────────────────────────────────
tab1, tab2, tab3, tab4 = st.tabs([
    "📊  Ringkasan",
    "📈  Profil Beban",
    "⚡  Profil Tegangan",
    "🗂️  Data Lengkap",
])


# ════════════════════════════════════════════════════════════════════════════
#  TAB 1 — RINGKASAN
# ════════════════════════════════════════════════════════════════════════════
with tab1:
    st.markdown('<div class="section-header">STATISTIK KESELURUHAN</div>', unsafe_allow_html=True)

    tot_days = df["tanggal"].nunique()
    tot_max  = df_max["total_mw"].max()
    tot_avg  = df_avg["total_mw"].mean()
    vR_avg   = df_avg["volt_r"].mean()
    pf_avg   = df_avg["total_pf"].mean()
    last_tgl = df["tanggal"].max()

    k1,k2,k3,k4,k5 = st.columns(5)
    kpi(k1, "Data Tersimpan",     f"{tot_days} Hari",   f"s/d {last_tgl}")
    kpi(k2, "Beban Tertinggi",    f"{tot_max:.2f} MW",  "sepanjang periode")
    kpi(k3, "Beban Rata-rata",    f"{tot_avg:.2f} MW",  "rata-rata harian")
    kpi(k4, "Tegangan Rata-rata", f"{vR_avg:.3f} kV",   "phasa R",
        "success" if volt_min <= vR_avg <= volt_max_v else "warning")
    kpi(k5, "Power Factor",       f"{pf_avg:.4f}",      "rata-rata",
        "success" if pf_avg >= 0.95 else "warning")

    st.markdown("")

    # Tren beban
    st.markdown('<div class="section-header">TREN BEBAN HARIAN (MW)</div>', unsafe_allow_html=True)
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=df_max["tgl_str"], y=df_max["total_mw"],
        mode="lines+markers", name="Beban Max",
        line=dict(color="#ff6b6b",width=2), marker=dict(size=6)))
    fig.add_trace(go.Scatter(x=df_avg["tgl_str"], y=df_avg["total_mw"],
        mode="lines+markers", name="Beban Rata-rata",
        line=dict(color="#00d4ff",width=2.5), marker=dict(size=6)))
    fig.add_trace(go.Scatter(x=df_min["tgl_str"], y=df_min["total_mw"],
        mode="lines+markers", name="Beban Min",
        line=dict(color="#00e676",width=2,dash="dot"), marker=dict(size=4)))
    fig.add_hline(y=beban_max, line_dash="dash", line_color="#ffb800",
                  annotation_text=f"Batas {beban_max} MW", annotation_font_color="#ffb800")
    fig.update_layout(**LAYOUT, height=350,
        xaxis=dict(**axis("Tanggal"), tickangle=-45),
        yaxis=axis("Total Beban (MW)"),
        legend=dict(bgcolor="#0a0e1a", bordercolor="#1e3a5f"))
    st.plotly_chart(fig, use_container_width=True)

    # Kontribusi per unit
    st.markdown('<div class="section-header">KONTRIBUSI BEBAN PER UNIT (MW)</div>', unsafe_allow_html=True)
    fig2 = go.Figure()
    fig2.add_trace(go.Bar(x=df_avg["tgl_str"], y=df_avg["tg1_mw"].fillna(0), name="TG1", marker_color="#00d4ff"))
    fig2.add_trace(go.Bar(x=df_avg["tgl_str"], y=df_avg["tg2_mw"].fillna(0), name="TG2", marker_color="#0077ff"))
    fig2.add_trace(go.Bar(x=df_avg["tgl_str"], y=df_avg["tg3_mw"].fillna(0), name="TG3", marker_color="#00e676"))
    fig2.update_layout(**LAYOUT, barmode="stack", height=300,
        xaxis=dict(**axis("Tanggal"), tickangle=-45),
        yaxis=axis("Beban (MW)"),
        legend=dict(bgcolor="#0a0e1a", bordercolor="#1e3a5f"))
    st.plotly_chart(fig2, use_container_width=True)

    c1, c2 = st.columns(2)
    with c1:
        st.markdown('<div class="section-header">POWER FACTOR HARIAN</div>', unsafe_allow_html=True)
        fig3 = go.Figure()
        fig3.add_trace(go.Scatter(x=df_avg["tgl_str"], y=df_avg["total_pf"],
            mode="lines+markers", name="PF Rata-rata",
            line=dict(color="#ffb800",width=2), marker=dict(size=5)))
        fig3.add_hline(y=0.95, line_dash="dot", line_color="#ff4444",
                       annotation_text="Min 0.95", annotation_font_color="#ff4444")
        fig3.update_layout(**LAYOUT, height=260,
            xaxis=dict(**axis(), tickangle=-45),
            yaxis=dict(**axis("Power Factor"), range=[0.93,1.0]))
        st.plotly_chart(fig3, use_container_width=True)

    with c2:
        st.markdown('<div class="section-header">DAYA REAKTIF Q (MVAr)</div>', unsafe_allow_html=True)
        fig4 = go.Figure()
        fig4.add_trace(go.Bar(x=df_avg["tgl_str"], y=df_avg["total_mvar"],
            name="Q Rata-rata", marker_color="#7c4dff"))
        fig4.add_trace(go.Scatter(x=df_max["tgl_str"], y=df_max["total_mvar"],
            mode="lines", name="Q Max", line=dict(color="#ff6b6b",dash="dot",width=2)))
        fig4.update_layout(**LAYOUT, height=260,
            xaxis=dict(**axis(), tickangle=-45),
            yaxis=axis("Q (MVAr)"))
        st.plotly_chart(fig4, use_container_width=True)


# ════════════════════════════════════════════════════════════════════════════
#  TAB 2 — PROFIL BEBAN HARIAN
# ════════════════════════════════════════════════════════════════════════════
with tab2:
    st.markdown('<div class="section-header">PROFIL BEBAN PER JAM</div>', unsafe_allow_html=True)

    tanggal_list = sorted(df["tanggal"].unique(), reverse=True)
    col_a, col_b = st.columns([1,3])

    with col_a:
        pilih_tgl = st.selectbox(
            "Pilih Tanggal", options=tanggal_list,
            format_func=lambda x: str(x))
        tampil_unit = st.multiselect(
            "Unit", ["TG1","TG2","TG3","Total"],
            default=["TG1","TG2","TG3","Total"])

    df_day = df[df["tanggal"] == pilih_tgl].copy()

    with col_b:
        d_max = df_max[df_max["tanggal"] == pilih_tgl]
        d_avg = df_avg[df_avg["tanggal"] == pilih_tgl]
        d_min = df_min[df_min["tanggal"] == pilih_tgl]
        k1,k2,k3,k4 = st.columns(4)
        if not d_max.empty:
            kpi(k1, "Beban Max", f"{d_max['total_mw'].values[0]:.2f} MW", str(pilih_tgl))
            kpi(k2, "Beban Avg", f"{d_avg['total_mw'].values[0]:.2f} MW" if not d_avg.empty else "-", "")
            kpi(k3, "Beban Min", f"{d_min['total_mw'].values[0]:.2f} MW" if not d_min.empty else "-", "")
            pf_v = d_avg['total_pf'].values[0] if not d_avg.empty else 0
            kpi(k4, "PF Rata-rata", f"{pf_v:.4f}", "",
                "success" if pf_v >= 0.95 else "warning")

    unit_colors = {"TG1":"#00d4ff","TG2":"#0077ff","TG3":"#00e676","Total":"#ffb800"}
    unit_cols   = {"TG1":"tg1_mw","TG2":"tg2_mw","TG3":"tg3_mw","Total":"total_mw"}

    fig = go.Figure()
    for unit in tampil_unit:
        col = unit_cols[unit]
        mask = df_day[col].notna() & (df_day[col] > 0)
        fig.add_trace(go.Scatter(
            x=df_day.loc[mask,"jam"], y=df_day.loc[mask,col],
            mode="lines+markers", name=unit,
            line=dict(color=unit_colors[unit],width=2.5), marker=dict(size=6)))
    fig.add_hline(y=beban_max, line_dash="dash", line_color="#ff4444",
                  annotation_text=f"Batas {beban_max} MW", annotation_font_color="#ff4444")
    fig.update_layout(**LAYOUT, height=380,
        title=f"Profil Beban — {pilih_tgl}",
        title_font=dict(color="#e0f0ff",size=14),
        xaxis=dict(**axis("Jam"), tickangle=-45),
        yaxis=axis("Beban (MW)"),
        legend=dict(bgcolor="#0a0e1a", bordercolor="#1e3a5f"))
    st.plotly_chart(fig, use_container_width=True)

    st.markdown('<div class="section-header">DATA PER JAM</div>', unsafe_allow_html=True)
    show_cols = ["jam","tg1_mw","tg2_mw","tg3_mw","total_mw","total_pf","total_mvar"]
    df_show   = df_day[show_cols].rename(columns={
        "jam":"Jam","tg1_mw":"TG1 (MW)","tg2_mw":"TG2 (MW)","tg3_mw":"TG3 (MW)",
        "total_mw":"Total (MW)","total_pf":"PF","total_mvar":"Q (MVAr)"})

    def hl(row):
        mw = row.get("Total (MW)", None)
        if mw and mw >= beban_max:
            return ["background-color:#2a0000;color:#ffcccc"] * len(row)
        return [""] * len(row)

    st.dataframe(df_show.style.apply(hl, axis=1).format(
        {"TG1 (MW)":"{:.2f}","TG2 (MW)":"{:.2f}","TG3 (MW)":"{:.2f}",
         "Total (MW)":"{:.2f}","PF":"{:.4f}","Q (MVAr)":"{:.3f}"}, na_rep="-"),
        use_container_width=True, hide_index=True, height=380)


# ════════════════════════════════════════════════════════════════════════════
#  TAB 3 — PROFIL TEGANGAN
# ════════════════════════════════════════════════════════════════════════════
with tab3:
    st.markdown('<div class="section-header">TREN TEGANGAN (kV)</div>', unsafe_allow_html=True)

    fig = go.Figure()
    fig.add_trace(go.Scatter(x=df_max["tgl_str"], y=df_max["volt_r"],
        mode="lines+markers", name="Max R",
        line=dict(color="#ff6b6b",width=2), marker=dict(size=5)))
    fig.add_trace(go.Scatter(x=df_avg["tgl_str"], y=df_avg["volt_r"],
        mode="lines+markers", name="Rata-rata R",
        line=dict(color="#00d4ff",width=2.5), marker=dict(size=5)))
    fig.add_trace(go.Scatter(x=df_min["tgl_str"], y=df_min["volt_r"],
        mode="lines+markers", name="Min R",
        line=dict(color="#00e676",width=2,dash="dot"), marker=dict(size=4)))
    fig.add_hrect(y0=volt_min, y1=volt_max_v,
        fillcolor="rgba(0,212,255,0.05)", line_width=0,
        annotation_text=f"Normal {volt_min}–{volt_max_v} kV",
        annotation_font_color="#4a7fa5", annotation_position="top left")
    fig.add_hline(y=volt_max_v, line_dash="dot", line_color="#ffb800", line_width=1)
    fig.add_hline(y=volt_min,   line_dash="dot", line_color="#ff4444", line_width=1)
    fig.update_layout(**LAYOUT, height=360,
        xaxis=dict(**axis("Tanggal"), tickangle=-45),
        yaxis=dict(**axis("Tegangan (kV)"), range=[volt_min-0.5, volt_max_v+0.3]),
        legend=dict(bgcolor="#0a0e1a", bordercolor="#1e3a5f"))
    st.plotly_chart(fig, use_container_width=True)

    st.markdown('<div class="section-header">PROFIL TEGANGAN PER JAM</div>', unsafe_allow_html=True)
    pilih_tgl_v = st.selectbox(
        "Pilih Tanggal", options=tanggal_list,
        format_func=lambda x: str(x), key="volt_day")
    df_dv = df[df["tanggal"] == pilih_tgl_v].copy()

    fig2 = go.Figure()
    ph = {"volt_r":"#ff6b6b","volt_s":"#ffb800","volt_t":"#00e676"}
    pn = {"volt_r":"Phasa R","volt_s":"Phasa S","volt_t":"Phasa T"}
    for col_name, color in ph.items():
        mask = df_dv[col_name].notna()
        fig2.add_trace(go.Scatter(
            x=df_dv.loc[mask,"jam"], y=df_dv.loc[mask,col_name],
            mode="lines+markers", name=pn[col_name],
            line=dict(color=color,width=2.5), marker=dict(size=5)))
    fig2.add_hrect(y0=volt_min, y1=volt_max_v, fillcolor="rgba(0,212,255,0.05)", line_width=0)
    fig2.add_hline(y=volt_max_v, line_dash="dot", line_color="#ffb800", line_width=1)
    fig2.add_hline(y=volt_min,   line_dash="dot", line_color="#ff4444", line_width=1)
    fig2.update_layout(**LAYOUT, height=340,
        title=f"Tegangan R/S/T — {pilih_tgl_v}",
        title_font=dict(color="#e0f0ff",size=14),
        xaxis=dict(**axis("Jam"), tickangle=-45),
        yaxis=dict(**axis("Tegangan (kV)"), range=[volt_min-1, volt_max_v+0.5]),
        legend=dict(bgcolor="#0a0e1a", bordercolor="#1e3a5f"))
    st.plotly_chart(fig2, use_container_width=True)

    mask_low  = df["volt_r"].notna() & (df["volt_r"] < volt_min)
    mask_high = df["volt_r"].notna() & (df["volt_r"] > volt_max_v)
    if (mask_low | mask_high).any():
        st.markdown('<div class="section-header">⚠️ PELANGGARAN BATAS TEGANGAN</div>', unsafe_allow_html=True)
        viol = df[mask_low | mask_high][["tanggal","jam","volt_r","volt_s","volt_t"]].copy()
        viol["Status"] = viol["volt_r"].apply(
            lambda v: "UNDERVOLTAGE" if v < volt_min else "OVERVOLTAGE")
        st.dataframe(viol, use_container_width=True, hide_index=True)
    else:
        st.success("✅ Tidak ada pelanggaran batas tegangan pada periode ini.")


# ════════════════════════════════════════════════════════════════════════════
#  TAB 4 — DATA LENGKAP
# ════════════════════════════════════════════════════════════════════════════
with tab4:
    st.markdown('<div class="section-header">DATA LENGKAP</div>', unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        tgl_dari = st.date_input("Dari Tanggal", value=df["tanggal"].min())
    with c2:
        tgl_ke   = st.date_input("Sampai Tanggal", value=df["tanggal"].max())

    df_filtered = df[(df["tanggal"] >= tgl_dari) & (df["tanggal"] <= tgl_ke)].copy()
    show_c = ["tanggal","jam","tg1_mw","tg2_mw","tg3_mw",
              "total_mw","total_pf","total_mvar","volt_r","volt_s","volt_t"]
    fmt = {c:"{:.3f}" for c in show_c if c not in ["tanggal","jam"]}
    fmt["total_pf"] = "{:.4f}"

    st.dataframe(df_filtered[show_c].style.format(fmt, na_rep="-"),
        use_container_width=True, hide_index=True, height=500)
    st.markdown(f"**Total: {len(df_filtered)} baris data**")

    st.markdown("")
    ex1, ex2 = st.columns(2)
    with ex1:
        csv_out = df_filtered[show_c].to_csv(index=False).encode()
        st.download_button("📥 Export CSV", data=csv_out,
            file_name=f"data_pltm_{tgl_dari}_{tgl_ke}.csv",
            mime="text/csv", use_container_width=True)
    with ex2:
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            df_filtered[show_c].to_excel(w, index=False, sheet_name="Data")
        st.download_button("📥 Export Excel", data=buf.getvalue(),
            file_name=f"data_pltm_{tgl_dari}_{tgl_ke}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)

    st.markdown("---")
    st.markdown('<div class="section-header">🗑️ HAPUS DATA</div>', unsafe_allow_html=True)
    st.warning("⚠️ Hati-hati! Data yang dihapus tidak bisa dikembalikan.")
    tgl_hapus = st.date_input("Pilih tanggal yang ingin dihapus")
    if st.button("🗑️ Hapus Data Tanggal Ini", type="secondary"):
        ok = sb_delete("data_harian", "tanggal", str(tgl_hapus))
        if ok:
            st.success(f"✅ Data tanggal {tgl_hapus} berhasil dihapus!")
            st.cache_data.clear()
            st.rerun()
        else:
            st.error("❌ Gagal menghapus data.")
