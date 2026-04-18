import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import numpy as np
from openpyxl import load_workbook
from datetime import timedelta
import io

# ── Page Config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Monitoring Profil Tegangan & Beban",
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


# ── Helpers ───────────────────────────────────────────────────────────────────
def td_to_str(val):
    if isinstance(val, timedelta):
        h = int(val.total_seconds()) // 3600
        return f"{h:02d}:00"
    return str(val) if val is not None else ""

def num(v):
    return v if isinstance(v, (int, float)) and not (isinstance(v, float) and np.isnan(v)) else None

DARK_BG  = "#0d1a2d"
GRID_COL = "#1e3a5f"
FONT_COL = "#aac4e0"
LAYOUT   = dict(plot_bgcolor=DARK_BG, paper_bgcolor=DARK_BG,
                font=dict(color=FONT_COL, family="Barlow"),
                margin=dict(t=30, b=50, l=10, r=10))

def axis(title=""):
    return dict(title=title, gridcolor=GRID_COL, color="#4a7fa5", zerolinecolor=GRID_COL)

def kpi(col, label, value, sub, cls=""):
    col.markdown(
        f'<div class="metric-card {cls}">'
        f'<div class="metric-label">{label}</div>'
        f'<div class="metric-value">{value}</div>'
        f'<div class="metric-sub">{sub}</div>'
        f'</div>', unsafe_allow_html=True)


# ── Load data ─────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner="Membaca data Excel…")
def load_all_sheets(file_bytes):
    wb = load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    daily_summary, hourly_all = [], []

    for sheet_name in [str(i) for i in range(1, 32)]:
        if sheet_name not in wb.sheetnames:
            continue
        ws  = wb[sheet_name]
        day = int(sheet_name)
        rows = list(ws.iter_rows(min_row=1, max_row=32, values_only=True))

        # Data per jam (baris index 3–26)
        for r in rows[3:27]:
            if r[0] is None:
                continue
            jam = td_to_str(r[0])
            hourly_all.append({
                "Tanggal": f"Jan-{day:02d}", "Hari": day, "Jam": jam,
                "TG1_MW"    : num(r[2]),
                "TG1_PF"    : num(r[3]),
                "TG2_MW"    : num(r[30]) if len(r) > 30 else None,
                "TG3_MW"    : num(r[58]) if len(r) > 58 else None,
                "Total_MW"  : num(r[86]) if len(r) > 86 else None,
                "Total_PF"  : num(r[87]) if len(r) > 87 else None,
                "Total_MVAR": num(r[88]) if len(r) > 88 else None,
                "Volt_R"    : num(r[89]) if len(r) > 89 else None,
                "Volt_S"    : num(r[90]) if len(r) > 90 else None,
                "Volt_T"    : num(r[91]) if len(r) > 91 else None,
            })

        # Summary Max/Rata2/Min (baris index 27–29)
        for r in rows[27:30]:
            label = r[0]
            if label not in ("Max", "Rata2", "Min"):
                continue
            daily_summary.append({
                "Tanggal": f"Jan-{day:02d}", "Hari": day, "Label": label,
                "TG1_MW"    : num(r[2]),
                "TG2_MW"    : num(r[30]) if len(r) > 30 else None,
                "TG3_MW"    : num(r[58]) if len(r) > 58 else None,
                "Total_MW"  : num(r[86]) if len(r) > 86 else None,
                "Total_PF"  : num(r[87]) if len(r) > 87 else None,
                "Total_MVAR": num(r[88]) if len(r) > 88 else None,
                "Volt_R"    : num(r[89]) if len(r) > 89 else None,
                "Volt_S"    : num(r[90]) if len(r) > 90 else None,
                "Volt_T"    : num(r[91]) if len(r) > 91 else None,
            })

    return pd.DataFrame(hourly_all), pd.DataFrame(daily_summary)


# ════════════════════════════════════════════════════════════════════════════
#  SIDEBAR
# ════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("### ⚡ MONITORING PENYULANG")
    st.markdown("---")
    st.markdown("**📂 Upload File Excel**")
    uploaded = st.file_uploader(
        "File Profil Tegangan & Beban",
        type=["xlsx", "xls"],
        help="Format: 01_Profil_Tegangan_dan_Beban_*.xlsx"
    )
    st.caption("Sheet: Rekap + sheet per tanggal (1–31)")
    st.markdown("---")
    st.markdown("**⚙️ Threshold Tegangan (kV)**")
    volt_min  = st.slider("Tegangan Minimum",  18.0, 21.0, 20.0, 0.1)
    volt_max  = st.slider("Tegangan Maksimum", 20.0, 23.0, 21.5, 0.1)
    st.markdown("**⚙️ Batas Beban Normal**")
    beban_max = st.slider("Beban Maks (MW)", 1.0, 5.0, 2.8, 0.1)
    st.markdown("---")
    st.markdown(
        "<div style='color:#4a7fa5;font-size:11px;line-height:1.7'>"
        "📊 Profil Tegangan & Beban<br>Penyulang — Januari 2025<br>"
        "Unit: TG1 · TG2 · TG3<br>Sistem 20 kV"
        "</div>", unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════════════
#  HEADER
# ════════════════════════════════════════════════════════════════════════════
st.markdown(
    '<div class="dash-title">⚡ PROFIL TEGANGAN & BEBAN PENYULANG</div>'
    '<div class="dash-subtitle">SISTEM MONITORING DISTRIBUSI — JANUARI 2025</div>',
    unsafe_allow_html=True)
st.markdown("")

if uploaded is None:
    st.info("👈 Upload file **Excel Profil Tegangan & Beban** di sidebar untuk memulai.")
    st.markdown("""
    ### 📋 Format File yang Didukung
    File Excel dengan struktur:
    - **Sheet `Rekap`** — ringkasan bulanan
    - **Sheet `1` s/d `31`** — data per tanggal, per jam (00:00–23:00)

    Setiap sheet harian memuat: **Power TG1/TG2/TG3 (MW)**, **Power Factor**,
    **Daya Reaktif Q (MVAr)**, **Tegangan R/S/T (kV)**, dan **Total** gabungan.
    """)
    st.stop()

# ── Load ──────────────────────────────────────────────────────────────────────
file_bytes = uploaded.read()
df_h, df_s = load_all_sheets(file_bytes)

if df_h.empty:
    st.error("Gagal membaca data. Pastikan format file sesuai.")
    st.stop()

df_avg = df_s[df_s["Label"] == "Rata2"].copy()
df_max = df_s[df_s["Label"] == "Max"].copy()
df_min = df_s[df_s["Label"] == "Min"].copy()

# ── Tabs ──────────────────────────────────────────────────────────────────────
tab1, tab2, tab3, tab4 = st.tabs([
    "📊  Ringkasan Bulanan",
    "📈  Profil Beban Harian",
    "⚡  Profil Tegangan",
    "🗂️  Data Lengkap",
])


# ════════════════════════════════════════════════════════════════════════════
#  TAB 1 — RINGKASAN BULANAN
# ════════════════════════════════════════════════════════════════════════════
with tab1:
    st.markdown('<div class="section-header">STATISTIK BULANAN — JANUARI 2025</div>', unsafe_allow_html=True)

    tot_max = df_max["Total_MW"].max()
    tot_avg = df_avg["Total_MW"].mean()
    tot_min = df_min["Total_MW"].min()
    vR_avg  = df_avg["Volt_R"].mean()
    vR_max  = df_max["Volt_R"].max()
    vR_min  = df_min["Volt_R"].min()
    pf_avg  = df_avg["Total_PF"].mean()

    k1,k2,k3,k4,k5 = st.columns(5)
    kpi(k1, "Beban Tertinggi",    f"{tot_max:.2f} MW", "sepanjang Januari")
    kpi(k2, "Beban Rata-rata",    f"{tot_avg:.2f} MW", "rata-rata harian")
    kpi(k3, "Beban Terendah",     f"{tot_min:.2f} MW", "sepanjang Januari",
        "success" if tot_min > 0 else "danger")
    kpi(k4, "Tegangan Rata-rata", f"{vR_avg:.3f} kV",
        f"Min {vR_min:.2f} / Maks {vR_max:.2f} kV",
        "success" if volt_min <= vR_avg <= volt_max else "warning")
    kpi(k5, "Power Factor",       f"{pf_avg:.4f}", "rata-rata bulanan",
        "success" if pf_avg >= 0.95 else "warning")

    st.markdown("")

    # Tren beban
    st.markdown('<div class="section-header">TREN BEBAN HARIAN (MW)</div>', unsafe_allow_html=True)
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=df_max["Tanggal"], y=df_max["Total_MW"],
        mode="lines+markers", name="Beban Max",
        line=dict(color="#ff6b6b", width=2), marker=dict(size=5)))
    fig.add_trace(go.Scatter(x=df_avg["Tanggal"], y=df_avg["Total_MW"],
        mode="lines+markers", name="Beban Rata-rata",
        line=dict(color="#00d4ff", width=2.5), marker=dict(size=5)))
    fig.add_trace(go.Scatter(x=df_min["Tanggal"], y=df_min["Total_MW"],
        mode="lines+markers", name="Beban Min",
        line=dict(color="#00e676", width=2, dash="dot"), marker=dict(size=4)))
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
    fig2.add_trace(go.Bar(x=df_avg["Tanggal"], y=df_avg["TG1_MW"].fillna(0), name="TG1", marker_color="#00d4ff"))
    fig2.add_trace(go.Bar(x=df_avg["Tanggal"], y=df_avg["TG2_MW"].fillna(0), name="TG2", marker_color="#0077ff"))
    fig2.add_trace(go.Bar(x=df_avg["Tanggal"], y=df_avg["TG3_MW"].fillna(0), name="TG3", marker_color="#00e676"))
    fig2.update_layout(**LAYOUT, barmode="stack", height=320,
        xaxis=dict(**axis("Tanggal"), tickangle=-45),
        yaxis=axis("Beban (MW)"),
        legend=dict(bgcolor="#0a0e1a", bordercolor="#1e3a5f"))
    st.plotly_chart(fig2, use_container_width=True)

    # PF & MVAR
    c1, c2 = st.columns(2)
    with c1:
        st.markdown('<div class="section-header">POWER FACTOR HARIAN</div>', unsafe_allow_html=True)
        fig3 = go.Figure()
        fig3.add_trace(go.Scatter(x=df_avg["Tanggal"], y=df_avg["Total_PF"],
            mode="lines+markers", name="PF Rata-rata",
            line=dict(color="#ffb800", width=2), marker=dict(size=4)))
        fig3.add_hline(y=0.95, line_dash="dot", line_color="#ff4444",
                       annotation_text="PF Min 0.95", annotation_font_color="#ff4444")
        fig3.update_layout(**LAYOUT, height=280,
            xaxis=dict(**axis(), tickangle=-45),
            yaxis=dict(**axis("Power Factor"), range=[0.94, 1.0]))
        st.plotly_chart(fig3, use_container_width=True)
    with c2:
        st.markdown('<div class="section-header">DAYA REAKTIF Q (MVAr)</div>', unsafe_allow_html=True)
        fig4 = go.Figure()
        fig4.add_trace(go.Bar(x=df_avg["Tanggal"], y=df_avg["Total_MVAR"],
            name="Q Rata-rata", marker_color="#7c4dff"))
        fig4.add_trace(go.Scatter(x=df_max["Tanggal"], y=df_max["Total_MVAR"],
            mode="lines", name="Q Max", line=dict(color="#ff6b6b", dash="dot", width=2)))
        fig4.update_layout(**LAYOUT, height=280,
            xaxis=dict(**axis(), tickangle=-45),
            yaxis=axis("Q (MVAr)"))
        st.plotly_chart(fig4, use_container_width=True)


# ════════════════════════════════════════════════════════════════════════════
#  TAB 2 — PROFIL BEBAN HARIAN (per jam)
# ════════════════════════════════════════════════════════════════════════════
with tab2:
    st.markdown('<div class="section-header">PROFIL BEBAN PER JAM</div>', unsafe_allow_html=True)

    col_a, col_b = st.columns([1, 3])
    with col_a:
        pilih_hari = st.selectbox(
            "Pilih Tanggal",
            options=sorted(df_h["Hari"].unique()),
            format_func=lambda x: f"Januari {x:02d}, 2025"
        )
        tampil_unit = st.multiselect(
            "Unit", ["TG1", "TG2", "TG3", "Total"],
            default=["TG1", "TG2", "TG3", "Total"])

    df_day = df_h[df_h["Hari"] == pilih_hari].copy()

    with col_b:
        d_max = df_max[df_max["Hari"] == pilih_hari]
        d_avg = df_avg[df_avg["Hari"] == pilih_hari]
        d_min = df_min[df_min["Hari"] == pilih_hari]
        k1,k2,k3,k4 = st.columns(4)
        if not d_max.empty:
            kpi(k1, "Beban Max", f"{d_max['Total_MW'].values[0]:.2f} MW", f"Jan-{pilih_hari:02d}")
            kpi(k2, "Beban Avg", f"{d_avg['Total_MW'].values[0]:.2f} MW" if not d_avg.empty else "-", "")
            kpi(k3, "Beban Min", f"{d_min['Total_MW'].values[0]:.2f} MW" if not d_min.empty else "-", "")
            pf_v = d_avg['Total_PF'].values[0] if not d_avg.empty else 0
            kpi(k4, "PF Rata-rata", f"{pf_v:.4f}", "",
                "success" if pf_v >= 0.95 else "warning")

    unit_colors = {"TG1":"#00d4ff","TG2":"#0077ff","TG3":"#00e676","Total":"#ffb800"}
    unit_cols   = {"TG1":"TG1_MW","TG2":"TG2_MW","TG3":"TG3_MW","Total":"Total_MW"}

    fig = go.Figure()
    for unit in tampil_unit:
        col = unit_cols[unit]
        mask = df_day[col].notna() & (df_day[col] > 0)
        fig.add_trace(go.Scatter(
            x=df_day.loc[mask,"Jam"], y=df_day.loc[mask,col],
            mode="lines+markers", name=unit,
            line=dict(color=unit_colors[unit], width=2.5), marker=dict(size=6)))
    fig.add_hline(y=beban_max, line_dash="dash", line_color="#ff4444",
                  annotation_text=f"Batas {beban_max} MW", annotation_font_color="#ff4444")
    fig.update_layout(**LAYOUT, height=380,
        title=f"Profil Beban — Januari {pilih_hari:02d}, 2025",
        title_font=dict(color="#e0f0ff", size=14),
        xaxis=dict(**axis("Jam"), tickangle=-45),
        yaxis=axis("Beban (MW)"),
        legend=dict(bgcolor="#0a0e1a", bordercolor="#1e3a5f"))
    st.plotly_chart(fig, use_container_width=True)

    # Tabel
    st.markdown('<div class="section-header">DATA PER JAM</div>', unsafe_allow_html=True)
    show_cols = ["Jam","TG1_MW","TG2_MW","TG3_MW","Total_MW","Total_PF","Total_MVAR"]
    df_show   = df_day[show_cols].rename(columns={
        "TG1_MW":"TG1 (MW)","TG2_MW":"TG2 (MW)","TG3_MW":"TG3 (MW)",
        "Total_MW":"Total (MW)","Total_PF":"PF","Total_MVAR":"Q (MVAr)"})

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
    st.markdown('<div class="section-header">TREN TEGANGAN BULANAN (kV)</div>', unsafe_allow_html=True)

    fig = go.Figure()
    fig.add_trace(go.Scatter(x=df_max["Tanggal"], y=df_max["Volt_R"],
        mode="lines+markers", name="Max Phasa R",
        line=dict(color="#ff6b6b",width=2), marker=dict(size=4)))
    fig.add_trace(go.Scatter(x=df_avg["Tanggal"], y=df_avg["Volt_R"],
        mode="lines+markers", name="Rata-rata Phasa R",
        line=dict(color="#00d4ff",width=2.5), marker=dict(size=4)))
    fig.add_trace(go.Scatter(x=df_min["Tanggal"], y=df_min["Volt_R"],
        mode="lines+markers", name="Min Phasa R",
        line=dict(color="#00e676",width=2,dash="dot"), marker=dict(size=4)))
    fig.add_hrect(y0=volt_min, y1=volt_max,
        fillcolor="rgba(0,212,255,0.05)", line_width=0,
        annotation_text=f"Normal {volt_min}–{volt_max} kV",
        annotation_font_color="#4a7fa5", annotation_position="top left")
    fig.add_hline(y=volt_max, line_dash="dot", line_color="#ffb800", line_width=1)
    fig.add_hline(y=volt_min, line_dash="dot", line_color="#ff4444", line_width=1)
    fig.update_layout(**LAYOUT, height=360,
        xaxis=dict(**axis("Tanggal"), tickangle=-45),
        yaxis=dict(**axis("Tegangan (kV)"), range=[volt_min-0.5, volt_max+0.3]),
        legend=dict(bgcolor="#0a0e1a", bordercolor="#1e3a5f"))
    st.plotly_chart(fig, use_container_width=True)

    # Profil per jam
    st.markdown('<div class="section-header">PROFIL TEGANGAN PER JAM</div>', unsafe_allow_html=True)
    pilih_hari_v = st.selectbox(
        "Pilih Tanggal",
        options=sorted(df_h["Hari"].unique()),
        format_func=lambda x: f"Januari {x:02d}, 2025",
        key="volt_day")
    df_dv = df_h[df_h["Hari"] == pilih_hari_v].copy()

    fig2 = go.Figure()
    ph = {"Volt_R":"#ff6b6b","Volt_S":"#ffb800","Volt_T":"#00e676"}
    pn = {"Volt_R":"Phasa R","Volt_S":"Phasa S","Volt_T":"Phasa T"}
    for col_name, color in ph.items():
        mask = df_dv[col_name].notna()
        fig2.add_trace(go.Scatter(
            x=df_dv.loc[mask,"Jam"], y=df_dv.loc[mask,col_name],
            mode="lines+markers", name=pn[col_name],
            line=dict(color=color,width=2.5), marker=dict(size=5)))
    fig2.add_hrect(y0=volt_min, y1=volt_max, fillcolor="rgba(0,212,255,0.05)", line_width=0)
    fig2.add_hline(y=volt_max, line_dash="dot", line_color="#ffb800", line_width=1)
    fig2.add_hline(y=volt_min, line_dash="dot", line_color="#ff4444", line_width=1)
    fig2.update_layout(**LAYOUT, height=340,
        title=f"Tegangan R/S/T — Januari {pilih_hari_v:02d}, 2025",
        title_font=dict(color="#e0f0ff", size=14),
        xaxis=dict(**axis("Jam"), tickangle=-45),
        yaxis=dict(**axis("Tegangan (kV)"), range=[volt_min-1, volt_max+0.5]),
        legend=dict(bgcolor="#0a0e1a", bordercolor="#1e3a5f"))
    st.plotly_chart(fig2, use_container_width=True)

    # Pelanggaran tegangan
    mask_low  = df_h["Volt_R"].notna() & (df_h["Volt_R"] < volt_min)
    mask_high = df_h["Volt_R"].notna() & (df_h["Volt_R"] > volt_max)
    if (mask_low | mask_high).any():
        st.markdown('<div class="section-header">⚠️ PELANGGARAN BATAS TEGANGAN</div>', unsafe_allow_html=True)
        viol = df_h[mask_low | mask_high][["Tanggal","Jam","Volt_R","Volt_S","Volt_T"]].copy()
        viol["Status"] = viol["Volt_R"].apply(
            lambda v: "UNDERVOLTAGE" if v < volt_min else "OVERVOLTAGE")
        st.dataframe(viol, use_container_width=True, hide_index=True)
    else:
        st.success("✅ Tidak ada pelanggaran batas tegangan pada periode ini.")


# ════════════════════════════════════════════════════════════════════════════
#  TAB 4 — DATA LENGKAP
# ════════════════════════════════════════════════════════════════════════════
with tab4:
    st.markdown('<div class="section-header">DATA LENGKAP HARIAN</div>', unsafe_allow_html=True)

    label_filter = st.radio("Tampilkan", ["Rata2","Max","Min"], horizontal=True)
    df_show2 = df_s[df_s["Label"] == label_filter].copy()
    show_c2  = ["Tanggal","TG1_MW","TG2_MW","TG3_MW",
                "Total_MW","Total_PF","Total_MVAR","Volt_R","Volt_S","Volt_T"]
    fmt2 = {c:"{:.3f}" for c in show_c2 if c != "Tanggal"}
    fmt2["Total_PF"] = "{:.4f}"

    st.dataframe(df_show2[show_c2].style.format(fmt2, na_rep="-"),
        use_container_width=True, hide_index=True, height=600)

    st.markdown("")
    c1, c2 = st.columns(2)
    with c1:
        csv_out = df_s[show_c2].to_csv(index=False).encode()
        st.download_button("📥 Export CSV", data=csv_out,
            file_name="profil_beban_jan2025.csv", mime="text/csv",
            use_container_width=True)
    with c2:
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            df_s[show_c2].to_excel(w, index=False, sheet_name="Summary")
            df_h.to_excel(w, index=False, sheet_name="Per Jam")
        st.download_button("📥 Export Excel", data=buf.getvalue(),
            file_name="profil_beban_jan2025.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)
