import os

import streamlit as st
import pandas as pd
import plotly.express as px

import data_store as ds

st.set_page_config(
    page_title="Analisis Bahan Tagih C0 & C1 FIFGROUP Cabang Rungkut",
    page_icon="📊",
    layout="wide"
)

# Letakkan file logo (logo_fifgroup.png) di folder yang sama dengan app.py
LOGO_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo_fifgroup.png")

# Styling ringan: tampilan terang & bersih
st.markdown(
    """
    <style>
    :root {color-scheme: light;}
    .stApp, [data-testid="stAppViewContainer"], [data-testid="stHeader"] {
        background-color: #FFFFFF !important;
        color: #0F172A !important;
    }
    [data-testid="stSidebar"], [data-testid="stSidebar"] > div {
        background-color: #F1F5F9 !important;
    }
    .stApp p, .stApp span, .stApp label, .stApp li,
    .stApp h1, .stApp h2, .stApp h3, .stApp h4,
    [data-testid="stMetricLabel"], [data-testid="stMetricValue"],
    [data-testid="stCaptionContainer"] {
        color: #0F172A !important;
    }
    [data-testid="stCaptionContainer"] {color: #64748B !important;}
    div[data-baseweb="select"] > div, div[data-baseweb="input"] > div {
        background-color: #FFFFFF !important;
        border-color: #CBD5E1 !important;
    }
    .block-container {padding-top: 2rem; max-width: 1300px;}
    [data-testid="stMetricValue"] {font-size: 1.45rem;}
    [data-testid="stMetricValue"], [data-testid="stMetricValue"] div,
    [data-testid="stMetricLabel"], [data-testid="stMetricLabel"] div, [data-testid="stMetricLabel"] p {
        overflow: visible !important;
        text-overflow: clip !important;
        white-space: normal !important;
        word-break: break-word;
    }
    [data-testid="stMetric"] {
        background: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 10px;
        padding: 14px 16px;
    }
    h1, h2, h3 {color: #0F172A;}
    </style>
    """,
    unsafe_allow_html=True
)

# ============================================================
# KONSTANTA
# ============================================================
STATUSES = ds.STATUSES
GAB = "Pelunasan + Pickup"
PCT_STATUSES = [GAB, "Rollback", "Rolling", "Settle"]

# True = higher is better, False = lower is better
HIGHER_BETTER = {GAB: True, "Rollback": True, "Rolling": False, "Settle": True}
DELTA_COLOR = {s: ("normal" if hb else "inverse") for s, hb in HIGHER_BETTER.items()}

PRODUK_PREFERRED = ["MMU", "MPF", "NMC", "REFI"]
PALETTE = ["#2563EB", "#F97316", "#10B981", "#8B5CF6", "#EF4444", "#0EA5E9", "#F59E0B", "#14B8A6"]


# ============================================================
# HEADER + LOGO
# ============================================================
logo_col, title_col = st.columns([1, 6], vertical_alignment="center")

with logo_col:
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, width=130)
    else:
        st.markdown("**FIFGROUP**  \n*member of ASTRA*")

title_col_container = title_col.container()

if os.path.exists(LOGO_PATH):
    st.sidebar.image(LOGO_PATH, width=90)

# Pesan hasil aksi terakhir (upload / hapus / restore)
if "flash" in st.session_state:
    st.success(st.session_state.pop("flash"))

# ============================================================
# KELOLA DATA (upload Excel, hapus bulan, backup)
# ============================================================
if "uploader_ver" not in st.session_state:
    st.session_state["uploader_ver"] = 0

data = ds.load_data()
existing_periods = ds.ordered_periods(data) if not data.empty else []

with st.expander("📂 Kelola Data (upload Excel, hapus bulan, backup)", expanded=data.empty):
    tab_up, tab_del, tab_bak = st.tabs(["Upload bulan baru", "Hapus bulan", "Backup & restore"])

    # ---- Upload ----
    with tab_up:
        st.caption(
            "Upload file pivot Excel (.xlsx) per bulan. Satu sheet = satu bulan. "
            "Bisa upload beberapa file sekaligus. Label periode terbaca otomatis dari sel "
            "seperti 'Bulan 7' (jadi B7) dan bisa diubah. Label yang sudah ada akan ditimpa."
        )
        ver = st.session_state["uploader_ver"]
        files = st.file_uploader(
            "Pilih file Excel", type=["xlsx"], accept_multiple_files=True, key=f"uploader_{ver}"
        )

        entries = []
        for f in files or []:
            try:
                parsed = ds.parse_workbook(f)
            except Exception as e:
                st.error(f"{f.name}: file tidak bisa dibaca ({e})")
                continue
            for res in parsed:
                with st.container(border=True):
                    st.markdown(f"**{f.name}** - sheet `{res['sheet']}`")
                    if res["error"]:
                        st.error(res["error"])
                        continue
                    label = st.text_input(
                        "Label periode", value=res["label"],
                        key=f"lbl_{ver}_{f.name}_{res['sheet']}",
                        help="Contoh: B10. Urutan periode mengikuti angka pada label."
                    ).strip()
                    rows = res["rows"].copy()
                    rows["Periode"] = label
                    n_prod = rows["Produk"].nunique()
                    n_grp = rows["Kelompok"].nunique()
                    total = int(rows[STATUSES].to_numpy().sum())
                    if label in existing_periods:
                        st.warning(f"Periode {label} sudah ada dan akan ditimpa.")
                    elif label:
                        st.info(f"Periode baru: {label}")
                    st.caption(
                        f"{len(rows)} baris ({n_grp} kelompok x {n_prod} produk). "
                        f"Grand Total: Rp {total:,.0f}".replace(",", ".")
                    )
                    for w in res["warnings"]:
                        st.warning(w)
                    entries.append((f"{f.name} / {res['sheet']}", label, rows))

        if entries and st.button("💾 Simpan ke data", type="primary"):
            labels = [e[1] for e in entries]
            if any(not l for l in labels):
                st.error("Ada label periode yang masih kosong.")
            elif len(set(labels)) != len(labels):
                st.error("Ada label periode yang sama di antara file yang diupload.")
            else:
                new_df = pd.concat([e[2] for e in entries], ignore_index=True)
                ds.save_data(ds.upsert_periods(data, new_df))
                st.session_state["uploader_ver"] += 1
                st.session_state["flash"] = "Data tersimpan: " + ", ".join(labels) + "."
                st.rerun()

    # ---- Hapus ----
    with tab_del:
        if not existing_periods:
            st.info("Belum ada data.")
        else:
            to_del = st.multiselect("Pilih bulan yang dihapus", existing_periods, key="del_sel")
            ok = st.checkbox("Saya yakin ingin menghapus bulan tersebut", key="del_ok")
            if st.button("🗑️ Hapus", disabled=not (to_del and ok)):
                ds.save_data(ds.delete_periods(data, to_del))
                st.session_state["flash"] = "Dihapus: " + ", ".join(to_del) + "."
                st.rerun()

    # ---- Backup ----
    with tab_bak:
        st.caption(
            "Data disimpan di file data_bahan_tagih.csv di folder aplikasi. Bila di-deploy ke hosting "
            "yang filenya tidak permanen (mis. Streamlit Community Cloud), download backup ini "
            "secara berkala dan restore bila data hilang."
        )
        st.download_button(
            "⬇️ Download backup (CSV)",
            data=data.to_csv(index=False).encode("utf-8"),
            file_name="data_bahan_tagih.csv", mime="text/csv", disabled=data.empty
        )
        bak = st.file_uploader("Restore dari backup CSV (mengganti seluruh data)", type=["csv"], key="restore_up")
        if bak is not None and st.button("♻️ Restore data dari CSV"):
            try:
                ds.save_data(ds.validate_backup(pd.read_csv(bak)))
                st.session_state["flash"] = "Data berhasil di-restore dari CSV."
                st.rerun()
            except Exception as e:
                st.error(f"Restore gagal: {e}")
        reset_ok = st.checkbox("Saya yakin ingin menghapus SEMUA data", key="reset_ok")
        if st.button("Hapus semua data", disabled=not (reset_ok and not data.empty)):
            ds.save_data(ds.empty_df())
            st.session_state["flash"] = "Semua data dihapus."
            st.rerun()

if data.empty:
    st.info("Belum ada data. Upload file pivot Excel pada bagian Kelola Data di atas.")
    st.stop()

# ============================================================
# DATA DINAMIS
# ============================================================
PERIODES = ds.ordered_periods(data)
GROUPS_ALL = sorted(data["Kelompok"].unique())
_prod_all = set(data["Produk"].unique())
PRODUKS = [p for p in PRODUK_PREFERRED if p in _prod_all] + sorted(_prod_all - set(PRODUK_PREFERRED))

COLORS_GROUP = {g: PALETTE[i % len(PALETTE)] for i, g in enumerate(GROUPS_ALL)}
if set(GROUPS_ALL) <= {"C0", "C1"}:
    COLORS_GROUP = {"C0": "#2563EB", "C1": "#F97316"}
COLORS_STATUS = ["#0EA5E9", "#EF4444", "#F59E0B", "#10B981"]
COLORS_PRODUK = {
    "MMU": "#2563EB", "MPF": "#10B981", "NMC": "#F59E0B", "REFI": "#8B5CF6"
}
for i, p in enumerate([p for p in PRODUKS if p not in COLORS_PRODUK]):
    COLORS_PRODUK[p] = PALETTE[(i + 4) % len(PALETTE)]

with title_col_container:
    st.title("Analisis (Pelunasan + Pickup), Rollback, Rolling & Settle")
    st.caption(
        f"Dashboard interaktif {' & '.join(GROUPS_ALL)} per produk ({', '.join(PRODUKS)}) "
        f"berdasarkan data {', '.join(PERIODES)}. "
        "Seluruh amount dalam Rupiah (Sum of PKK_AW). "
        "Pelunasan + Pickup = (Pelunasan + Pickup) / Grand Total."
    )

raw_df = data.copy()
raw_df["Grand Total"] = raw_df[STATUSES].sum(axis=1)
raw_df[GAB] = raw_df["Pelunasan"] + raw_df["Pickup"]
raw_df["Pembagi"] = raw_df["Grand Total"]  # penyebut persentase = Grand Total baris itu


# ============================================================
# FORMAT & STYLE
# ============================================================
def rupiah(value):
    return f"Rp {value:,.0f}".replace(",", ".")

def pct_fmt(value):
    return f"{value:.2f}".replace(".", ",") + "%"

def delta_text(d):
    return f"{d:+.2f} poin".replace(".", ",")

def style_fig(fig, pct=False, amount=False):
    fig.update_layout(
        template="plotly_white",
        separators=",.",
        legend_title_text="",
        margin=dict(t=40, b=20, l=10, r=10),
        font=dict(size=13, color="#1E293B"),
    )
    if pct:
        fig.update_yaxes(ticksuffix="%")
    if amount:
        fig.update_yaxes(tickformat=",.0f")
    fig.for_each_annotation(lambda a: a.update(text=a.text.split("=")[-1]))
    return fig

def aggregate(frame, by):
    """Jumlahkan amount & pembagi, lalu hitung persentase (bukan rata-rata persentase)."""
    cols = STATUSES + [GAB, "Grand Total", "Pembagi"]
    g = frame.groupby(by, as_index=False)[cols].sum()
    for s in PCT_STATUSES:
        g[f"{s} %"] = (g[s] / g["Pembagi"].replace(0, float("nan")) * 100).fillna(0)
    return g

def pick(frame, **conds):
    """Ambil satu baris sesuai kondisi; None bila tidak ada (mis. produk baru belum ada di bulan lama)."""
    mask = pd.Series(True, index=frame.index)
    for k, v in conds.items():
        mask &= frame[k] == v
    sub = frame[mask]
    return None if sub.empty else sub.iloc[0]

def kualitas(status, d):
    if abs(d) < 0.005:
        return "tetap"
    return "membaik" if (d > 0) == HIGHER_BETTER[status] else "memburuk"

def metric_row(r, prev, g, p, label):
    """Satu baris metric persentase (dengan selisih vs bulan sebelumnya bila ada)."""
    st.markdown(f"**{g}**")
    cols = st.columns(len(PCT_STATUSES))
    for c, status in zip(cols, PCT_STATUSES):
        delta = None
        if prev is not None:
            delta = delta_text(r[f"{status} %"] - prev[f"{status} %"])
        if status == GAB:
            pembilang = (f"({rupiah(r['Pelunasan'])} + {rupiah(r['Pickup'])}) "
                         f"= {rupiah(r[GAB])}")
        else:
            pembilang = rupiah(r[status])
        c.metric(
            status, pct_fmt(r[f"{status} %"]), delta=delta, delta_color=DELTA_COLOR[status],
            help=(f"{status} {label} {g} {p}  \n"
                  f"Pembilang: {pembilang}  \n"
                  f"Grand Total (pembagi): {rupiah(r['Pembagi'])}  \n"
                  f"= {pct_fmt(r[f'{status} %'])}")
        )

ARAH_CAPTION = (
    "Warna selisih: hijau = membaik, merah = memburuk. "
    "Higher is better: Pelunasan + Pickup, Rollback, Settle. Lower is better: Rolling."
)

# ============================================================
# SIDEBAR
# ============================================================
st.sidebar.header("Filter")

periode = st.sidebar.multiselect("Pilih periode", PERIODES, default=PERIODES[-3:])
kelompok = st.sidebar.radio("Pilih kelompok", ["Semua"] + GROUPS_ALL, index=0)
produk = st.sidebar.multiselect("Pilih produk", PRODUKS, default=PRODUKS)

st.sidebar.caption(ARAH_CAPTION)

if not periode or not produk:
    st.warning("Pilih minimal satu periode dan satu produk.")
    st.stop()

groups = GROUPS_ALL if kelompok == "Semua" else [kelompok]
periods_sel = [p for p in PERIODES if p in periode]
produk_sel = [x for x in PRODUKS if x in produk]

fdf = raw_df[
    raw_df["Periode"].isin(periods_sel)
    & raw_df["Kelompok"].isin(groups)
    & raw_df["Produk"].isin(produk_sel)
].copy()

if fdf.empty:
    st.warning("Tidak ada data untuk kombinasi filter ini.")
    st.stop()

monthly = aggregate(fdf, ["Periode", "Kelompok"])
prod_df = aggregate(fdf, ["Periode", "Kelompok", "Produk"])

# ============================================================
# KPI PERSENTASE
# ============================================================
st.subheader("Ringkasan Persentase terhadap Grand Total per Bulan")

for i, p in enumerate(periods_sel):
    st.markdown(f"#### {p}")
    for g in groups:
        r = pick(monthly, Periode=p, Kelompok=g)
        if r is None:
            continue
        prev = pick(monthly, Periode=periods_sel[i - 1], Kelompok=g) if i > 0 else None
        metric_row(r, prev, g, p, "")

st.caption(
    "Angka kecil = selisih terhadap bulan sebelumnya dalam poin persentase. "
    + ARAH_CAPTION
    + " Arahkan kursor ke tanda tanya (?) untuk melihat amount dan Grand Total pembaginya."
)

st.divider()

# ============================================================
# RINGKASAN PERSENTASE PER PRODUK PER BULAN
# ============================================================
st.subheader("Ringkasan Persentase per Produk per Bulan")

prod_tabs = st.tabs(produk_sel)
for tab, prod_name in zip(prod_tabs, produk_sel):
    with tab:
        for i, p in enumerate(periods_sel):
            st.markdown(f"#### {prod_name} - {p}")
            for g in groups:
                r = pick(prod_df, Periode=p, Kelompok=g, Produk=prod_name)
                if r is None:
                    continue
                prev = (pick(prod_df, Periode=periods_sel[i - 1], Kelompok=g, Produk=prod_name)
                        if i > 0 else None)
                metric_row(r, prev, g, p, prod_name)

st.caption(
    "Setiap produk dihitung terhadap Grand Total produk itu sendiri pada kelompok dan bulan yang sama. "
    "Klik tab produk untuk berpindah antar produk. " + ARAH_CAPTION
)

st.divider()

# ============================================================
# KPI AMOUNT
# ============================================================
st.subheader("Ringkasan Amount PKK_AW")

k = fdf[STATUSES + [GAB, "Grand Total"]].sum()
a1, a2, a3 = st.columns(3)
a1.metric("Grand Total", rupiah(k["Grand Total"]))
a2.metric(
    GAB, rupiah(k[GAB]),
    help=f"Pelunasan {rupiah(k['Pelunasan'])} + Pickup {rupiah(k['Pickup'])}"
)
a3.metric("Rollback", rupiah(k["Rollback"]))

b1, b2, _ = st.columns(3)
b1.metric("Rolling", rupiah(k["Rolling"]))
b2.metric("Settle", rupiah(k["Settle"]))

st.caption("Amount menggunakan nilai Sum of PKK_AW dalam Rupiah, sesuai filter periode, kelompok, dan produk.")

st.divider()

# ============================================================
# 1. PERSENTASE PER STATUS PER PERIODE
# ============================================================
st.subheader("1. Persentase (Pelunasan + Pickup), Rollback, Rolling & Settle per Periode")

long_period = monthly.melt(
    id_vars=["Periode", "Kelompok"],
    value_vars=[f"{s} %" for s in PCT_STATUSES],
    var_name="Status",
    value_name="Persentase"
)
long_period["Status"] = long_period["Status"].str.replace(" %", "", regex=False)

fig_period = px.bar(
    long_period, x="Periode", y="Persentase", color="Kelompok",
    facet_col="Status", barmode="group", text="Persentase",
    color_discrete_map=COLORS_GROUP,
    category_orders={"Status": PCT_STATUSES, "Kelompok": GROUPS_ALL, "Periode": PERIODES},
    labels={"Persentase": "Persentase (%)"}
)
fig_period.update_traces(texttemplate="%{text:.2f}%", textposition="outside", cliponaxis=False)
fig_period.update_yaxes(matches=None, showticklabels=True)
fig_period.update_layout(height=460)
st.plotly_chart(style_fig(fig_period, pct=True), use_container_width=True)
st.caption("Skala sumbu Y tiap status dibuat terpisah agar status dengan persentase kecil tetap terbaca.")

# ============================================================
# 2. PERSENTASE PER PRODUK
# ============================================================
st.subheader("2. Persentase per Produk")

status_prod = st.selectbox("Pilih status", PCT_STATUSES, index=0, key="status_prod")
col_prod = f"{status_prod} %"

fig_prod = px.bar(
    prod_df, x="Produk", y=col_prod, color="Kelompok",
    facet_col="Periode", barmode="group", text=col_prod,
    color_discrete_map=COLORS_GROUP,
    category_orders={"Produk": PRODUKS, "Kelompok": GROUPS_ALL, "Periode": PERIODES},
    labels={col_prod: f"{status_prod} (%)"}
)
fig_prod.update_traces(texttemplate="%{text:.2f}%", textposition="outside", cliponaxis=False)
fig_prod.update_layout(height=460)
st.plotly_chart(style_fig(fig_prod, pct=True), use_container_width=True)
st.caption(
    f"{status_prod}: " + ("higher is better (semakin tinggi semakin baik)."
                          if HIGHER_BETTER[status_prod]
                          else "lower is better (semakin rendah semakin baik).")
)

# ============================================================
# 3. TREND PER PRODUK
# ============================================================
st.subheader("3. Trend Persentase per Produk")

status_trend = st.selectbox("Pilih indikator", PCT_STATUSES, index=0, key="status_trend")
col_trend = f"{status_trend} %"

fig_trend = px.line(
    prod_df, x="Periode", y=col_trend, color="Produk", line_dash="Kelompok",
    markers=True,
    color_discrete_map=COLORS_PRODUK,
    category_orders={"Produk": PRODUKS, "Kelompok": GROUPS_ALL, "Periode": PERIODES},
    labels={col_trend: f"{status_trend} (%)"}
)
fig_trend.update_layout(height=440)
st.plotly_chart(style_fig(fig_trend, pct=True), use_container_width=True)
st.caption(
    "Bentuk garis membedakan kelompok (C0 / C1), warna membedakan produk. "
    + ("Higher is better." if HIGHER_BETTER[status_trend] else "Lower is better.")
)

# ============================================================
# 4. AMOUNT PER STATUS
# ============================================================
st.subheader("4. Komposisi Amount PKK_AW per Status")

amount_long = monthly.melt(
    id_vars=["Periode", "Kelompok"], value_vars=PCT_STATUSES,
    var_name="Status", value_name="Amount"
)

fig_amount = px.bar(
    amount_long, x="Periode", y="Amount", color="Status",
    facet_col="Kelompok", barmode="stack",
    color_discrete_sequence=COLORS_STATUS,
    category_orders={"Status": PCT_STATUSES, "Kelompok": GROUPS_ALL, "Periode": PERIODES},
    labels={"Amount": "PKK_AW (Rp)"}
)
st.plotly_chart(style_fig(fig_amount, amount=True), use_container_width=True)

# ============================================================
# 5. TABEL PERSENTASE
# ============================================================
st.subheader("5. Tabel Persentase per Produk")

pct_table = prod_df[["Periode", "Kelompok", "Produk"] + [f"{s} %" for s in PCT_STATUSES]].copy()
for s in PCT_STATUSES:
    pct_table[f"{s} %"] = pct_table[f"{s} %"].map(pct_fmt)
st.dataframe(pct_table, use_container_width=True, hide_index=True)

# ============================================================
# 6. DATA DETAIL (AMOUNT)
# ============================================================
st.subheader("6. Data Detail Amount")

amount_cols = STATUSES + [GAB, "Grand Total", "Pembagi"]
detail = fdf.sort_values(["Periode", "Kelompok", "Produk"])[["Periode", "Kelompok", "Produk"] + amount_cols].copy()
for c in amount_cols:
    detail[c] = detail[c].apply(rupiah)
amount_cfg = {c: st.column_config.TextColumn(c, width="medium") for c in amount_cols}
st.dataframe(detail, use_container_width=True, hide_index=True, column_config=amount_cfg)
st.caption(
    "Kolom Pembagi = Grand Total yang dipakai sebagai penyebut persentase. "
    "Pelunasan + Pickup = Pelunasan dan Pickup dijumlahkan."
)

# ============================================================
# 7. INSIGHT
# ============================================================
st.subheader("7. Insight")
st.caption(ARAH_CAPTION)

st.markdown("### Sorotan")
for g in groups:
    sub = monthly[monthly["Kelompok"] == g]
    if sub.empty:
        continue
    st.markdown(f"**{g}**")
    for s in PCT_STATUSES:
        col = f"{s} %"
        if HIGHER_BETTER[s]:
            best = sub.loc[sub[col].idxmax()]
            worst = sub.loc[sub[col].idxmin()]
        else:
            best = sub.loc[sub[col].idxmin()]
            worst = sub.loc[sub[col].idxmax()]
        if len(sub) == 1:
            st.write(f"• {s}: {pct_fmt(best[col])} pada **{best['Periode']}**.")
        else:
            st.write(
                f"• {s}: terbaik pada **{best['Periode']}** ({pct_fmt(best[col])}), "
                f"terburuk pada **{worst['Periode']}** ({pct_fmt(worst[col])})."
            )

st.markdown("### Per Periode")
for i, p in enumerate(periods_sel):
    st.markdown(f"#### {p}")
    for g in groups:
        r = pick(monthly, Periode=p, Kelompok=g)
        if r is None:
            continue
        ringkas = ", ".join(
            f"{s} {pct_fmt(r[f'{s} %'])} ({rupiah(r[s])})" for s in PCT_STATUSES
        )
        st.write(f"**{g}:** {ringkas}.")

        sub = prod_df[(prod_df["Periode"] == p) & (prod_df["Kelompok"] == g)]
        if sub.empty:
            continue
        top_gab = sub.loc[sub[f"{GAB} %"].idxmax()]
        if top_gab[f"{GAB} %"] > 0:
            st.write(
                f"• {GAB} tertinggi (baik) pada produk **{top_gab['Produk']}**: "
                f"{pct_fmt(top_gab[f'{GAB} %'])} ({rupiah(top_gab[GAB])})."
            )
        top_rb = sub.loc[sub["Rollback %"].idxmax()]
        if top_rb["Rollback %"] > 0:
            st.write(
                f"• Rollback tertinggi (baik) pada produk **{top_rb['Produk']}**: "
                f"{pct_fmt(top_rb['Rollback %'])} ({rupiah(top_rb['Rollback'])})."
            )
        top_rl = sub.loc[sub["Rolling %"].idxmax()]
        if top_rl["Rolling %"] > 0:
            st.write(
                f"• Rolling tertinggi (perlu perhatian) pada produk **{top_rl['Produk']}**: "
                f"{pct_fmt(top_rl['Rolling %'])} ({rupiah(top_rl['Rolling'])})."
            )

        if i > 0:
            prev_p = periods_sel[i - 1]
            pr = pick(monthly, Periode=prev_p, Kelompok=g)
            if pr is not None:
                parts = []
                for s in PCT_STATUSES:
                    d = r[f"{s} %"] - pr[f"{s} %"]
                    parts.append(f"{s} {kualitas(s, d)} ({delta_text(d)})")
                st.write(f"• Dibanding {prev_p}: " + ", ".join(parts) + ".")
