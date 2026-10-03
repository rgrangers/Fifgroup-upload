"""Penyimpanan & pembacaan data bahan tagih (tanpa dependensi Streamlit)."""
import os
import re

import pandas as pd

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_PATH = os.path.join(BASE_DIR, "data_bahan_tagih.csv")

STATUSES = ["Pelunasan", "Pickup", "Rollback", "Rolling", "Settle"]
KEYS = ["Periode", "Kelompok", "Produk"]
COLS = KEYS + STATUSES

# ------------------------------------------------------------
# Simpan / muat
# ------------------------------------------------------------
def _clean(df):
    df = df.copy()
    for c in STATUSES:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0).round().astype("int64")
    for c in KEYS:
        df[c] = df[c].astype(str).str.strip()
    return df[COLS].reset_index(drop=True)


def save_data(df):
    _clean(df).to_csv(DATA_PATH, index=False)


def empty_df():
    return pd.DataFrame(columns=COLS)


def load_data():
    """Baca CSV data. Bila belum ada (belum pernah upload), data kosong."""
    if not os.path.exists(DATA_PATH):
        return empty_df()
    df = pd.read_csv(DATA_PATH)
    if df.empty:
        return empty_df()
    return _clean(df)


def period_sort_key(label):
    m = re.search(r"\d+", str(label))
    return int(m.group()) if m else 10**6


def ordered_periods(df):
    """Urutan periode: berdasarkan angka pada label (B7 < B8 < B10), lalu urutan masuk."""
    first_seen = {}
    for p in df["Periode"]:
        first_seen.setdefault(p, len(first_seen))
    return sorted(first_seen, key=lambda p: (period_sort_key(p), first_seen[p]))


def upsert_periods(df, new_rows):
    """Tambah periode baru; periode dengan label sama ditimpa."""
    periods = new_rows["Periode"].unique()
    kept = df[~df["Periode"].isin(periods)]
    return _clean(pd.concat([kept, new_rows[COLS]], ignore_index=True))


def delete_periods(df, periods):
    return _clean(df[~df["Periode"].isin(periods)])


def validate_backup(df):
    missing = [c for c in COLS if c not in df.columns]
    if missing:
        raise ValueError("Kolom CSV kurang: " + ", ".join(missing))
    return _clean(df)


# ------------------------------------------------------------
# Parser pivot Excel
# ------------------------------------------------------------
def _num(v):
    n = pd.to_numeric(v, errors="coerce")
    return 0 if pd.isna(n) else float(n)


def parse_pivot(raw):
    """raw: DataFrame sheet mentah (header=None).
    Format: baris label C0/C1 (kelompok) lalu produk di bawahnya, kolom status,
    sel kosong = 0. Mengembalikan (label_tebakan, df_baris, daftar_peringatan)."""
    warnings = []

    # label bulan, mis. "Bulan 7" -> "B7"
    label = ""
    for r in range(min(len(raw), 10)):
        for v in raw.iloc[r].tolist():
            if isinstance(v, str):
                m = re.fullmatch(r"\s*(?:bulan|b)\s*[-_]?\s*(\d{1,2})\s*", v, re.I)
                if m:
                    label = f"B{int(m.group(1))}"
                    break
        if label:
            break

    # baris header
    hr = None
    for r in range(len(raw)):
        v = raw.iloc[r, 0]
        if isinstance(v, str) and v.strip().lower() == "row labels":
            hr = r
            break
    if hr is None:
        raise ValueError("Baris header 'Row Labels' tidak ditemukan.")

    col_map, gt_col = {}, None
    for j, h in enumerate(raw.iloc[hr].tolist()):
        if j == 0 or not isinstance(h, str):
            continue
        name = h.strip().title()
        if name in STATUSES:
            col_map[name] = j
        elif name == "Grand Total":
            gt_col = j
    if not col_map:
        raise ValueError("Kolom status (PELUNASAN, PICKUP, ROLLBACK, ROLLING, SETTLE) tidak ditemukan.")
    absent = [s for s in STATUSES if s not in col_map]
    if absent:
        warnings.append("Kolom tidak ada di pivot (dianggap 0): " + ", ".join(absent))

    rows, group = [], None
    for r in range(hr + 1, len(raw)):
        first = raw.iloc[r, 0]
        if not isinstance(first, str) or not first.strip():
            continue
        name = first.strip()
        if name.lower() == "grand total":
            break
        if re.fullmatch(r"C\d+", name, re.I):
            group = name.upper()
            continue
        if group is None:
            continue
        vals = {s: _num(raw.iloc[r, col_map[s]]) if s in col_map else 0 for s in STATUSES}
        if gt_col is not None:
            gt = _num(raw.iloc[r, gt_col])
            if abs(sum(vals.values()) - gt) > 1:
                warnings.append(f"{group} {name}: jumlah status tidak sama dengan Grand Total di Excel.")
        rows.append({"Periode": label, "Kelompok": group, "Produk": name.upper(), **vals})

    if not rows:
        raise ValueError("Tidak ada baris produk yang terbaca (cek struktur C0/C1 dan produk).")
    return label, pd.DataFrame(rows, columns=COLS), warnings


def parse_workbook(file):
    """Baca semua sheet pada file Excel; satu sheet = satu bulan."""
    sheets = pd.read_excel(file, sheet_name=None, header=None)
    results = []
    for sheet, raw in sheets.items():
        try:
            label, rows, warns = parse_pivot(raw)
            results.append({"sheet": sheet, "label": label, "rows": rows,
                            "warnings": warns, "error": None})
        except ValueError as e:
            results.append({"sheet": sheet, "label": "", "rows": None,
                            "warnings": [], "error": str(e)})
    return results
