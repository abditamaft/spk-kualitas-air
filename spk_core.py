"""
spk_core.py
Modul inti SPK Perangkingan Kualitas Air Kolam/Tambak (Weighted Product & SAW).
Dipakai oleh app_spk.py (GUI) dan buat_laporan.py (generator laporan).
"""
import glob
import os

import numpy as np
import pandas as pd
from scipy.stats import kendalltau, spearmanr

# ---------------------------------------------------------------
# 1. DEFINISI KRITERIA
#    mode "skor" : nilai asli dikonversi ke skala 1-5 (kriteria berentang optimum)
#    mode "raw"  : nilai asli (hasil rata-rata) dipakai langsung
# ---------------------------------------------------------------
KRITERIA = [
    {"kode": "C1", "nama": "Suhu Air", "satuan": "°C", "tipe": "Benefit", "mode": "skor", "kepentingan": 4},
    {"kode": "C2", "nama": "Kekeruhan", "satuan": "NTU", "tipe": "Cost", "mode": "raw", "kepentingan": 3},
    {"kode": "C3", "nama": "pH Air", "satuan": "", "tipe": "Benefit", "mode": "skor", "kepentingan": 5},
    {"kode": "C4", "nama": "Oksigen Terlarut (DO)", "satuan": "mg/L", "tipe": "Benefit", "mode": "raw", "kepentingan": 4},
]

# Aturan konversi skor (rentang optimum -> skor 5, tiap "langkah" jarak menurunkan 1 skor)
ATURAN_SKOR = {
    "C1": {"inti": (28.0, 30.0), "langkah": 2.0},
    "C3": {"inti": (7.0, 8.0), "langkah": 0.5},
}


def skor_rentang(x, inti, langkah):
    lo, hi = inti
    jarak = 0.0 if lo <= x <= hi else (lo - x if x < lo else x - hi)
    if jarak == 0:
        return 5
    s = 5 - int(np.ceil(jarak / langkah))
    return max(1, s)


# ---------------------------------------------------------------
# 2. DATA
# ---------------------------------------------------------------
def data_demo():
    """DATA ILUSTRASI (BUKAN data Kaggle) untuk menguji alur perhitungan."""
    return pd.DataFrame(
        {
            "Suhu Air": [28.4, 30.6, 27.1, 33.2, 29.3, 25.2],
            "Kekeruhan": [18.2, 32.5, 12.4, 45.0, 9.8, 27.9],
            "pH Air": [7.2, 6.8, 7.6, 8.7, 7.9, 6.3],
            "Oksigen Terlarut (DO)": [6.1, 4.7, 5.4, 3.6, 6.8, 5.1],
        },
        index=[f"Kolam {i}" for i in range(1, 7)],
    )


def _cari_kolom(df, kunci, persis=False):
    for c in df.columns:
        nama = str(c).strip().lower()
        if (persis and nama == kunci) or (not persis and kunci in nama):
            return c
    return None


def agregasi_frames(frames, n=None):
    """frames: dict {nama_kolam: DataFrame mentah Kaggle}. Hasil: rata-rata per kolam."""
    baris, ref = {}, {}
    for nama, df in frames.items():
        c_suhu = _cari_kolom(df, "temp")
        c_keruh = _cari_kolom(df, "turb")
        c_ph = next((c for c in df.columns if str(c).strip().lower().startswith("ph")), None)
        c_do = _cari_kolom(df, "oxygen") or _cari_kolom(df, "dissolved")
        if None in (c_suhu, c_keruh, c_ph, c_do):
            continue
        sub = df[[c_suhu, c_keruh, c_ph, c_do]].apply(pd.to_numeric, errors="coerce").dropna()
        # pembersihan sederhana: buang nilai tidak fisik
        sub = sub[(sub[c_ph].between(0, 14)) & (sub[c_suhu] > 0) & (sub[c_keruh] >= 0) & (sub[c_do] >= 0)]
        if sub.empty:
            continue
        baris[nama] = [sub[c_suhu].mean(), sub[c_keruh].mean(), sub[c_ph].mean(), sub[c_do].mean()]
        c_w = _cari_kolom(df, "weight")
        if c_w is not None:
            w = pd.to_numeric(df[c_w], errors="coerce").dropna()
            if len(w):
                ref[nama] = w.mean()
    out = pd.DataFrame.from_dict(
        baris, orient="index", columns=[k["nama"] for k in KRITERIA]
    )
    if n:
        out = out.iloc[:n]
    refs = pd.Series(ref).reindex(out.index) if ref else None
    return out, refs


def load_folder(folder, n=None):
    files = sorted(glob.glob(os.path.join(folder, "*.csv")))
    frames = {}
    for i, f in enumerate(files, 1):
        frames[f"Kolam {i}"] = pd.read_csv(f)
    return agregasi_frames(frames, n)


# ---------------------------------------------------------------
# 3. PERHITUNGAN
# ---------------------------------------------------------------
def konversi_matriks(raw, kriteria=KRITERIA):
    X = pd.DataFrame(index=raw.index)
    for k in kriteria:
        kol = raw[k["nama"]]
        if k["mode"] == "skor":
            a = ATURAN_SKOR[k["kode"]]
            X[k["kode"]] = kol.apply(lambda v: skor_rentang(v, a["inti"], a["langkah"]))
        else:
            X[k["kode"]] = kol.astype(float)
    return X


def bobot_normal(kriteria=KRITERIA):
    kep = pd.Series({k["kode"]: k["kepentingan"] for k in kriteria}, dtype=float)
    return kep / kep.sum()


def hitung_wp(X, W, kriteria=KRITERIA):
    tipe = {k["kode"]: k["tipe"] for k in kriteria}
    pangkat = pd.Series({c: (W[c] if tipe[c] == "Benefit" else -W[c]) for c in X.columns})
    S = (X.astype(float) ** pangkat).prod(axis=1)
    V = S / S.sum()
    return pangkat, S, V


def hitung_saw(X, W, kriteria=KRITERIA):
    tipe = {k["kode"]: k["tipe"] for k in kriteria}
    R = pd.DataFrame(index=X.index)
    for c in X.columns:
        if tipe[c] == "Benefit":
            R[c] = X[c] / X[c].max()
        else:
            R[c] = X[c].min() / X[c]
    V = (R * W).sum(axis=1)
    return R, V


def peringkat(V):
    return V.rank(ascending=False, method="min").astype(int)


def hitung_semua(raw, kriteria=KRITERIA, refs=None):
    X = konversi_matriks(raw, kriteria)
    W = bobot_normal(kriteria)
    pangkat, S, V_wp = hitung_wp(X, W, kriteria)
    R, V_saw = hitung_saw(X, W, kriteria)
    r_wp, r_saw = peringkat(V_wp), peringkat(V_saw)
    rho, p_rho = spearmanr(V_wp, V_saw)
    tau, p_tau = kendalltau(V_wp, V_saw)
    hasil = {
        "raw": raw, "X": X, "W": W, "pangkat": pangkat, "S": S, "V_wp": V_wp, "R": R,
        "V_saw": V_saw, "rank_wp": r_wp, "rank_saw": r_saw,
        "rho": float(rho), "tau": float(tau),
        "sama_top1": r_wp.idxmin() == r_saw.idxmin(),
        "selisih_rank": int((r_wp - r_saw).abs().sum()),
    }
    hasil["sensitivitas"] = sensitivitas(X, kriteria, r_wp)
    if refs is not None and refs.notna().sum() >= 3:
        ok = refs.notna()
        hasil["valid_rho_wp"] = float(spearmanr(V_wp[ok], refs[ok])[0])
        hasil["valid_rho_saw"] = float(spearmanr(V_saw[ok], refs[ok])[0])
        hasil["refs"] = refs
    return hasil


def sensitivitas(X, kriteria, rank_dasar):
    """Ubah tingkat kepentingan tiap kriteria +1 / -1 (batas 1..5), hitung ulang peringkat."""
    baris = [("Dasar", None, 0)]
    for k in kriteria:
        baris.append((f"{k['kode']} +1", k["kode"], +1))
        baris.append((f"{k['kode']} -1", k["kode"], -1))
    out = []
    for nama, kode, d in baris:
        kr = [dict(k) for k in kriteria]
        if kode:
            for k in kr:
                if k["kode"] == kode:
                    k["kepentingan"] = int(min(5, max(1, k["kepentingan"] + d)))
        W = bobot_normal(kr)
        _, _, Vw = hitung_wp(X, W, kr)
        _, Vs = hitung_saw(X, W, kr)
        rw, rs = peringkat(Vw), peringkat(Vs)
        out.append({
            "Skenario": nama,
            "Peringkat 1 WP": rw.idxmin(),
            "Peringkat 1 SAW": rs.idxmin(),
            "Spearman WP vs dasar": float(spearmanr(rw, rank_dasar)[0]) if rw.nunique() > 1 else 1.0,
        })
    return pd.DataFrame(out)


def uraian_skor():
    """Tabel aturan konversi skor untuk laporan."""
    rows = []
    for kode, a in ATURAN_SKOR.items():
        lo, hi = a["inti"]
        st = a["langkah"]
        rows.append([kode, f"{lo:g} – {hi:g}", f"{lo - st:g} – <{lo:g} atau >{hi:g} – {hi + st:g}",
                     f"{lo - 2 * st:g} – <{lo - st:g} atau >{hi + st:g} – {hi + 2 * st:g}",
                     f"{lo - 3 * st:g} – <{lo - 2 * st:g} atau >{hi + 2 * st:g} – {hi + 3 * st:g}",
                     "di luar rentang tersebut"])
    return rows


if __name__ == "__main__":
    raw = data_demo()
    h = hitung_semua(raw)
    pd.set_option("display.width", 200)
    print(h["raw"], "\n", h["X"], "\n", h["W"], "\n", h["S"], "\n", h["V_wp"], "\n", h["R"], "\n", h["V_saw"])
    print(h["rank_wp"].tolist(), h["rank_saw"].tolist(), h["rho"], h["tau"])
    print(h["sensitivitas"])
