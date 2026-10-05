import copy

import pandas as pd
import streamlit as st

import spk_core as core

st.set_page_config(page_title="SPK Kualitas Air Kolam", page_icon="💧", layout="wide")

MENU = ["1. Beranda", "2. Input Data", "3. Kriteria dan Bobot", "4. Hasil WP", "5. Hasil SAW",
        "6. Perbandingan dan Validasi"]

# ---------- state ----------
if "raw" not in st.session_state:
    st.session_state.raw = None
    st.session_state.refs = None
if "kriteria" not in st.session_state:
    st.session_state.kriteria = copy.deepcopy(core.KRITERIA)


def hasil():
    if st.session_state.raw is None:
        st.warning("Data belum dimuat. Buka form **2. Input Data** terlebih dahulu.")
        st.stop()
    return core.hitung_semua(st.session_state.raw, st.session_state.kriteria, st.session_state.refs)


st.sidebar.title("💧 SPK Kualitas Air")
halaman = st.sidebar.radio("Menu", MENU)

# ---------- 1. Beranda ----------
if halaman == MENU[0]:
    st.title("SPK Perangkingan Kualitas Air Kolam")
    st.write("Aplikasi ini mengurutkan kolam budidaya berdasarkan kualitas air menggunakan metode "
             "**Weighted Product (WP)** dan **Simple Additive Weighting (SAW)**.")
    st.markdown("""
**Langkah penggunaan**
1. **Input Data** – unggah berkas CSV Kaggle (satu berkas = satu kolam) atau pakai data contoh.
2. **Kriteria dan Bobot** – atur tingkat kepentingan tiap kriteria (1–5).
3. **Hasil WP** dan 4. **Hasil SAW** – telusuri langkah perhitungan dan peringkat.
5. **Perbandingan dan Validasi** – korelasi peringkat dan analisis sensitivitas.
""")

# ---------- 2. Input Data ----------
elif halaman == MENU[1]:
    st.header("Input Data")
    sumber = st.radio("Sumber data", ["Unggah CSV Kaggle", "Data contoh (ilustrasi)"], horizontal=True)
    if sumber == "Unggah CSV Kaggle":
        files = st.file_uploader("Pilih beberapa berkas CSV (satu berkas = satu kolam)", type="csv",
                                 accept_multiple_files=True)
        n = st.number_input("Jumlah kolam yang dipakai (0 = semua)", 0, 50, 0)
        if files and st.button("Muat data"):
            frames = {f"Kolam {i}": pd.read_csv(f) for i, f in enumerate(sorted(files, key=lambda x: x.name), 1)}
            raw, refs = core.agregasi_frames(frames, int(n) or None)
            if raw.empty:
                st.error("Kolom suhu/kekeruhan/pH/oksigen tidak terdeteksi. Periksa nama kolom CSV.")
            else:
                st.session_state.raw, st.session_state.refs = raw, refs
                st.success(f"{len(raw)} kolam berhasil dimuat.")
    else:
        if st.button("Muat data contoh"):
            st.session_state.raw, st.session_state.refs = core.data_demo(), None
            st.info("Data contoh adalah data ILUSTRASI, bukan data Kaggle.")
    if st.session_state.raw is not None:
        st.subheader("Matriks keputusan awal (rata-rata per kolam)")
        st.dataframe(st.session_state.raw.round(2), use_container_width=True)

# ---------- 3. Kriteria dan Bobot ----------
elif halaman == MENU[2]:
    st.header("Kriteria dan Bobot")
    for k in st.session_state.kriteria:
        k["kepentingan"] = st.slider(f"{k['kode']} – {k['nama']} ({k['tipe']})", 1, 5, int(k["kepentingan"]))
    W = core.bobot_normal(st.session_state.kriteria)
    tabel = pd.DataFrame({
        "Kriteria": [k["nama"] for k in st.session_state.kriteria],
        "Tipe": [k["tipe"] for k in st.session_state.kriteria],
        "Kepentingan": [k["kepentingan"] for k in st.session_state.kriteria],
        "Bobot normal": W.round(4).values,
    }, index=W.index)
    st.dataframe(tabel, use_container_width=True)
    st.metric("Jumlah bobot normal", f"{W.sum():.4f}")

# ---------- 4. Hasil WP ----------
elif halaman == MENU[3]:
    h = hasil()
    st.header("Hasil Weighted Product (WP)")
    st.subheader("Matriks keputusan X (setelah konversi skor)")
    st.dataframe(h["X"].round(2), use_container_width=True)
    st.subheader("Pangkat bobot")
    st.dataframe(h["pangkat"].round(4).to_frame("Pangkat").T, use_container_width=True)
    st.subheader("Vektor S, V, dan peringkat")
    out = pd.DataFrame({"S": h["S"].round(4), "V": h["V_wp"].round(4), "Peringkat": h["rank_wp"]})
    st.dataframe(out.sort_values("Peringkat"), use_container_width=True)
    st.bar_chart(h["V_wp"])

# ---------- 5. Hasil SAW ----------
elif halaman == MENU[4]:
    h = hasil()
    st.header("Hasil Simple Additive Weighting (SAW)")
    st.subheader("Matriks ternormalisasi R")
    st.dataframe(h["R"].round(4), use_container_width=True)
    st.subheader("Kontribusi w × r")
    st.dataframe((h["R"] * h["W"]).round(4), use_container_width=True)
    st.subheader("Nilai V dan peringkat")
    out = pd.DataFrame({"V": h["V_saw"].round(4), "Peringkat": h["rank_saw"]})
    st.dataframe(out.sort_values("Peringkat"), use_container_width=True)
    st.bar_chart(h["V_saw"])

# ---------- 6. Perbandingan dan Validasi ----------
else:
    h = hasil()
    st.header("Perbandingan dan Validasi")
    perb = pd.DataFrame({"V WP": h["V_wp"].round(4), "Rank WP": h["rank_wp"],
                         "V SAW": h["V_saw"].round(4), "Rank SAW": h["rank_saw"]})
    st.dataframe(perb.sort_values("Rank WP"), use_container_width=True)
    c1, c2, c3 = st.columns(3)
    c1.metric("Spearman ρ", f"{h['rho']:.4f}")
    c2.metric("Kendall τ", f"{h['tau']:.4f}")
    c3.metric("Peringkat 1 sama?", "Ya" if h["sama_top1"] else "Tidak")
    st.subheader("Analisis sensitivitas bobot")
    st.dataframe(h["sensitivitas"].round(4), use_container_width=True)
    if "valid_rho_wp" in h:
        st.subheader("Validasi eksternal (rata-rata berat ikan)")
        st.write(f"Spearman V(WP) vs berat ikan: **{h['valid_rho_wp']:.4f}**  |  "
                 f"V(SAW) vs berat ikan: **{h['valid_rho_saw']:.4f}**")
    else:
        st.caption("Kolom berat ikan tidak tersedia, validasi eksternal dilewati.")