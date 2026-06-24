"""
PPN WAPU → PSIAP — Streamlit pilot (RO Palembang).

Flow: upload SAP → grid pre-fills from SAP → Salsa pastes/fills the Coretax
masa + status → review flagged rows → download the PSIAP template.

Run locally / internal server (NOT public cloud — bank-sensitive data):
    streamlit run app.py
"""
import streamlit as st
import pandas as pd

from engine import (Config, read_sap, build_doc_index, normalize_coretax,
                    coretax_seed_from_sap, read_coretax, reconcile, BULAN_ID)
from writer import fm_import_bytes, workbook_bytes

st.set_page_config(page_title="PPN WAPU → PSIAP", page_icon="🧾", layout="wide")
st.title("🧾 PPN WAPU → Template Impor PSIAP")
st.caption("Pilot RO Palembang · diproses lokal, data tidak dikirim ke mana pun")

# ---------------------------------------------------------------- sidebar: masa + SAP
with st.sidebar:
    st.header("① Masa & SAP")
    ro = st.text_input("RO", "PALEMBANG")
    masa = st.selectbox("Masa setor (bulan)", list(BULAN_ID), index=3,
                        format_func=lambda m: f"{m:02d} — {BULAN_ID[m]}")
    tahun = st.number_input("Tahun", 2024, 2030, 2026)
    sap_file = st.file_uploader("Extract SAP PPN WAPU (.xlsx)", type=["xlsx"])
    st.caption("Sheet1 + BULAN SBLMNYA/BERIKUTNYA dipakai untuk doc number.")

cfg = Config(ro_name=ro.strip().upper(), masa=int(masa), tahun=int(tahun))

if not sap_file:
    st.info("⬅️ Mulai dengan upload file SAP. Grid Coretax akan otomatis terisi dari SAP.")
    st.stop()

sap = read_sap(sap_file)
by_faktur, by_amt = build_doc_index(sap_file)

# ---------------------------------------------------------------- step 2: Coretax data
st.subheader("② Data Coretax (Pajak Masukan)")
mode = st.radio("Sumber data Coretax", ["Ketik / paste di grid", "Upload file"],
                horizontal=True, label_visibility="collapsed")

coretax = None
if mode == "Upload file":
    ct_file = st.file_uploader("File Coretax (.xlsx)", type=["xlsx"], key="ct")
    ct_sheet = st.text_input("Nama sheet", "Faktur Masukan_1")
    if ct_file:
        coretax = read_coretax(ct_file, ct_sheet)
else:
    st.caption("Grid sudah terisi dari SAP — Salsa tinggal isi kolom **masa** "
               "(dari Coretax) & betulkan **status** kalau belum approved. "
               "Faktur yang hanya ada di Coretax: tambah baris di bawah.")
    seed = coretax_seed_from_sap(sap, cfg)
    edited = st.data_editor(
        seed, num_rows="dynamic", use_container_width=True, height=340,
        column_config={
            "nomor_faktur": st.column_config.TextColumn("Nomor Faktur", width="medium"),
            "npwp_penjual": st.column_config.TextColumn("NPWP Penjual"),
            "nama_penjual": st.column_config.TextColumn("Nama Vendor"),
            "masa": st.column_config.NumberColumn("Masa ⬅️Coretax", min_value=1, max_value=12, step=1),
            "tahun": st.column_config.NumberColumn("Tahun", step=1, format="%d"),
            "dpp": st.column_config.NumberColumn("DPP", format="%.0f"),
            "ppn": st.column_config.NumberColumn("PPN", format="%.0f"),
            "status": st.column_config.SelectboxColumn("Status", options=["approved", "not approved"]),
            "konfirmasi": st.column_config.SelectboxColumn("Konfirmasi", options=["uncredited", "credited"]),
        })
    coretax = normalize_coretax(edited)

if coretax is None or len(coretax) == 0:
    st.warning("Belum ada data Coretax.")
    st.stop()

missing_masa = int(coretax["masa"].isna().sum()) if "masa" in coretax else 0
if missing_masa:
    st.warning(f"⏳ {missing_masa} baris belum ada **masa**-nya — isi dulu dari Coretax "
               "supaya MASA_PAJAK template benar.")

# ---------------------------------------------------------------- step 3: process
res = reconcile(coretax, sap, cfg, by_faktur, by_amt)
s = res.stats

st.subheader("③ Hasil")
c1, c2, c3, c4 = st.columns(4)
c1.metric("Baris template", s["fm_import_rows"])
c2.metric("Doc number terisi", f"{s['doc_filled']}/{s['fm_import_rows']}")
c3.metric("Perlu direview", s["exceptions"])
c4.metric("SAP INVOICE", s["sap_invoice_rows"])

tab1, tab2 = st.tabs([f"⚠️ Pengecualian ({s['exceptions']})", "✅ Template FM-Import"])
with tab1:
    if len(res.exceptions):
        st.caption("Hanya baris ini yang perlu mata manusia — sisanya sudah otomatis.")
        for jenis, grp in res.exceptions.groupby("JENIS"):
            with st.expander(f"{jenis} — {len(grp)} baris", expanded=jenis.startswith("DOC")):
                st.dataframe(grp.drop(columns=["JENIS"]), use_container_width=True, hide_index=True)
    else:
        st.success("Tidak ada pengecualian — semua bersih ✨")
with tab2:
    st.dataframe(res.fm_import, use_container_width=True, hide_index=True)

st.divider()
fname = f"Template Impor PPN WAPU PSIAP RO {cfg.ro_name} {BULAN_ID[cfg.masa]} {cfg.tahun}.xlsx"
d1, d2 = st.columns(2)
d1.download_button("⬇️ Template PSIAP (siap upload)", fm_import_bytes(res), file_name=fname,
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                   use_container_width=True)
d2.download_button("⬇️ Workbook review (template + pengecualian)", workbook_bytes(res),
                   file_name=f"Review PPN WAPU {cfg.ro_name} {BULAN_ID[cfg.masa]} {cfg.tahun}.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                   use_container_width=True)
