"""
PPN WAPU → PSIAP — Streamlit app.

Flow: upload SAP → the grid pre-fills from SAP → fill the Coretax masa + status
→ review flagged exceptions → download the PSIAP import template.

Run locally / on an internal server (NOT public cloud — data is bank-sensitive):
    streamlit run app.py
"""
import streamlit as st
import pandas as pd

from engine import (Config, read_sap, build_doc_index, normalize_coretax,
                    coretax_seed_from_sap, read_coretax, reconcile, BULAN_ID)
from writer import fm_import_bytes, workbook_bytes

# ---- branding (name combines Alkaina + Farhan; one line to swap) ------------
APP_NAME = "ALFA"
APP_TAGLINE = "PPN WAPU reconciliation & PSIAP export"
MONTHS_EN = {1: "January", 2: "February", 3: "March", 4: "April", 5: "May",
             6: "June", 7: "July", 8: "August", 9: "September", 10: "October",
             11: "November", 12: "December"}

st.set_page_config(page_title=APP_NAME, page_icon="🧾", layout="wide")
st.title(APP_NAME)
st.caption(f"{APP_TAGLINE} · runs locally — no data leaves this machine")

# ---------------------------------------------------------------- sidebar: period + SAP
with st.sidebar:
    st.header("1 · Period & source")
    ro = st.text_input("Regional Office", "PALEMBANG")
    masa = st.selectbox("Tax period (month)", list(BULAN_ID), index=3,
                        format_func=lambda m: f"{m:02d} — {MONTHS_EN[m]}")
    tahun = st.number_input("Year", 2024, 2030, 2026)
    sap_file = st.file_uploader("SAP PPN WAPU extract (.xlsx)", type=["xlsx"])
    st.caption("Current and adjacent-month sheets are detected automatically "
               "for document-number matching.")

cfg = Config(ro_name=ro.strip().upper(), masa=int(masa), tahun=int(tahun))

if not sap_file:
    st.info("⬅️ Start by uploading the SAP extract. The Coretax grid will "
            "pre-fill from it automatically.")
    st.stop()

sap = read_sap(sap_file)
by_faktur, by_amt = build_doc_index(sap_file)

# ---------------------------------------------------------------- step 2: Coretax data
st.subheader("2 · Coretax data (Pajak Masukan)")
mode = st.radio("Coretax source", ["Enter / paste in grid", "Upload file"],
                horizontal=True, label_visibility="collapsed")

coretax = None
if mode == "Upload file":
    ct_file = st.file_uploader("Coretax file (.xlsx)", type=["xlsx"], key="ct")
    ct_sheet = st.text_input("Sheet name", "Faktur Masukan_1")
    if ct_file:
        coretax = read_coretax(ct_file, ct_sheet)
else:
    st.caption("The grid is pre-filled from SAP. Fill in the **Masa** column "
               "from Coretax and correct the **Status** for any faktur that "
               "isn't approved. For fakturs that exist only in Coretax, add a "
               "row at the bottom.")
    seed = coretax_seed_from_sap(sap, cfg)
    edited = st.data_editor(
        seed, num_rows="dynamic", use_container_width=True, height=340,
        column_config={
            "nomor_faktur": st.column_config.TextColumn("Nomor Faktur", width="medium"),
            "npwp_penjual": st.column_config.TextColumn("NPWP Penjual"),
            "nama_penjual": st.column_config.TextColumn("Nama Vendor"),
            "masa": st.column_config.NumberColumn("Masa (dari Coretax)", min_value=1, max_value=12, step=1),
            "tahun": st.column_config.NumberColumn("Tahun", step=1, format="%d"),
            "dpp": st.column_config.NumberColumn("DPP", format="%.0f"),
            "ppn": st.column_config.NumberColumn("PPN", format="%.0f"),
            "status": st.column_config.SelectboxColumn("Status", options=["approved", "not approved"]),
            "konfirmasi": st.column_config.SelectboxColumn("Konfirmasi", options=["uncredited", "credited"]),
        })
    coretax = normalize_coretax(edited)

    # --- audit trail: what was changed from the SAP-seeded values? ---
    def _changed(col, old, new):
        if pd.isna(new) or str(old).strip().lower() in ("", "nan", "none"):
            return False
        if col == "dpp":                              # numeric: ignore formatting
            try:
                return round(float(old)) != round(float(new))
            except (TypeError, ValueError):
                pass
        return str(old).strip() != str(new).strip()   # text: ignore whitespace

    # compare by ROW POSITION (faktur isn't unique — blank/zero fakturs exist)
    base = seed.reset_index(drop=True)
    overrides, added = [], []
    for idx, row in edited.reset_index(drop=True).iterrows():
        if idx >= len(base):                          # rows appended in the editor
            if str(row["nomor_faktur"]).strip():
                added.append(str(row["nomor_faktur"]).strip())
            continue
        for col, lbl in [("npwp_penjual", "NPWP Penjual"), ("dpp", "DPP"), ("nama_penjual", "Nama Vendor")]:
            old, new = base.loc[idx, col], row[col]
            if _changed(col, old, new):
                overrides.append({"Nomor Faktur": str(row["nomor_faktur"]), "Kolom": lbl,
                                  "Dari SAP": old, "Diisi dari Coretax": new})
    if overrides or added:
        with st.expander(f"Changes from SAP — {len(overrides)} value(s) overridden, "
                         f"{len(added)} faktur(s) added", expanded=bool(overrides)):
            st.caption("Audit trail: values changed because Coretax differs from SAP.")
            if overrides:
                st.dataframe(pd.DataFrame(overrides), use_container_width=True, hide_index=True)
            if added:
                st.write("Coretax-only fakturs added:", ", ".join(added))

if coretax is None or len(coretax) == 0:
    st.warning("No Coretax data yet.")
    st.stop()

missing_masa = int(coretax["masa"].isna().sum()) if "masa" in coretax else 0
if missing_masa:
    st.warning(f"⏳ {missing_masa} row(s) have no **Masa** yet — fill it from "
               "Coretax so the template's tax period is correct.")

# ---------------------------------------------------------------- step 3: result
res = reconcile(coretax, sap, cfg, by_faktur, by_amt)
s = res.stats

st.subheader("3 · Result")
c1, c2, c3, c4 = st.columns(4)
c1.metric("Template rows", s["fm_import_rows"])
c2.metric("Doc numbers filled", f"{s['doc_filled']}/{s['fm_import_rows']}")
c3.metric("Needs review", s["exceptions"])
c4.metric("SAP invoices", s["sap_invoice_rows"])

tab1, tab2 = st.tabs([f"⚠️ Exceptions ({s['exceptions']})", "✅ FM-Import template"])
with tab1:
    if len(res.exceptions):
        st.caption("Only these rows need a human — everything else is automated.")
        for jenis, grp in res.exceptions.groupby("Jenis"):
            with st.expander(f"{jenis} — {len(grp)} row(s)",
                             expanded=jenis.startswith("Doc number") or jenis.startswith("Nomor faktur")):
                st.dataframe(grp.drop(columns=["Jenis"]), use_container_width=True, hide_index=True)
    else:
        st.success("No exceptions — everything reconciles. ✨")
with tab2:
    st.dataframe(res.fm_import, use_container_width=True, hide_index=True)

st.divider()
fname = f"Template Impor PPN WAPU PSIAP RO {cfg.ro_name} {BULAN_ID[cfg.masa]} {cfg.tahun}.xlsx"
d1, d2 = st.columns(2)
d1.download_button("⬇️ PSIAP template (ready to upload)", fm_import_bytes(res), file_name=fname,
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                   use_container_width=True)
d2.download_button("⬇️ Review workbook (template + exceptions)", workbook_bytes(res),
                   file_name=f"Review PPN WAPU {cfg.ro_name} {BULAN_ID[cfg.masa]} {cfg.tahun}.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                   use_container_width=True)

st.divider()
st.caption(f"{APP_NAME} · made by Farhan, for Alkaina 🤍")
