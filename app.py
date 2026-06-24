"""
PPN WAPU → PSIAP — Streamlit app.

Flow: upload SAP → the grid pre-fills from SAP → fill the Coretax masa + status
→ review flagged exceptions → download the PSIAP import template.

Run locally / on an internal server (NOT public cloud — data is bank-sensitive):
    streamlit run app.py
"""
import os
from pathlib import Path

import streamlit as st
import pandas as pd

from engine import (Config, read_sap, build_doc_index, normalize_coretax,
                    coretax_seed_from_sap, read_coretax, reconcile, BULAN_ID,
                    read_etb, build_cabang_index, build_rekon, ETB_SHEET_BY_RO,
                    read_uker_names)
from writer import fm_import_bytes, fm_import_full_bytes, workbook_bytes, rekon_bytes

# ---- branding (name combines Alkaina + Farhan) ------------------------------
APP_NAME = "Alfa"
APP_TAGLINE = "PPN WAPU reconciliation & PSIAP export"
LOGO = str(Path(__file__).parent / "assets" / "alfa-logo.png")
_HAS_LOGO = os.path.exists(LOGO)
MONTHS_EN = {1: "January", 2: "February", 3: "March", 4: "April", 5: "May",
             6: "June", 7: "July", 8: "August", 9: "September", 10: "October",
             11: "November", 12: "December"}

st.set_page_config(page_title=APP_NAME, page_icon=LOGO if _HAS_LOGO else "🧾", layout="wide")
if _HAS_LOGO:
    st.logo(LOGO)


def _password_ok() -> bool:
    """Optional gate. If an 'app_password' secret is set (e.g. on a web deploy),
    require it; with no secret (local run) the app is open."""
    try:
        pw = st.secrets["app_password"]
    except Exception:
        return True
    if st.session_state.get("auth_ok"):
        return True
    if _HAS_LOGO:
        st.image(LOGO, width=140)
    else:
        st.title(APP_NAME)
    with st.form("login"):
        entered = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Enter")
    if submitted and entered == pw:
        st.session_state["auth_ok"] = True
        st.rerun()
    elif submitted:
        st.error("Incorrect password.")
    return False


if not _password_ok():
    st.stop()

if _HAS_LOGO:
    st.image(LOGO, width=150)
else:
    st.title(APP_NAME)
st.caption(APP_TAGLINE)

# ---------------------------------------------------------------- sidebar: period + SAP
with st.sidebar:
    st.header("1 · Period & source")
    ro = st.selectbox("Regional Office", ["PALEMBANG", "YOGYAKARTA"],
                      help="Label PSIAP & sheet ETB mengikuti RO yang dipilih.")
    masa = st.selectbox("Masa setor (bulan pelaporan)", list(BULAN_ID), index=3,
                        format_func=lambda m: f"{m:02d} — {MONTHS_EN[m]}",
                        help="Masa Pajak yang dilaporkan/disetor bulan ini (masa "
                             "pelaporan). Beda dari kolom **Masa** di grid, yang "
                             "adalah masa pajak faktur dari Coretax.")
    tahun = st.number_input("Year", 2024, 2030, 2026)
    sap_file = st.file_uploader("SAP PPN WAPU extract (.xlsx)", type=["xlsx"])
    st.caption("Current and adjacent-month sheets are detected automatically "
               "for document-number matching.")
    etb_file = st.file_uploader("ETB file (.xlsx) — for the Rekon tab", type=["xlsx"], key="etb")
    etb_sheet = None
    if etb_file is not None:
        try:
            _etb_sheets = pd.ExcelFile(etb_file).sheet_names
            _default = ETB_SHEET_BY_RO.get(ro.strip().upper())
            etb_sheet = st.selectbox(
                "Sheet ETB", _etb_sheets,
                index=_etb_sheets.index(_default) if _default in _etb_sheets else 0,
                help="Pilih sheet yang berisi saldo ETB per uker (mis. PLG / YOG).")
        except Exception as e:  # noqa
            st.caption(f"Tidak bisa baca sheet ETB: {e}")

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
    st.caption("The grid is pre-filled from SAP. Fill in **Masa** (the faktur's "
               "*Masa Pajak*, from Coretax) and fix **Status** for any faktur "
               "that isn't *approved*. For fakturs that exist only in Coretax, "
               "add a row at the bottom.")
    seed = coretax_seed_from_sap(sap, cfg)
    edited = st.data_editor(
        seed, num_rows="dynamic", use_container_width=True, height=340,
        column_config={
            "nomor_faktur": st.column_config.TextColumn(
                "Nomor Faktur", width="medium", help="Nomor Faktur Pajak (e-faktur)."),
            "npwp_penjual": st.column_config.TextColumn(
                "NPWP Penjual", help="NPWP penjual / vendor."),
            "nama_penjual": st.column_config.TextColumn(
                "Nama Vendor", help="Nama penjual / vendor."),
            "masa": st.column_config.NumberColumn(
                "Masa", min_value=1, max_value=12, step=1,
                help="Masa Pajak Faktur — masa pajak dari faktur di Coretax. "
                     "Sering bulan sebelumnya, jadi beda dari masa setor di sidebar."),
            "tahun": st.column_config.NumberColumn(
                "Tahun", step=1, format="%d", help="Tahun Pajak faktur."),
            "dpp": st.column_config.NumberColumn(
                "DPP", format="%.0f", help="Dasar Pengenaan Pajak."),
            "ppn": st.column_config.NumberColumn(
                "PPN", format="%.0f",
                help="PPN faktur dari Coretax (dipakai di tab Rekon)."),
            "status": st.column_config.SelectboxColumn(
                "Status", options=["approved", "not approved"],
                help="Status validasi faktur di Coretax."),
            "konfirmasi": st.column_config.SelectboxColumn(
                "Konfirmasi", options=["uncredited", "credited"],
                help="Status pengkreditan faktur (konfirmasi)."),
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
    st.warning(f"⏳ {missing_masa} row(s) have no **Masa** yet — fill the faktur's "
               "*Masa Pajak* from Coretax (or switch the Coretax source to "
               "**Upload file**, which fills masa + PPN automatically) so "
               "`MASA_PAJAK` is correct.")

# ---------------------------------------------------------------- step 3: result
cbf, cba, cbs = build_cabang_index(sap_file)          # for kode uker + the Rekon
uker_names = {}
if etb_file:
    try:
        etb_file.seek(0)
        uker_names = read_uker_names(etb_file, sheet=etb_sheet, ro_name=cfg.ro_name)
    except Exception:
        uker_names = {}
res = reconcile(coretax, sap, cfg, by_faktur, by_amt, cbf, cba, cbs, uker_names)
s = res.stats

st.subheader("3 · Result")
c1, c2, c3, c4 = st.columns(4)
c1.metric("Template rows", s["fm_import_rows"])
c2.metric("Doc numbers filled", f"{s['doc_filled']}/{s['fm_import_rows']}")
c3.metric("Needs review", s["exceptions"])
c4.metric("SAP invoices", s["sap_invoice_rows"])

tab1, tab2, tab3 = st.tabs(
    [f"⚠️ Exceptions ({s['exceptions']})", "✅ FM-Import template", "📊 Rekon"])
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
with tab3:
    if not etb_file:
        st.info("Upload the **ETB file** in the sidebar to build the Rekon (PPN "
                "per uker × masa vs the ETB *saldo*). Note: the ETB file gives the "
                "balance, **not** the per-faktur PPN.")
    elif "masa" not in coretax or coretax["masa"].isna().all():
        st.warning("The Rekon groups PPN by uker × **masa**, but no **Masa** is "
                   "filled yet. Fill **Masa** from Coretax in the grid, or switch "
                   "the Coretax source above to **Upload file** (which has masa + PPN).")
    elif "ppn" not in coretax or coretax["ppn"].fillna(0).eq(0).all():
        st.warning("The **PPN** column is empty — the Rekon sums PPN per uker. "
                   "Fill PPN (from Coretax) in the grid, or use Upload-file mode.")
    else:
        try:
            etb_file.seek(0)
            etb = read_etb(etb_file, sheet=etb_sheet, ro_name=cfg.ro_name)
            rekon_df, reclass_flag = build_rekon(coretax, sap, etb, cfg, cbf, cba, cbs)
            st.caption("PPN per uker × masa, joined to the ETB balance. "
                       "**SELISIH** ≠ 0 → investigate. Branches are folded into "
                       "their parent uker via the reclass; anything that couldn't "
                       "be mapped is listed below.")
            st.dataframe(rekon_df, use_container_width=True, hide_index=True)
            if len(reclass_flag):
                with st.expander(f"⚠️ Perlu reclass / uker tidak terpetakan — {len(reclass_flag)} faktur",
                                 expanded=True):
                    st.dataframe(reclass_flag, use_container_width=True, hide_index=True)
            st.download_button(
                "⬇️ Rekon (.xlsx)", rekon_bytes(rekon_df, reclass_flag),
                file_name=f"Rekon PPN WAPU {cfg.ro_name} {BULAN_ID[cfg.masa]} {cfg.tahun}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:  # noqa
            st.error(f"Gagal membuat Rekon: {e}")

st.divider()
st.caption("Dua file: **template upload** (12 kolom, langsung ke PSIAP) dan "
           "**versi full** (+ kolom review: vendor, uker, DPP, tarif, pajak faktur "
           "vs SAP, selisih). Nama Uker terisi kalau file ETB di-upload.")
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
ro_masa = f"{cfg.ro_name} {BULAN_ID[cfg.masa]} {cfg.tahun}"
d1, d2 = st.columns(2)
d1.download_button("⬇️ Template PSIAP (siap upload)", fm_import_bytes(res),
                   file_name=f"Template Impor PPN WAPU PSIAP RO {ro_masa}.xlsx",
                   mime=XLSX, use_container_width=True)
d2.download_button("⬇️ Versi full (dengan kolom review)", fm_import_full_bytes(res),
                   file_name=f"FM-Import FULL PPN WAPU {ro_masa}.xlsx",
                   mime=XLSX, use_container_width=True)
st.download_button("⬇️ Review workbook (template + exceptions)", workbook_bytes(res),
                   file_name=f"Review PPN WAPU {ro_masa}.xlsx",
                   mime=XLSX, use_container_width=True)

st.divider()
st.caption(f"{APP_NAME} · made by Farhan, for Alkaina 🤍")
