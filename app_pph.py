"""
PPh Unifikasi page for the Alfa app.

Flow: upload the SAP pulls (22/23/4 ayat 2, auto-detected from the file name) +
the SIPO/BRITAX per-uker exports + optional DIO manual Excel → review the grids
(fill flagged NPWP etc.) → download the PSIAP templates + rekon.
"""
import re

import pandas as pd
import streamlit as st

from engine import BULAN_ID
from engine_pph import (PphConfig, read_sap_pph, read_sipo, read_dio_pph,
                        build_template_sap, build_template_sipo, build_template_dio,
                        build_data_olah_pph, read_etb_pph, build_rekon_pph,
                        TEMPLATE_COLUMNS)
from grid import editable_grid
from writer_pph import (template_pph_bytes, workbook_pph_bytes, rekon_pph_bytes,
                        data_olah_pph_bytes)

XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
PASAL_LABEL = {"22": "PPH 22", "23": "PPH 23", "4A2": "PPH 4 AYAT 2"}


def detect_pasal(name: str):
    up = name.upper()
    if any(t in up for t in ["PPH 4 AYAT 2", "PPH 4-2", "4 AYAT 2", "PPH 4 "]):
        return "4A2"
    if "PPH 23" in up:
        return "23"
    if "PPH 22" in up:
        return "22"
    return None


def _postprocess(df: pd.DataFrame) -> pd.DataFrame:
    """After grid edits: NITKU penerima follows the (possibly corrected) NPWP,
    and rows added by hand get the constant columns filled from the first row."""
    df = df.copy()
    if len(df) == 0:
        return df
    consts = ["NPWP Pemotong", "Masa Pajak", "Tahun Pajak", "Jenis PPh",
              "Fasilitas Insentif", "Nomor Setifikat Insentif", "Tarif Fasilitas",
              "Jenis Dokumen Referensi", "NPWP Penandatangan"]
    first = df.iloc[0]
    for c in consts:
        df[c] = df[c].map(lambda v: first[c] if v in (None, "") or pd.isna(v) else v)
    def nitku(row):
        npwp = re.sub(r"\D", "", str(row["NPWP Penerima Penghasilan"] or ""))
        cur = str(row["NITKU Penerima Penghasilan (22 Digit)"] or "")
        if len(npwp) == 16 and not cur.startswith(npwp):
            return npwp + "000000"
        return cur if not pd.isna(row["NITKU Penerima Penghasilan (22 Digit)"]) else ""
    df["NITKU Penerima Penghasilan (22 Digit)"] = df.apply(nitku, axis=1)
    return df


DATE_COLS = ["Tanggal Dokumen Referensi", "Tanggal Pemotongan"]


def _grid(label, res, cfg):
    """Editable template grid; returns the post-processed DataFrame."""
    st.caption("Semua kolom bisa diedit — isi **NPWP Penerima** yang di-flag "
               "(0000/kosong) dari konfirmasi uker + Coretax; NITKU penerima "
               "mengikuti otomatis. Baris manual bisa ditambah di bawah.")
    df = res.template.copy()
    for c in DATE_COLS:
        # uniform dtype: SAP rows carry datetimes, DIO/added rows empty strings —
        # mixed object columns break the Arrow serializer behind data_editor
        df[c] = pd.to_datetime(df[c], errors="coerce")
    edited = editable_grid(
        df, height=320,
        # key ties the editor's edit-state to THIS dataset — without the
        # ro/masa/rowcount suffix, edits made before switching RO/masa/files
        # would silently re-apply to the new rows
        key=f"grid_{label}_{cfg.ro_name}_{cfg.masa}_{cfg.tahun}_{len(df)}",
        column_config={
            "NPWP Penerima Penghasilan": st.column_config.TextColumn(
                "NPWP Penerima Penghasilan", width="medium",
                help="16 digit. Kalau kosong/0000 → konfirmasi ke uker, validasi Coretax."),
            "Penghasilan Bruto": st.column_config.NumberColumn(
                "Penghasilan Bruto", format="%.0f"),
            **{c: st.column_config.DateColumn(c) for c in DATE_COLS},
        })
    return _postprocess(edited)


def render(ro: str, masa: int, tahun: int, app_name: str = "Alfa"):
    cfg = PphConfig(ro_name=ro.strip().upper(), masa=int(masa), tahun=int(tahun))
    bulan = BULAN_ID[cfg.masa]

    with st.sidebar:
        st.header("2 · PPh source files")
        sap_files = st.file_uploader(
            "SAP PPh (22 / 23 / 4 ayat 2) — boleh beberapa file",
            type=["xlsx"], accept_multiple_files=True, key="pph_sap",
            help="Pasal terdeteksi dari nama file (PPH 22 / PPH 23 / PPH 4 AYAT 2). "
                 "Semua sheet tarikan (RO/KANINS/SENDIK) digabung otomatis.")
        sipo_files = st.file_uploader(
            "SIPO / BRITAX per-uker (.xls) — pilih semua sekaligus",
            type=["xls"], accept_multiple_files=True, key="pph_sipo",
            help="File tarikan BRITAX per kode uker (128.xls, 240.xls, …). "
                 "Duplikat re-download (152(1).xls) dibuang otomatis.")
        dio_files = st.file_uploader(
            "Manual dari DIO (.xlsx) — opsional",
            type=["xlsx"], accept_multiple_files=True, key="pph_dio",
            help="Excel kiriman DIO (Pajak_PPH_23_Unifikasi_Manual_…). Nomor & "
                 "tanggal dokumen tetap diisi dari PDF bukti fisik.")
        etb_file = st.file_uploader(
            "ETB PPh Unifikasi (.xlsx) — untuk rekon", type=["xlsx"], key="pph_etb")

    if not (sap_files or sipo_files or dio_files):
        st.info("⬅️ Upload the SAP PPh pulls (and/or the SIPO per-uker files) to start. "
                "Each pasal becomes its own upload-ready PSIAP template.")
        return

    # ---------------------------------------------------------------- build streams
    results = {}          # label -> PphResult
    sap_raw = {}          # label -> (sap df, pasal) for the Data Olah download
    for f in sap_files or []:
        pasal = detect_pasal(f.name)
        if pasal is None:
            pasal = st.selectbox(f"Pasal untuk «{f.name}» (tidak terdeteksi)",
                                 ["22", "23", "4A2"], key=f"pasal_{f.name}")
        try:
            sap = read_sap_pph(f)
            if sap.empty:
                st.warning(f"«{f.name}»: tidak ada sheet tarikan SAP yang dikenali.")
                continue
            label = PASAL_LABEL[pasal]
            if label in sap_raw:
                sap = pd.concat([sap_raw[label][0], sap], ignore_index=True)
            sap_raw[label] = (sap, pasal)
            # sap is cumulative per pasal, so the rebuild covers every file
            results[label] = build_template_sap(sap, cfg, pasal)
        except Exception as e:  # noqa
            st.error(f"Gagal baca «{f.name}»: {e}")
    if sipo_files:
        try:
            sipo = read_sipo(sipo_files)
            if len(sipo):
                results["SIPOBRI"] = build_template_sipo(sipo, cfg)
        except Exception as e:  # noqa
            st.error(f"Gagal baca file SIPO: {e}")
    if dio_files:
        frames = []
        for f in dio_files:
            try:
                d = read_dio_pph(f)
                if len(d):
                    frames.append(d)
            except Exception as e:  # noqa
                st.error(f"Gagal baca «{f.name}»: {e}")
        if frames:
            results["PPH 23 MANUAL"] = build_template_dio(
                pd.concat(frames, ignore_index=True), cfg, "23")

    if not results:
        st.warning("Belum ada data yang terbaca.")
        return

    # ---------------------------------------------------------------- metrics
    st.subheader("3 · Result")
    n_exc = sum(len(r.exceptions) for r in results.values())
    cols = st.columns(len(results) + 1)
    for c, (label, r) in zip(cols, results.items()):
        c.metric(label, f"{len(r.template)} rows")
    cols[-1].metric("Needs review", n_exc)

    # ---------------------------------------------------------------- tabs
    labels = list(results)
    tabs = st.tabs([f"⚠️ Exceptions ({n_exc})"] + [f"📄 {l}" for l in labels]
                   + ["📊 Rekon"])
    with tabs[0]:
        if n_exc:
            st.caption("Only these rows need a human — everything else is automated.")
            for label, r in results.items():
                if not len(r.exceptions):
                    continue
                for jenis, grp in r.exceptions.groupby("Jenis"):
                    with st.expander(f"{label} · {jenis} — {len(grp)} row(s)",
                                     expanded="NPWP" in jenis):
                        st.dataframe(grp.drop(columns=["Jenis"]),
                                     use_container_width=True, hide_index=True)
        else:
            st.success("No exceptions — everything derived cleanly. ✨")

    edited = {}
    for tab, label in zip(tabs[1:-1], labels):
        with tab:
            edited[label] = _grid(label, results[label], cfg)

    rekons = {}
    with tabs[-1]:
        sap_labels = [l for l in labels if l in PASAL_LABEL.values()]
        if not etb_file:
            st.info("Upload the **ETB PPh Unifikasi** file in the sidebar to build "
                    "the rekon (Utang per uker vs PAJAK per uker, per pasal).")
        elif not sap_labels:
            st.info("Rekon dibangun dari tarikan SAP per pasal — upload file SAP "
                    "PPh dulu.")
        else:
            st.caption("Per pasal: **SELISIH = Utang (ETB) − PAJAK** per uker; 0 = "
                       "cocok. Angka PAJAK dari tarikan SAP (baris tanpa KOP ikut "
                       "dihitung — deposito/reward tetap nyata di buku besar).")
            for label in sap_labels:
                pasal = next(k for k, v in PASAL_LABEL.items() if v == label)
                try:
                    etb_file.seek(0)
                    utang = read_etb_pph(etb_file, cfg.ro_name, pasal)
                    if not utang:
                        st.warning(f"{label}: kolom Utang pasal ini tidak ketemu di ETB.")
                        continue
                    rk = build_rekon_pph(results[label].recon_rows, utang, cfg.ro_name)
                    rekons[label] = rk
                    n_off = int((rk["SELISIH"].fillna(0) != 0).sum())
                    with st.expander(f"{label} — {n_off} uker selisih ≠ 0",
                                     expanded=n_off > 0):
                        st.dataframe(rk, use_container_width=True, hide_index=True)
                except Exception as e:  # noqa
                    st.error(f"Gagal rekon {label}: {e}")
            if rekons:
                st.download_button(
                    "⬇️ Rekon PPh (.xlsx)", rekon_pph_bytes(rekons),
                    file_name=f"Rekon PPh {cfg.ro_name} {bulan} {cfg.tahun}.xlsx",
                    mime=XLSX)

    # ---------------------------------------------------------------- downloads
    st.divider()
    st.caption("Satu template siap upload per stream — nama file mengikuti pola "
               "kerja yang sudah ada.")
    fname = {
        "PPH 22": f"NEW TEMPLATE PSIAP PPH 22 RO {cfg.ro_name} {bulan} {cfg.tahun}.xlsx",
        "PPH 23": f"NEW TEMPLATE PSIAP PPH 23 RO {cfg.ro_name} {bulan} {cfg.tahun}.xlsx",
        "PPH 4 AYAT 2": f"NEW TEMPLATE PSIAP PPH 4 AYAT 2 RO {cfg.ro_name} {bulan} {cfg.tahun}.xlsx",
        "SIPOBRI": f"NEW TEMPLATE PSIAP SIPOBRI RO {cfg.ro_name} {bulan} {cfg.tahun}.xlsx",
        "PPH 23 MANUAL": f"NEW TEMPLATE PSIAP MANUAL PPH 23 RO {cfg.ro_name} {bulan} {cfg.tahun}.xlsx",
    }
    dcols = st.columns(min(len(edited), 3))
    for i, (label, df) in enumerate(edited.items()):
        dcols[i % len(dcols)].download_button(
            f"⬇️ Template {label}", template_pph_bytes(df),
            file_name=fname.get(label, f"Template {label}.xlsx"),
            mime=XLSX, use_container_width=True, key=f"dl_{label}")
    w1, w2 = st.columns(2)
    w1.download_button(
        "⬇️ Review workbook (semua stream + exceptions + rekon)",
        workbook_pph_bytes(edited, {l: r.exceptions for l, r in results.items()}, rekons),
        file_name=f"Review PPh {cfg.ro_name} {bulan} {cfg.tahun}.xlsx",
        mime=XLSX, use_container_width=True)
    if sap_raw:
        olahs = {label: build_data_olah_pph(sap, cfg, pasal)
                 for label, (sap, pasal) in sap_raw.items()}
        w2.download_button(
            "⬇️ Data Olah (olahan SAP per pasal + gap setoran)",
            data_olah_pph_bytes(olahs),
            file_name=f"DATA OLAH PPh {cfg.ro_name} {bulan} {cfg.tahun}.xlsx",
            mime=XLSX, use_container_width=True,
            help="Semua baris tarikan SAP + uker/NITKU/referensi + "
                 "PPh hitung (DPP × tarif) vs setoran SAP per baris.")
