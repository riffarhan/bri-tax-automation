"""Write the PPh engine results to PSIAP-ready Excel files."""
import io

import pandas as pd


def template_pph_bytes(result, sheet_name="Template") -> bytes:
    """The upload-ready PSIAP bukti-potong template (one stream)."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xl:
        result.template.to_excel(xl, sheet_name=sheet_name, index=False)
    return buf.getvalue()


def rekon_pph_bytes(rekon_df) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xl:
        rekon_df.to_excel(xl, sheet_name="REKON", index=False)
    return buf.getvalue()


def workbook_pph_bytes(results: dict, rekons: dict = None) -> bytes:
    """One review workbook: a Template sheet per stream + exceptions + rekon.
    `results` maps stream label (PPH 22 / PPH 23 / PPH 4A2 / SIPOBRI) -> PphResult."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xl:
        exc_frames = []
        for label, res in results.items():
            res.template.to_excel(xl, sheet_name=label[:31], index=False)
            if len(res.exceptions):
                e = res.exceptions.copy()
                e.insert(0, "Stream", label)
                exc_frames.append(e)
        (pd.concat(exc_frames, ignore_index=True) if exc_frames
         else pd.DataFrame(columns=["Stream", "Baris", "Vendor", "Jenis", "Keterangan"])
         ).to_excel(xl, sheet_name="Pengecualian", index=False)
        for label, rk in (rekons or {}).items():
            rk.to_excel(xl, sheet_name=f"REKON {label}"[:31], index=False)
    return buf.getvalue()
