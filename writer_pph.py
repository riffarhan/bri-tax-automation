"""Write the PPh engine results to PSIAP-ready Excel files."""
import io

import pandas as pd


def _df(obj):
    """Accept a PphResult or a plain (possibly grid-edited) DataFrame."""
    return obj.template if hasattr(obj, "template") else obj


def template_pph_bytes(result_or_df, sheet_name="Template") -> bytes:
    """The upload-ready PSIAP bukti-potong template (one stream)."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xl:
        _df(result_or_df).to_excel(xl, sheet_name=sheet_name, index=False)
    return buf.getvalue()


def rekon_pph_bytes(rekons: dict) -> bytes:
    """One workbook, a REKON sheet per pasal."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xl:
        for label, rk in rekons.items():
            rk.to_excel(xl, sheet_name=f"REKON {label}"[:31], index=False)
    return buf.getvalue()


def workbook_pph_bytes(templates: dict, exceptions: dict = None,
                       rekons: dict = None) -> bytes:
    """One review workbook: a Template sheet per stream + exceptions + rekon.
    `templates` maps stream label (PPH 22 / PPH 23 / ...) -> DataFrame/PphResult."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xl:
        exc_frames = []
        for label, obj in templates.items():
            _df(obj).to_excel(xl, sheet_name=label[:31], index=False)
        for label, e in (exceptions or {}).items():
            if e is not None and len(e):
                e = e.copy()
                e.insert(0, "Stream", label)
                exc_frames.append(e)
        (pd.concat(exc_frames, ignore_index=True) if exc_frames
         else pd.DataFrame(columns=["Stream", "Baris", "Vendor", "Jenis", "Keterangan"])
         ).to_excel(xl, sheet_name="Pengecualian", index=False)
        for label, rk in (rekons or {}).items():
            rk.to_excel(xl, sheet_name=f"REKON {label}"[:31], index=False)
    return buf.getvalue()
