"""Write the engine Result to PSIAP-ready Excel files."""
import io
import pandas as pd
from engine import FM_IMPORT_COLUMNS


def fm_import_bytes(result, sheet_name="FM - Import") -> bytes:
    """The upload-ready PSIAP template (just the FM-Import sheet)."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xl:
        result.fm_import.to_excel(xl, sheet_name=sheet_name, index=False)
    return buf.getvalue()


def workbook_bytes(result) -> bytes:
    """Full review workbook: template + exceptions + run stats."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xl:
        result.fm_import.to_excel(xl, sheet_name="FM - Import", index=False)
        (result.exceptions if len(result.exceptions)
         else pd.DataFrame(columns=["Faktur", "Type", "Detail"])
         ).to_excel(xl, sheet_name="Exceptions", index=False)
        pd.DataFrame([result.stats]).T.rename(columns={0: "value"}).to_excel(
            xl, sheet_name="Summary")
    return buf.getvalue()
