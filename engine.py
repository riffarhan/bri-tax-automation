"""
PPN WAPU PSIAP — transformation & reconciliation engine.

Pilot scope: RO Palembang. Produces the PSIAP `FM-Import` template plus an
exception report, from two inputs:

  1. SAP Pajak Masukan extract  (the "...SAP RO..." workbook, sheet "Sheet1")
  2. Coretax Pajak Masukan list (NPWP penjual, faktur, masa, tahun, dpp, ppn,
     status, konfirmasi, valid) — supplied today as an upload; later fetched
     automatically by the Coretax connector (RPA).

Design notes (validated against April 2026 Palembang, see validate_april.py):
  * FM-Import is a 1:1 reshape of the Coretax faktur set. masa/tahun/npwp_penjual
    all come from Coretax — SAP does NOT carry the faktur masa.
  * SAP is used for reconciliation: the FT2 doc-number tag (Dokumen Invoice,
    joined by faktur) and to surface discrepancies (DPP diff, booked-not-reported).
  * Nothing is silently dropped: ambiguous rows go to the exception report.
"""
from __future__ import annotations
import re
import datetime as dt
from dataclasses import dataclass, field

import pandas as pd

# ----------------------------------------------------------------------------- config / constants
BULAN_ID = {1: "JANUARI", 2: "FEBRUARI", 3: "MARET", 4: "APRIL", 5: "MEI",
            6: "JUNI", 7: "JULI", 8: "AGUSTUS", 9: "SEPTEMBER", 10: "OKTOBER",
            11: "NOVEMBER", 12: "DESEMBER"}

FM_IMPORT_COLUMNS = ["FM", "NPWP_WP", "ID_TKU_WP", "NOMOR_FAKTUR", "KONFIRMASI",
                     "MASA_PAJAK", "TAHUN_PAJAK", "MASA_PENGKREDITAN",
                     "TAHUN_PENGKREDITAN", "NPWP_PENJUAL",
                     "FIELD_TAMBAHAN_1", "FIELD_TAMBAHAN_2"]

# Maps the Coretax "konfirmasi" status to the PSIAP import code.
KONFIRMASI_CODE = {"uncredited": 2, "credited": 1}  # TODO confirm 'credited' code w/ Salsa


@dataclass
class Config:
    ro_name: str = "PALEMBANG"          # appears in the FT2 label
    masa: int = 4                        # filing masa (month) — drives the label
    tahun: int = 2026
    npwp_wp: str = "0010016087093000"    # BRI HO
    id_tku_wp: str = "000000"
    branch_tag: str = "HO"               # FIELD_TAMBAHAN_1
    label_prefix: str = "PPNWAPUSAPRO"   # full label = prefix + RO + MONTH + YEAR

    @property
    def label(self) -> str:
        return f"{self.label_prefix}{self.ro_name}{BULAN_ID[self.masa]}{self.tahun}"


# ----------------------------------------------------------------------------- normalisation helpers
def norm_id(v) -> str:
    """Normalise an NPWP/NITKU/TIN-like value to a clean digit string."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    if isinstance(v, float):                       # e.g. 1.0016087093e+19
        v = f"{v:.0f}"
    s = str(v).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s


def norm_faktur(v) -> str:
    s = norm_id(v)
    return re.sub(r"\D", "", s)                     # keep digits only


def faktur_key(faktur: str, n: int = 6) -> str:
    """Dedup key — Coretax flags only 'similar' fakturs, so Salsa matches on the
    last few digits. Default 6 (between her 5–7)."""
    f = norm_faktur(faktur)
    return f[-n:] if len(f) >= n else f


def to_month(v):
    if isinstance(v, (dt.datetime, dt.date)):
        return v.month
    if isinstance(v, str):
        m = re.match(r"\s*(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})", v)
        if m:                                       # dd/mm/yyyy
            return int(m.group(2))
    return None


def to_num(v) -> float:
    if v is None or v == "":
        return 0.0
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


# ----------------------------------------------------------------------------- readers
SAP_RENAME = {
    "Kode Cabang Transaksi": "kode_cabang", "Dokumen Invoice": "dokumen_invoice",
    "Dokumen Pembayaran": "dokumen_pembayaran", "Dokumen Status": "dokumen_status",
    "Masa Pajak": "masa_sap", "Tahun Pajak": "tahun_sap",
    "DPP Amount Loc Currency VAT": "dpp_sap", "NIK/NPWP/TIN": "npwp_sap",
    "Nama": "nama_sap", "Tanggal Faktur": "tgl_faktur_sap", "Nomor FP": "nomor_faktur",
}

# Canonical Coretax columns the engine consumes, and the aliases it accepts
# (the April validation source is the "Faktur Masukan" sheet).
CORETAX_ALIASES = {
    "npwp_penjual": ["NPWP_PENJUAL", "npwp_penjual"],
    "nama_penjual": ["NAMA_PENJUAL", "nama_penjual", "NAMA"],
    "nomor_faktur": ["NOMOR_FAKTUR", "nomor_faktur"],
    "tanggal_faktur": ["TANGGAl_FAKTUR", "TANGGAL_FAKTUR", "tanggal_faktur"],
    "masa": ["MASA", "masa"],
    "tahun": ["TAHUN", "tahun"],
    "dpp": ["DPP", "dpp"],
    "ppn": ["PPN", "ppn"],
    "status": ["STATUS", "status"],
    "konfirmasi": ["KONFIRMASI", "konfirmasi"],
    "valid": ["VALID", "valid"],
    "branch": ["BRANCH", "branch", "FIELD_TAMBAHAN_1"],
}


def _read_sap_sheet(path, sheet) -> pd.DataFrame:
    """Read one SAP sheet, auto-detecting the header row (Sheet1 = row 1,
    BULAN SBLMNYA / BERIKUTNYA = row 2)."""
    probe = pd.read_excel(path, sheet_name=sheet, header=None, nrows=3, dtype=object)
    hdr_row = 0
    for i in range(min(3, len(probe))):
        if probe.iloc[i].astype(str).str.strip().eq("Nomor FP").any():
            hdr_row = i
            break
    df = pd.read_excel(path, sheet_name=sheet, header=hdr_row, dtype=object)
    df = df.rename(columns={k: v for k, v in SAP_RENAME.items() if k in df.columns})
    keep = [c for c in SAP_RENAME.values() if c in df.columns]
    df = df[keep].copy()
    df["_sheet"] = sheet
    for c in ("nomor_faktur",):
        if c in df: df[c] = df[c].map(norm_faktur)
    for c in ("npwp_sap", "dokumen_invoice"):
        if c in df: df[c] = df[c].map(norm_id)
    if "dpp_sap" in df: df["dpp_sap"] = df["dpp_sap"].map(to_num)
    return df


_SAP_MARKERS = {"Nomor FP", "Dokumen Invoice"}          # columns the doc-index needs


def _is_doc_source_sheet(path, sheet) -> bool:
    """True for any sheet carrying faktur→doc data — detected by having both
    'Nomor FP' and 'Dokumen Invoice' columns. Catches Sheet1, partial-download
    tabs (19-20, 15-19, 21-21, …) and the adjacent-month worked sheets
    (BULAN SBLMNYA/BERIKUTNYA), which hold the doc numbers for fakturs booked in
    a neighbouring month. REKON / DUPLIKAT lack these columns and drop out.
    Name-agnostic, so it survives month-to-month tab renames."""
    try:
        probe = pd.read_excel(path, sheet_name=sheet, header=None, nrows=3, dtype=object)
    except Exception:
        return False
    for i in range(min(3, len(probe))):
        vals = {str(x).strip() for x in probe.iloc[i].tolist()}
        if _SAP_MARKERS <= vals:
            return True
    return False


def sap_pull_sheets(path) -> list:
    import openpyxl
    names = openpyxl.load_workbook(path, read_only=True).sheetnames
    # 'DATA OLAH' is the current-month OUTPUT — exclude it so we don't reuse
    # Salsa's own answers (it's empty on a fresh month anyway). BULAN
    # SBLMNYA/BERIKUTNYA are kept: they hold real adjacent-month doc numbers.
    return [s for s in names
            if "OLAH" not in s.upper() and _is_doc_source_sheet(path, s)]


def read_sap(path, sheet="Sheet1") -> pd.DataFrame:
    """Current-masa SAP rows, used for reconciliation (DPP diff, booked-not-reported)."""
    import openpyxl
    available = openpyxl.load_workbook(path, read_only=True).sheetnames
    if sheet not in available:                       # fall back to the first SAP pull tab
        pulls = sap_pull_sheets(path)
        sheet = pulls[0] if pulls else available[0]
    return _read_sap_sheet(path, sheet)


def build_doc_index(path):
    """Faktur/doc lookup spanning every SAP pull tab (current + adjacent-month +
    partial downloads), since a faktur reported this masa may have been booked
    in an adjacent month."""
    frames = [_read_sap_sheet(path, s) for s in sap_pull_sheets(path)]
    if not frames:
        return {}, {}
    allrows = pd.concat(frames, ignore_index=True)
    allrows = allrows[allrows.get("dokumen_invoice", "").astype(str) != ""]
    by_faktur = {r["nomor_faktur"]: r["dokumen_invoice"]
                 for _, r in allrows.iterrows() if r.get("nomor_faktur")}
    # NPWP + DPP (rounded) -> doc, only when unambiguous
    by_amt = {}
    for _, r in allrows.iterrows():
        k = (norm_id(r.get("npwp_sap")), round(to_num(r.get("dpp_sap"))))
        if not k[0] or k[1] == 0:
            continue
        by_amt.setdefault(k, set()).add(r["dokumen_invoice"])
    by_amt = {k: next(iter(v)) for k, v in by_amt.items() if len(v) == 1}
    return by_faktur, by_amt


def resolve_doc(faktur, npwp, dpp, by_faktur, by_amt):
    """Three-tier: exact faktur -> NPWP+amount -> none. Returns (doc, how)."""
    if faktur in by_faktur:
        return by_faktur[faktur], "faktur"
    k = (norm_id(npwp), round(to_num(dpp)))
    if k in by_amt:
        return by_amt[k], "npwp+nominal"
    return "", "tidak ketemu"


def normalize_coretax(raw: pd.DataFrame) -> pd.DataFrame:
    """Canonicalise a Coretax table (from a file or a pasted/edited grid)."""
    if "FM" in raw.columns:                          # 'Faktur Masukan' sheet marker
        raw = raw[raw["FM"] == "FM"]
    out = {}
    for canon, aliases in CORETAX_ALIASES.items():
        col = next((a for a in aliases if a in raw.columns), None)
        out[canon] = raw[col].values if col is not None else None
    df = pd.DataFrame(out)
    df["nomor_faktur"] = df["nomor_faktur"].map(norm_faktur)
    df["npwp_penjual"] = df["npwp_penjual"].map(norm_id)
    df["dpp"] = df["dpp"].map(to_num)
    df["ppn"] = df["ppn"].map(to_num)
    df = df[df["nomor_faktur"] != ""].reset_index(drop=True)
    return df


def read_coretax(path, sheet=0) -> pd.DataFrame:
    return normalize_coretax(pd.read_excel(path, sheet_name=sheet, dtype=object))


def coretax_seed_from_sap(sap: pd.DataFrame, cfg: "Config") -> pd.DataFrame:
    """Pre-fill an editable grid from SAP so Salsa mostly only types the masa
    (and fixes status) she reads off Coretax — not the whole row."""
    inv = sap[sap["dokumen_status"] == "INVOICE"]
    return pd.DataFrame({
        "nomor_faktur": inv["nomor_faktur"].astype(str).values,
        "npwp_penjual": inv["npwp_sap"].astype(str).values,
        "nama_penjual": inv["nama_sap"].astype(str).values,
        "masa": [None] * len(inv),                   # <- from Coretax (key field)
        "tahun": [cfg.tahun] * len(inv),
        "dpp": inv["dpp_sap"].values,
        "ppn": [None] * len(inv),
        "status": ["approved"] * len(inv),
        "konfirmasi": ["uncredited"] * len(inv),
    })


# ----------------------------------------------------------------------------- core
@dataclass
class Result:
    fm_import: pd.DataFrame
    exceptions: pd.DataFrame
    stats: dict = field(default_factory=dict)


def reconcile(coretax: pd.DataFrame, sap: pd.DataFrame, cfg: Config,
              by_faktur=None, by_amt=None) -> Result:
    by_faktur = by_faktur or {}
    by_amt = by_amt or {}
    sap_inv = sap[sap["dokumen_status"] == "INVOICE"].copy()
    sap_by_faktur = {r["nomor_faktur"]: r for _, r in sap_inv.iterrows()
                     if r["nomor_faktur"]}

    rows, exc = [], []

    # duplicate detection (last-6-digit key) within the Coretax set
    dup_keys = (coretax.assign(_k=coretax["nomor_faktur"].map(faktur_key))
                .groupby("_k")["nomor_faktur"].transform("nunique"))
    coretax = coretax.assign(_dupgroup=coretax["nomor_faktur"].map(faktur_key),
                             _dupcount=dup_keys)

    for _, c in coretax.iterrows():
        fk = c["nomor_faktur"]
        s = sap_by_faktur.get(fk)
        dok, how = resolve_doc(fk, c["npwp_penjual"], c["dpp"], by_faktur, by_amt)

        # --- build the FM-Import row (every Coretax faktur is reportable) ---
        rows.append({
            "FM": "FM",
            "NPWP_WP": cfg.npwp_wp,
            "ID_TKU_WP": cfg.id_tku_wp,
            "NOMOR_FAKTUR": fk,
            "KONFIRMASI": KONFIRMASI_CODE.get(str(c["konfirmasi"]).strip().lower(), 2),
            "MASA_PAJAK": int(c["masa"]) if pd.notna(c["masa"]) else None,
            "TAHUN_PAJAK": int(c["tahun"]) if pd.notna(c["tahun"]) else None,
            "MASA_PENGKREDITAN": int(c["masa"]) if pd.notna(c["masa"]) else None,
            "TAHUN_PENGKREDITAN": int(c["tahun"]) if pd.notna(c["tahun"]) else None,
            "NPWP_PENJUAL": c["npwp_penjual"],
            "FIELD_TAMBAHAN_1": (norm_id(c["branch"]) or cfg.branch_tag),
            "FIELD_TAMBAHAN_2": f"{cfg.label}_{dok}" if dok else f"{cfg.label}_",
        })

        # --- exception checks (flag for review, never block) ---
        if how == "tidak ketemu":
            exc.append(_ex(fk, c, "Doc number not found",
                           "Not found in SAP by faktur number or by NPWP+amount — the document number must be filled in / verified manually."))
        elif how == "npwp+nominal":
            exc.append(_ex(fk, c, "Matched by NPWP + amount",
                           f"Faktur did not match SAP exactly (likely a typo); document {dok} was matched via NPWP + amount — please verify."))
        if s is not None:
            diff = to_num(c["dpp"]) - to_num(s["dpp_sap"])
            if abs(diff) >= 1:
                exc.append(_ex(fk, c, "DPP mismatch",
                               f"Coretax DPP {to_num(c['dpp']):,.0f} vs SAP {to_num(s['dpp_sap']):,.0f} (difference {diff:,.0f})"))
        if str(c["status"]).strip().lower() != "approved":
            exc.append(_ex(fk, c, "Not approved",
                           f"Coretax status = {c['status']} — faktur cannot be credited yet"))
        if c["_dupcount"] > 1:
            exc.append(_ex(fk, c, "Possible duplicate",
                           f"Faktur number is similar (last 6 digits = {c['_dupgroup']}) to another faktur"))

    # SAP fakturs that never appear in Coretax = booked but not reported
    coretax_keys = set(coretax["nomor_faktur"])
    for fk, s in sap_by_faktur.items():
        if fk not in coretax_keys:
            exc.append({"Faktur": fk, "Seller NPWP": s["npwp_sap"],
                        "Vendor": s.get("nama_sap"), "Masa": s.get("masa_sap"),
                        "Type": "Not in Coretax",
                        "Detail": "Present in SAP (INVOICE) but not found in Coretax — check whether it should be reported / different masa"})

    fm_import = pd.DataFrame(rows, columns=FM_IMPORT_COLUMNS)
    exceptions = pd.DataFrame(exc, columns=["Faktur", "Seller NPWP", "Vendor",
                                            "Masa", "Type", "Detail"])
    stats = {
        "coretax_fakturs": len(coretax),
        "sap_invoice_rows": len(sap_inv),
        "fm_import_rows": len(fm_import),
        "doc_filled": sum(1 for r in rows if r["FIELD_TAMBAHAN_2"] != f"{cfg.label}_"),
        "exceptions": len(exceptions),
    }
    return Result(fm_import, exceptions, stats)


def _ex(fk, c, etype, detail):
    return {"Faktur": fk, "Seller NPWP": c["npwp_penjual"],
            "Vendor": c.get("nama_penjual"), "Masa": c.get("masa"),
            "Type": etype, "Detail": detail}


def run(sap_path, coretax_path, cfg: Config,
        sap_sheet="Sheet1", coretax_sheet=0) -> Result:
    sap = read_sap(sap_path, sap_sheet)
    by_faktur, by_amt = build_doc_index(sap_path)
    coretax = read_coretax(coretax_path, coretax_sheet)
    return reconcile(coretax, sap, cfg, by_faktur, by_amt)
