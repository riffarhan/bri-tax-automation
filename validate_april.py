"""
Validate the engine against Salsa's REAL April 2026 Palembang output.

Feeds the engine the actual SAP extract + the Coretax "Faktur Masukan" data,
then compares the generated FM-Import against the template Salsa produced by
hand, cell-for-cell.
"""
import sys
import pandas as pd
from engine import Config, run

BASE = "../WORK MAY 2026/PPN PALEMBANG/"
SAP = BASE + "21 APR - 20 MEI PPN WAPU SAP RO PALEMBANG APRIL 2026.xlsx"
CORETAX = BASE + "EKSPOR PSIAP PPN WAPU PALEMBANG APRIL 2026.xlsx"   # sheet 'Faktur Masukan_1'
EXPECTED = BASE + "Template Impor PPN WAPU PSIAP RO PALEMBANG APRIL 2026 - PSIAP.xlsx"

cfg = Config(ro_name="PALEMBANG", masa=4, tahun=2026)
res = run(SAP, CORETAX, cfg, coretax_sheet="Faktur Masukan_1")

print("STATS:", res.stats)

# load expected FM-Import
exp = pd.read_excel(EXPECTED, sheet_name="FM - Import", dtype=object)
exp = exp[exp["FM"] == "FM"].reset_index(drop=True)

got = res.fm_import.copy()

# normalise both to comparable strings, keyed by nomor faktur
def keyframe(df):
    d = df.copy()
    for c in ["NOMOR_FAKTUR", "NPWP_WP", "ID_TKU_WP", "NPWP_PENJUAL", "FIELD_TAMBAHAN_2", "FIELD_TAMBAHAN_1"]:
        d[c] = d[c].astype(str).str.strip()
    for c in ["KONFIRMASI", "MASA_PAJAK", "TAHUN_PAJAK", "MASA_PENGKREDITAN", "TAHUN_PENGKREDITAN"]:
        d[c] = pd.to_numeric(d[c], errors="coerce").astype("Int64").astype(str)
    return d.set_index("NOMOR_FAKTUR")

E, G = keyframe(exp), keyframe(got)

print(f"\nRow counts — expected {len(E)}, generated {len(G)}")
only_e = set(E.index) - set(G.index)
only_g = set(G.index) - set(E.index)
print(f"Fakturs only in EXPECTED: {len(only_e)}  | only in GENERATED: {len(only_g)}")
if only_e: print("   missing from generated:", list(only_e)[:5])
if only_g: print("   extra in generated:", list(only_g)[:5])

cols = [c for c in E.columns if c in G.columns]
common = sorted(set(E.index) & set(G.index))
print(f"\nCell match on {len(common)} shared fakturs:")
total = mismatch = 0
mismatch_by_col = {}
for fk in common:
    for c in cols:
        total += 1
        if str(E.loc[fk, c]) != str(G.loc[fk, c]):
            mismatch += 1
            mismatch_by_col[c] = mismatch_by_col.get(c, 0) + 1
print(f"   {total - mismatch}/{total} cells match ({100*(total-mismatch)/total:.1f}%)")
if mismatch_by_col:
    print("   mismatches by column:", mismatch_by_col)
    # show a few examples for the worst column
    worst = max(mismatch_by_col, key=mismatch_by_col.get)
    print(f"\n   sample mismatches in '{worst}':")
    n = 0
    for fk in common:
        if str(E.loc[fk, worst]) != str(G.loc[fk, worst]):
            print(f"      {fk}: expected={E.loc[fk, worst]!r}  got={G.loc[fk, worst]!r}")
            n += 1
            if n >= 5: break

print(f"\nEXCEPTIONS flagged: {len(res.exceptions)}")
print(res.exceptions["Type"].value_counts().to_string() if len(res.exceptions) else "  (none)")
