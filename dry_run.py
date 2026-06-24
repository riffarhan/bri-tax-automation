"""
Dry-run the PPN WAPU pipeline on a SAP extract — no Coretax data required.

Shows: detected doc-source sheets, SAP stats, the doc-index, the grid seed
Salsa would start from, and confirms the pipeline runs end-to-end. Useful to
sanity-check a new month / RO before Salsa fills in the Coretax masa.

    python dry_run.py "<sap.xlsx>" [RO] [masa] [tahun]
"""
import sys
import warnings

warnings.filterwarnings("ignore")
from engine import (Config, read_sap, build_doc_index, coretax_seed_from_sap,
                    normalize_coretax, reconcile, resolve_doc, sap_pull_sheets)


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return
    path = sys.argv[1]
    cfg = Config(
        ro_name=(sys.argv[2] if len(sys.argv) > 2 else "PALEMBANG").upper(),
        masa=int(sys.argv[3]) if len(sys.argv) > 3 else 5,
        tahun=int(sys.argv[4]) if len(sys.argv) > 4 else 2026,
    )
    print(f"=== DRY-RUN: {cfg.ro_name} masa {cfg.masa}/{cfg.tahun} ===")
    print("Doc-source sheets:", sap_pull_sheets(path))

    sap = read_sap(path)
    inv = int((sap["dokumen_status"] == "INVOICE").sum())
    rec = int((sap["dokumen_status"] == "RECLASS").sum())
    print(f"SAP parsed: {len(sap)} rows ({inv} INVOICE, {rec} RECLASS)")

    by_faktur, by_amt = build_doc_index(path)
    print(f"Doc-index: {len(by_faktur)} faktur→doc, {len(by_amt)} npwp+nominal→doc")

    seed = coretax_seed_from_sap(sap, cfg)
    resolved = sum(1 for _, r in seed.iterrows()
                   if resolve_doc(r["nomor_faktur"], r["npwp_penjual"], r["dpp"],
                                  by_faktur, by_amt)[0])
    print(f"Grid seed: {len(seed)} fakturs pre-filled from SAP "
          f"({resolved}/{len(seed)} doc numbers auto-resolved)")

    res = reconcile(normalize_coretax(seed), sap, cfg, by_faktur, by_amt)
    print(f"Pipeline OK → template {res.stats['fm_import_rows']} rows; "
          f"Salsa only fills 'masa' from Coretax.")


if __name__ == "__main__":
    main()
