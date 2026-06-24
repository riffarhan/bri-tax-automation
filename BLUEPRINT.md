# PPN WAPU → PSIAP — Blueprint (Pilot: RO Palembang)

## Goal
Salsa drops in the month's raw exports for an RO and gets back the **upload-ready
PSIAP `FM-Import` template** plus a **short exception list** of only the rows that
need a human. No VLOOKUPs, no manual re-keying, no manual dedup.

- **Platform:** Streamlit web app, **run internally** (localhost or a BRI internal
  server) — bank data never leaves BRI. Salsa opens it in a browser, desktop or laptop.
- **Accuracy model:** auto-fill everything that's certain; **flag the ambiguous**.
- **Scope:** Palembang + Yogyakarta, PPN WAPU first (this pilot), then PPh.

## What the data actually says (validated, not assumed)
- The `FM-Import` template is a **1:1 reshape of the Coretax Pajak Masukan data**
  (NPWP penjual, masa, tahun, nomor faktur — 81/81 in April).
- **`MASA_PAJAK` comes from Coretax, not SAP** (71/71 vs 16/71). A vendor issues the
  e-faktur in (say) March; BRI posts it in April/May. The filing uses the *faktur's*
  masa, which only Coretax knows.
- **SAP is the reconciliation cross-check**, not the template source: it supplies the
  `FIELD_TAMBAHAN_2` doc-number tag (Dokumen Invoice) and surfaces discrepancies.
- **Coretax has no bulk export** — it's one faktur at a time. (See "Open" below.)

## Architecture — 3 modules
1. **Template generator** `engine.reconcile` — Coretax faktur set → `FM-Import`. Deterministic.
2. **Validation engine** — dedup (last-6-digit key), DPP diff vs SAP, booked-not-reported,
   not-approved, doc-number resolution. Produces the exception report.
3. **Coretax connector (RPA)** — *not built yet.* A browser bot that logs into Coretax and
   pulls each faktur's data/status, replacing Salsa's manual lookups. The only browser piece.

Today the engine takes the Coretax data as an **upload** (the `Faktur Masukan` sheet);
Module 3 will later feed the same data automatically.

## FM-Import column mapping
| FM-Import column | Source | Rule |
|---|---|---|
| `FM` | const | `"FM"` |
| `NPWP_WP` | const | `0010016087093000` (BRI HO) |
| `ID_TKU_WP` | const | `000000` |
| `NOMOR_FAKTUR` | Coretax | faktur number (digits only) |
| `KONFIRMASI` | Coretax | `uncredited → 2` *(confirm `credited` code)* |
| `MASA_PAJAK` | Coretax | faktur masa |
| `TAHUN_PAJAK` | Coretax | faktur tahun |
| `MASA_PENGKREDITAN` | = MASA_PAJAK | |
| `TAHUN_PENGKREDITAN` | = TAHUN_PAJAK | |
| `NPWP_PENJUAL` | Coretax | vendor NPWP (16 digit) |
| `FIELD_TAMBAHAN_1` | Coretax/const | branch tag (`HO`) |
| `FIELD_TAMBAHAN_2` | derived | `PPNWAPUSAPRO{RO}{BULAN}{TAHUN}_{SAP Dokumen Invoice}` |

**Doc-number resolution (FT2 suffix), 3 tiers:** exact faktur (across Sheet1 +
BULAN SBLMNYA/BERIKUTNYA) → NPWP + nominal (DPP) → flag for manual.

## Validation vs Salsa's real April output
- Inputs: real April SAP extract + real Coretax `Faktur Masukan`. Target: her hand-made template.
- **Result: 885/891 cells (99.3%). Tax-critical fields (faktur, NPWP, masa, tahun,
  konfirmasi, constants) = 100%.**
- The 6 differences are all `FIELD_TAMBAHAN_2` (the reference tag) on ambiguous rows
  (duplicate/replacement fakturs + SAP faktur typos) — the engine flags these.
- Exceptions surfaced: 29 DPP diffs, 11 booked-not-in-Coretax, 3 doc-not-found, 1 fuzzy doc match.

## Open questions
1. **Coretax retrieval** (the remaining core piece): is there *any* Pajak Masukan
   list/download in Coretax, or do we build the RPA bot? Determines Module 3.
2. `KONFIRMASI` code for `credited` fakturs (all April rows were `uncredited → 2`).
3. Which SAP DPP column is authoritative for the recon (`DPP Amount Loc Currency VAT` used now).
4. Duplicate/replacement faktur handling — should the tool keep each faktur's own doc
   (current) or merge like Salsa does manually?

## Files
- `engine.py` — readers, reconciliation, FM-Import generation
- `writer.py` — Excel output (template + review workbook)
- `app.py` — Streamlit UI
- `validate_april.py` — regression test against the real April output
- `out_*.xlsx` — sample generated output for April Palembang
