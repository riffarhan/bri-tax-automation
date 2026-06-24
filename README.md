# PPN WAPU → PSIAP tool (pilot)

Turns the monthly SAP + Coretax exports for an RO into the upload-ready PSIAP
`FM-Import` template + an exception report. See `BLUEPRINT.md` for the design.

## Setup
```bash
cd ppn-wapu-tool
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
```

## Run the app (locally / internal server — not public cloud)
```bash
streamlit run app.py
```
Upload the SAP extract + Coretax Pajak Masukan file, review the flagged rows,
download the template.

## Check it still matches Salsa's real April output
```bash
python3 validate_april.py
```
Expect: 99.3% cell match, tax-critical fields 100%. (Reads the files in
`../WORK MAY 2026/PPN PALEMBANG/`.)

## Status
- ✅ Module 1 (template generator) + Module 2 (validation/exceptions) — built & validated
- ⛔ Module 3 (Coretax RPA connector) — not built; Coretax data is uploaded manually for now
