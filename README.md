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

## Sanity-check a new month / RO (no Coretax data needed)
```bash
python3 dry_run.py "<SAP extract.xlsx>" PALEMBANG 5 2026
```
Prints the detected sheets, SAP stats, doc-index, and the grid seed Salsa
would start from. Sheet detection is name-agnostic, so partial-download tabs
that get renamed each month (19-20, 15-19, 21-21, SENDIK SAP, …) are handled
automatically; the current-month `DATA OLAH` output sheet is excluded.

## Deploy to the web (Streamlit Community Cloud)

⚠️ **Data-sensitivity warning.** Uploaded files contain real NPWPs and amounts.
On Streamlit Community Cloud the app runs on a third-party (US) cloud, so any
file a user uploads is processed off-premises — it does **not** stay inside BRI.
For real bank data, prefer an internal host (a BRI machine / VM on the network).
If you deploy to the public cloud, at minimum: make the app **private**, set a
**password**, and ideally use it only with test/non-sensitive data.

Steps:
1. Push to GitHub (done — `riffarhan/bri-tax-automation`, private).
2. Go to https://share.streamlit.io → **Create app** → pick this repo / branch
   `main` / main file `app.py`.
3. Under **Advanced settings → Secrets**, add a password:
   ```toml
   app_password = "choose-a-strong-password"
   ```
   The app stays open with no secret; with this secret it shows a password gate.
4. In app **Settings → Sharing**, set it to specific viewers (not public) if the
   plan allows.

## Status
- ✅ Module 1 (template generator) + Module 2 (validation/exceptions) — built & validated
- ✅ Optional password gate for web deploys (set `app_password` secret)
- ⛔ Module 3 (Coretax RPA connector) — not built; Coretax data is uploaded manually for now
