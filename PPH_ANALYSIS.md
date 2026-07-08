# PPh Unifikasi — hasil reverse-engineering

Digali dari: walkthrough call (transcript 53, 8 Jul 2026), jawaban WA Salsa (8 Jul 2026),
dan file `WORK MAY 2026/PPH PALEMBANG` (masa April 2026) + `WORK APRIL 2026` (masa Maret).

## Peran BRI

PPN WAPU: BRI **pemungut** (mungut PPN vendor). PPh Unifikasi: BRI **pemotong** —
motong PPh dari penghasilan pihak lain dan nerbitin **bukti potong**, diimpor ke
PSIAP/Pajakku pakai template per pasal.

## 4 stream output

| Stream | Sumber data | Template |
|---|---|---|
| PPh 22 | SAP | `NEW TEMPLATE PSIAP PPH 22 ...xlsx` |
| PPh 23 | SAP (tarikan RO + KANINS digabung) | `NEW TEMPLATE PSIAP PPH 23 ...xlsx` |
| PPh 4 ayat 2 | SAP | `NEW TEMPLATE PSIAP PPH 4 AYAT 2 ...xlsx` |
| SIPOBRI | **BRITAX** (bukan SAP) | `NEW TEMPLATE PSIAP SIPOBRI ...xlsx` |

Plus jalur samping: **transaksi manual** (penyetoran manual, nggak nyangkut SAP) —
datanya dari **DIO**, tidak diolah, langsung diketik ke template impor.
Palembang jarang ada; **Yogyakarta pasti ada tiap bulan**.

## Pipeline per stream (mirror PPN)

File SAP per pasal punya 3 sheet — pola sama persis kayak PPN WAPU:

1. **`Sheet1`** — tarikan SAP per masa. Kolom: Kode Cabang Transaksi, Nomor GL,
   Dokumen Invoice/Pembayaran + tanggal, Dokumen Status, Masa/Tahun Pajak,
   DPP, Amt.in loc.cur., **Rate Pajak**, **KOP**, NIK, NPWP, Nama.
   Untuk PPh semua Dokumen Status = "pembayaran" → **tidak ada reclass cabang→induk**
   (bagian tersulit PPN nggak ada di PPh).
   Tarikan Kanins & Kanwil terpisah, digabung ke satu sheet (format tarikan beda).
2. **`DATA OLAH`** — enrich. Kolom biru = yang Salsa isi manual: NPWP+nama penerima
   dari Coretax **hanya saat SAP-nya 0000/kosong**, plus NITKU (logic sama kayak PPN).
   Fokus PPh menurut Salsa: **validasi NPWP**.
3. **Template PSIAP** (bukti potong) — kolom kunci & sumbernya:
   - `NPWP Pemotong` = konstanta BRI HO `0010016087093000` (sama kayak PPN)
   - `NITKU Pemotong (6 Digit Terakhir)` = 6 digit terakhir NITKU uker → **dari master
     NPWP & NITKU Uker** (belum dibundle; master udah kita pegang)
   - `Masa/Tahun Pajak` = dari SAP
   - `NPWP/NITKU/Nama Penerima` = SAP; kalau 0000 → konfirmasi uker + validasi Coretax
   - `Kode Objek Pajak` = kolom `KOP` SAP **verbatim** (22-900-01 → 22-900-01, no mapping)
   - `Penghasilan Bruto` = DPP SAP; tarif udah ada di SAP (`Rate Pajak`)
   - `Dokumen Referensi` (jenis/nomor/tanggal) = dokumen SAP
4. **`Sheet2`** — rekon: `PCA L2 Desc | Utang Pajak PPh | PAJAK | SELISIH`.
   Logika sama kayak REKON PPN: PAJAK = akumulasi per uker, utang dari
   **ETB PPh Unifikasi** (sheet `PALEMBANG` / `YOGYAKARTA` — auto-detect ETB kita
   udah cocok), SELISIH harus 0.

## Lapisan rekon kedua (khas PPh): PSIAP × Coretax

Endpoint bulanan: `Rekonsiliasi PSIAP X Coretax RO ... Masa ....xlsx`
(sheet `SUMMARY KANWIL / DATA CORETAX / DATA PSIAP`; rumus dari tim Salsa).
Nge-balance dua hasil upload: data yang masuk PSIAP vs bukti potong di Coretax.

Buat "DATA CORETAX": Salsa narik **per NITKU satu-satu** dari Coretax (filter NITKU
per uker → tarik ±15 record) → hasilnya folder file per-uker
(`EKSPOR CORETAX .../150.xlsx, 438.xlsx, ...`; SIPO idem `SIPO PALEMBANG/138.xls, ...`)
→ dikonsolidasi pakai Power Query "Get Data" → dipindah ke file rekon.

## Keputusan desain (dari Farhan, 8 Jul 2026)

- **Logic sama kayak PPN**: template pre-fill dari SAP, Salsa validasi NPWP di grid
  (Coretax = validasi, bukan sumber). Narikan Coretax per-NITKU cuma dibutuhkan buat
  rekon akhir — bukan blocker template.
- **Otomasi yang bisa diotomasi**: konsolidasi file ekspor per-uker (gantiin Get Data)
  + rekon PSIAP×Coretax dibangun; narikan per-NITKU-nya sendiri tetap manual
  (butuh login Coretax/BRITAX; RPA nanti dulu).
- **Jalur manual (DIO)**: tool harus bisa nambah baris manual ke template
  (grid dynamic rows — pola yang sama udah ada di PPN).

## Yang udah kejawab, jangan ditanya lagi

- Kode Objek Pajak = KOP SAP verbatim (cek file, no mapping).
- NITKU Pemotong = 6 digit terakhir NITKU uker (master), NPWP Pemotong = BRI HO.
- SIPOBRI sumbernya BRITAX; olahnya cuma "kolom hitam" tambahan Salsa.
- Manual dari DIO, langsung ke template, Yogya selalu ada.
- ETB PPh = `ETB PPh Unifikasi ...xlsx` sheet PALEMBANG/YOGYAKARTA (Salsa konfirmasi).
- **Struktur RO** (Salsa, 8 Jul): satu RO contains **KANINS + SENDIK + RO-nya sendiri**.
  Nama file kayak "RO & KANINS" / "SENDIK & RO" cuma nandain tarikan mana yang ada
  isinya bulan itu — **pengolahan datanya identik, diimpor bareng**. Reader kita
  tinggal concat semua tarikan, nggak ada format khusus per entitas.

## BRITAX / SIPO — struktur kebuka (8 Jul 2026)

Konfirmasi Salsa: hasil tarikan BRITAX = folder `SIPO PALEMBANG` / `SIPO YOGYAKARTA`
(file `.xls` per kode uker, ditarik satu-satu), lalu dia Get Data lagi jadi satu sheet
(`GET DATA SIPO ...xlsx`: sheet `SIPO PALEMBANG` 380 baris → `DATA OLAH`).

Kolom per-uker `.xls` (sheet `Worksheet`): `Jenis Pajak | Branch | Masa Pajak |
Nama Vendor | NPWP Vendor | Jumlah Penghasilan | Jumlah PPH | Kode Objek Pajak |
Tarif | Jenis Dok Reff | No Dok Reff | Tanggal Dok Reff` — **udah bawa hampir semua
kolom template SIPOBRI** (kode objek, tarif, bruto, dok reff). "Kolom hitam" Salsa
kemungkinan tinggal NITKU/identitas — bisa kita derive dengan diff template vs export.

Catatan teknis: file `.xls`-nya BIFF lama yang ke-flag corrupt oleh xlrd →
baca pakai `xlrd.open_workbook(..., ignore_workbook_corruption=True)`.
Konsolidator kita: drop folder/multi-file → concat semua `Worksheet` → template.
NPWP vendor formatnya lama (`01.920.247.2-062.000`) → perlu normalisasi 15/16 digit.

## Masih perlu ditanya ke Salsa

1. **DIO** — data manualnya dikasih dalam bentuk apa (file/email/WA)?
2. **Prioritas** — pasal mana yang paling banyak baris / paling makan waktu?
   (pilot mulai dari situ; dugaan: PPh 22)
3. **Email / Fasilitas Insentif / Metode Pembayaran** di template — konstanta?
4. **NPWP 0000** — kira-kira berapa kasus per bulan?

## STATUS BUILD (8 Jul 2026) — engine SELESAI & tervalidasi

`engine_pph.py` + `writer_pph.py` + `validate_pph.py` dibangun dan divalidasi
lawan template real (April + Maret × Palembang + Yogyakarta = 16 kombinasi):

| Stream | Field kritis | Catatan |
|---|---|---|
| PPh 22 | 94.6–100% | jumlah baris pas |
| PPh 23 | 98.7–99.5% | 180/180 & 279/279 baris Yogya persis |
| PPh 4A2 | 94.0–98.8% | baris KOP kosong dibuang + di-flag (sesuai pola Salsa) |
| SIPOBRI | 99.6–100% | jumlah baris SEMUA persis (333/1307/334/1322) |

Sisa gap = koreksian manual Salsa (NPWP 0000 / NPWP BRI sendiri / typo SAP) —
semua ke-flag di exceptions, diisi lewat grid (pola PPN).

Aturan yang kebukti waktu build:
- **NITKU Pemotong**: SAP PPh bawa kolom NITKU sendiri (ambil 6 digit terakhir);
  master `data/uker_nitku.csv` jadi fallback + dipakai SIPO.
- **SIPO: uker = NAMA FILE** (tarikan per uker), bukan kolom Branch (verifikasi
  180/180). Referensi SIPOBRI = `{tag}_{nomor urut}`; SAP = `{tag}_{Dokumen Invoice}`.
- **Tanggal Dok Referensi** = SAP `Tanggal Invoice` (Maret 18/18); April Salsa
  override manual → biarkan editable.
- **Nama Penerima**: template real pakai dummy `NAMA111...` (PSIAP match by NPWP);
  kita isi nama SAP (lebih baik), beda ini nggak dihitung kritis.
- **KOP kosong** (deposito/reward, banyak di 4A2): keluar dari template + flag,
  tapi TETAP dihitung di rekon (pajaknya nyata di GL).
- **Rekon per pasal** = mirror Sheet2: `SELISIH = Utang (ETB kolom
  "Utang Pajak - PPh Pasal NN - SAP") - PAJAK (akumulasi per uker)`.
  Validasi April: PPh23 PLG PAJAK 8/8 + SELISIH 23/23, 4A2 SELISIH 23/23,
  PPh22 Yogya 19/19. Yogya 4A2 beda karena bucket REWARD/PHS/DEPOSITO
  dikategorikan manual oleh Salsa.
- ETB PPh label uker `00008 -- KC Baturaja` + `KANWIL X (Branch)` → parser
  `_uker_from_label`.
- Residual SELISIH ≠ SIPO (dites, bukan): itu memang selisih yang ditelusuri manual.

## Sisa kerjaan

1. **UI**: tab/halaman PPh di app Streamlit (upload SAP per pasal + folder SIPO,
   grid editable per stream, tambah baris manual DIO, download template+rekon).
2. Rekon lapisan 2 (PSIAP × Coretax) — konsolidator EKSPOR CORETAX per-uker.
3. Konfirmasi Salsa: DIO format, prioritas pasal, frekuensi NPWP 0000.
