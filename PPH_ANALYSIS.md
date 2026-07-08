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
| SIPOBRI | **BIRTAX** (bukan SAP) | `NEW TEMPLATE PSIAP SIPOBRI ...xlsx` |

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
  (butuh login Coretax/BIRTAX; RPA nanti dulu).
- **Jalur manual (DIO)**: tool harus bisa nambah baris manual ke template
  (grid dynamic rows — pola yang sama udah ada di PPN).

## Yang udah kejawab, jangan ditanya lagi

- Kode Objek Pajak = KOP SAP verbatim (cek file, no mapping).
- NITKU Pemotong = 6 digit terakhir NITKU uker (master), NPWP Pemotong = BRI HO.
- SIPOBRI sumbernya BIRTAX; olahnya cuma "kolom hitam" tambahan Salsa.
- Manual dari DIO, langsung ke template, Yogya selalu ada.
- ETB PPh = `ETB PPh Unifikasi ...xlsx` sheet PALEMBANG/YOGYAKARTA (Salsa konfirmasi).
- **Struktur RO** (Salsa, 8 Jul): satu RO contains **KANINS + SENDIK + RO-nya sendiri**.
  Nama file kayak "RO & KANINS" / "SENDIK & RO" cuma nandain tarikan mana yang ada
  isinya bulan itu — **pengolahan datanya identik, diimpor bareng**. Reader kita
  tinggal concat semua tarikan, nggak ada format khusus per entitas.

## Masih perlu ditanya ke Salsa

1. **BIRTAX export** — bentuk filenya kayak apa, dan "kolom hitam" yang kamu
   tambahin itu kolom apa aja? (minta 1 contoh file)
2. **DIO** — data manualnya dikasih dalam bentuk apa (file/email/WA)?
3. **Prioritas** — pasal mana yang paling banyak baris / paling makan waktu?
   (pilot mulai dari situ; dugaan: PPh 22)
4. **Email / Fasilitas Insentif / Metode Pembayaran** di template — konstanta?
5. **NPWP 0000** — kira-kira berapa kasus per bulan?

## Rencana build (bertahap, kayak PPN dulu)

1. **Pilot: PPh 22 Palembang** — template generator (SAP `Sheet1` → sheet `Template`),
   validasi vs template real April/Maret. Buildable sekarang, nggak nunggu jawaban.
2. Rekon ETB PPh (reuse engine rekon PPN, per uker, tanpa reclass).
3. Scale ke PPh 23 (setelah jawab #1) + 4 ayat 2, lalu Yogyakarta + baris manual DIO.
4. Konsolidator file ekspor Coretax/SIPO per-uker + rekon PSIAP×Coretax.
5. SIPOBRI (setelah lihat contoh BIRTAX).
