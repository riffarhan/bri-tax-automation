# Alfa — Panduan Pengguna (PPN WAPU → PSIAP)

Alfa mengotomasi proses bulanan PPN WAPU: dari **extract SAP** + **data Coretax**
menjadi **Template Impor PSIAP** yang siap di-upload, plus daftar hal yang perlu
dicek manual. Yang dulu manual (VLOOKUP, isi ulang data vendor, susun template,
rekon) sekarang otomatis — Alfa hanya minta Anda mengisi bagian yang memang
butuh mata manusia.

> Terminologi & alur di panduan ini mengikuti istilah yang dipakai di lapangan
> (lihat glosarium). Narasi tombol/judul memakai bahasa Inggris; nama kolom &
> istilah pajak tetap bahasa Indonesia.

---

## Glosarium istilah

| Istilah | Arti |
|---|---|
| **Masa Pajak** | Periode pajak (1 bulan, format MM). Disingkat "masa". |
| **Masa Pajak Pelaporan** (*masa setor*) | Masa yang sedang dilaporkan/disetor — yaitu **Tax period (month)** di sidebar. |
| **Masa Pajak Faktur** | Masa milik faktur itu sendiri (dari Coretax) — yaitu kolom **Masa** di grid. Sering beda dari masa setor karena masa setoran PPN ±3 bulan. |
| **PPN WAPU** | PPN yang dipungut BRI sebagai Wajib Pungut. |
| **SAP** | Sumber data transaksi (extract bulanan; `Sheet1` = master data). |
| **Coretax / Pajakku** | Sistem DJP tempat memvalidasi faktur (status *approved*) dan data resmi faktur. |
| **Faktur Pajak / Nomor Faktur** | Nomor faktur pajak (e-faktur). |
| **Template Impor PSIAP** | Format file yang di-upload ke PSIAP. Mengikuti **Petunjuk Isian**. |
| **Data Olah** | Sheet kerja hasil olahan SAP + Coretax (di proses manual lama). |
| **Rekon** | Cocokkan PPN per uker per masa dengan saldo **ETB**. |
| **ETB** | File saldo buku besar (utang PPN WAPU) per uker. |
| **Uker / induk / cabang** | Unit kerja. Cabang (mis. `5747`) di-reclass ke induknya (mis. `0059`). |
| **Reclass** | Memindahkan nominal cabang ke induk — dicocokkan lewat nominal yang sama. |
| **NPWP / NITKU / DPP** | NPWP penjual, NITKU, dan Dasar Pengenaan Pajak. |

---

## Apa yang perlu di-upload

| Input | Dari mana | Untuk |
|---|---|---|
| **Extract SAP PPN WAPU** (.xlsx) | Download dari SAP | Template + Rekon. `Sheet1` + sheet bulan sebelumnya/berikutnya dibaca otomatis untuk nomor dokumen. |
| **Data Coretax** | Dari Coretax (Pajak Masukan), atau diisi di grid | Sumber masa faktur, NPWP, DPP, PPN, status. |
| **File ETB** (.xlsx) | File saldo ETB | Hanya untuk tab **Rekon**. |

---

## Langkah pakai

**1 · Period & source** (sidebar)
- Isi **Regional Office** (mis. PALEMBANG), **Tax period (month)** = masa setor,
  dan **Year**.
- Upload **extract SAP**. Grid Coretax langsung terisi dari SAP.
- (Opsional) Upload **file ETB** kalau mau pakai tab Rekon.

**2 · Coretax data (Pajak Masukan)**
- **Mode grid (default):** grid sudah terisi dari SAP (Nomor Faktur, NPWP, Nama
  Vendor, DPP). Anda tinggal isi kolom **Masa** = *Masa Pajak Faktur* dari Coretax,
  dan betulkan **Status** kalau ada faktur yang belum *approved*. Untuk faktur yang
  hanya ada di Coretax (tidak di SAP), tambah baris di bawah. Kalau ada nilai yang
  beda dari SAP dan Anda timpa, perubahannya tercatat di panel **"Changes from SAP"**.
- **Mode upload file:** kalau sudah punya data Coretax dalam Excel, upload saja.
  Mode ini juga mengisi **PPN** otomatis (dibutuhkan untuk Rekon).

> **Kenapa Coretax tetap dibutuhkan?** Seperti di pembahasan: Coretax itu
> **validasi** — datanya sebenarnya sudah ada di SAP, tapi status *approved* dan
> masa faktur resmi hanya bisa dipastikan dari Coretax.

**3 · Result** — tiga tab:
- **⚠️ Exceptions** — hanya baris yang perlu dicek manusia. Jenis: *DPP beda*,
  *Doc number tidak ketemu*, *Belum approved*, *Dugaan duplikat*, *Tidak ada di
  Coretax / SAP*, *Nomor faktur kosong*.
- **✅ FM-Import template** — hasil akhir, siap upload ke PSIAP. Tombol download
  ada di bawah ("PSIAP template (ready to upload)").
- **📊 Rekon** — PPN per uker × masa vs saldo ETB. **SELISIH ≠ 0 → perlu dicek.**
  Cabang otomatis dilipat ke induknya lewat reclass; yang tak bisa dipetakan
  muncul di daftar "Perlu reclass".

---

## Soal "masa" (penting, sering ketukar)

Ada **dua masa** yang beda:
1. **Tax period (month)** di sidebar = **Masa Pajak Pelaporan** (masa setor) — mis. April.
2. Kolom **Masa** di grid = **Masa Pajak Faktur** — masa milik faktur dari Coretax,
   sering bulan sebelumnya (faktur Maret dilaporkan di setoran April).

Kolom yang Anda isi dari Coretax adalah **Masa Pajak Faktur**; ini yang masuk ke
`MASA_PAJAK` di template.

---

## Catatan

- **Reclass otomatis:** pasangan cabang↔induk dicocokkan dari baris RECLASS di SAP
  (nominal yang sama) — jadi tidak perlu mencocokkan manual. Kalau ada yang tidak
  ketemu, akan ditandai untuk dicek.
- **Bagian manual yang tersisa:** mengambil/validasi data Coretax masih manual
  (belum ada bulk export dari Coretax). Alfa mengerjakan sisanya.
- **Data sensitif:** file berisi NPWP & nominal asli. Saat di-host di cloud publik,
  data yang di-upload diproses di luar BRI — pertimbangkan password & data uji.
