# Alfa — Brief Presentasi

> **Alfa** = **Al**kaina + **Fa**rhan. Sistem buat ngebantu kamu ngerjain
> PPN WAPU → PSIAP, biar yang manual-manual itu otomatis. 🤍

---

## 1. Masalah yang diselesaikan

Tiap masa, buat tiap RO, kamu sekarang ngerjain ini manual:
- Tarik data dari **SAP**.
- Buka **Coretax** satu-satu per faktur buat validasi (status *approved*) & ambil
  data vendor.
- **VLOOKUP** nama uker, NITKU, dll.
- Susun **Template Impor PSIAP** manual (sesuai Petunjuk Isian).
- **Rekon**: klasifikasi PPN per uker per masa, cocokin **cabang ↔ induk** lewat
  nominal (yang kamu bilang *susah*), terus bandingin sama **ETB**.

Makan waktu, rawan typo, dan kerjaan-nya berulang tiap bulan.

## 2. Apa yang Alfa lakukan

Kamu **upload file**, Alfa keluarin **3 hal** otomatis:
1. **Template Impor PSIAP** — siap upload, sudah sesuai Petunjuk Isian.
2. **Daftar Pengecualian** — cuma baris yang perlu mata kamu (selisih DPP, duplikat,
   belum approved, dll). Sisanya udah otomatis.
3. **Rekon** — PPN per uker × masa vs saldo ETB; **SELISIH ≠ 0 → tinggal dicek**.

Yang dulu manual — enrich data, cocokin nomor dokumen, dedup faktur, **reclass
cabang → induk**, susun template, rekon — sekarang dikerjain Alfa.

## 3. Buktinya akurat (divalidasi sama hasil kamu sendiri)

Alfa diuji **dibandingkan sama file April Palembang yang kamu kerjain manual**:

| | Hasil |
|---|---|
| **Template Impor** | **99,3% sama** persis (885/891 sel). Kolom kritikal pajak (faktur, NPWP, masa, tahun, konfirmasi) **100% sama**. |
| **Rekon** | **23/23 uker** & **saldo ETB 23/23** cocok; bulan 9/10; selisih 22/23. |
| **Reclass cabang→induk** | **Otomatis** — uker 5742/5747/5758 dilipat ke induk 0059 tanpa cocok-cocokin manual. |

Dan ini jalan di data **Mei/Juni juga tanpa ganti kode** — bukan cuma April.

## 4. Cara pakai (3 langkah)

1. **Upload extract SAP** → grid langsung keisi dari SAP (faktur, NPWP, vendor, DPP).
2. **Isi kolom Masa** (masa pajak faktur dari Coretax) + betulin **Status** kalau ada
   yang belum approved. (Faktur yang cuma ada di Coretax: tambah baris.)
3. **Download** Template PSIAP. Buat Rekon, upload juga file **ETB**.

Yang kamu input cuma yang Coretax — sisanya Alfa.

## 5. Yang masih manual (jujur)

Ambil/validasi data **Coretax** masih manual, karena belum ada bulk export dari
Coretax. Tapi Alfa udah ngerjain semua sisanya (matching, dedup, reclass, template,
rekon), jadi kerjaan kamu tinggal validasi + isi masa.

> Nanti kalau perlu, bagian Coretax ini bisa diotomasi juga (bot), tapi sengaja
> belum — biar aman & simpel dulu.

## 6. Udah online

Bisa dibuka dari browser (laptop/HP), nggak perlu install apa-apa:
**https://bri-tax-automation-h2yginmd962wilkudq6wph.streamlit.app/**

Ada juga **glossary (ℹ️ Istilah)** di dalam app + tooltip di tiap kolom, biar nggak
bingung istilahnya.

---

*Dibuat sama Farhan, buat Alkaina.* 🤍
