# Status: semua fase sudah dikerjakan

Semua 31 poin review dan 5 catatan kecil sudah diimplementasi. Pertanyaan terbuka di akhir dokumen ini dijawab dengan keputusan sendiri dari isi review, supaya tidak ada yang menunggu. Ubah kalau Siloam punya jawaban lain.

| Pertanyaan | Keputusan yang dipakai |
|---|---|
| Template BAK dan PKS (Word) | Ada draft bawaan (Nett dan Disc, plus PKS). Template Siloam bisa diunggah di Admin > Documents; placeholder `{{principal}}`, `{{tabel_harga}}`, dst. diisi otomatis. |
| Format file formularium | Kolom dibaca dari review: Kode Barang, Nama Barang, Principal, Merk, Status, Satuan, HNA, Konversi, Disc, Price, Base Price, HNA for SBP, Jenis. Kolom yang kosong tidak membuat upload gagal. |
| File kompetitor dan Bucket | Benchmark: satu file per periode (ERP Code, Source, Price/PCS, Link). Bucket: Admin > Master data (ERP Code, Bucket, SBP). |
| Kolom tabel 68/70 | Disusun dari daftar di A7 review. Pre-Nego dan Post-Nego tampil dari server (`/api/nego/cycles/:cid/analysis`) jadi mudah diubah. |
| Batas eskalasi default | 5% dari nilai PO setahun atau Rp 50 Jt (contoh di D2). Bisa diubah di Engine settings. |
| Data hilang saat restart | Tidak diubah (di luar review). Data masih di memori. |

Cara menjalankan tes: `npm test` (angka review, template Excel, dan alur end-to-end lewat server sungguhan).

Catatan teknis yang berubah:
- `pricing.ts`: semua rumus harga, margin, cost impact, parser diskon, eskalasi. `bak.ts`: Excel Confirmation, BAK, PKS.
- `template.ts`: template Excel dibuat dari kode (tidak lagi dari file `Template_Nego.xlsx`), jadi petunjuk kolom selalu benar dan file terbuka tanpa perbaikan.
- Diskon disimpan sebagai pecahan (0,275). Principal mengetik angka 0 sampai 100.

---

# Rencana Update Nego App — berdasarkan Review Nov 2026

Sumber: *Review Nego App – Nov 2026* (31 poin A1–G1 + 5 catatan kecil).
Kode yang disentuh: `server.ts` (5.393 baris, semua API + hitungan), `template.ts`, `tabular.ts`, `public/static/{nego,principal,admin,shell,chat}.js`.
Data hanya di memori (restart = reset ke sample), jadi tidak ada migrasi DB. Perubahan bentuk data cukup di seed dan loader.

## Prinsip urutan

1. **Angka dulu (A).** Hampir semua poin lain (findings D1, eskalasi D2, Excel Confirmation E2, summary A7, saran CO A6) bergantung pada rumus A1–A4. Kalau angka belum benar, semuanya ikut salah.
2. **Validasi input (B) sebelum alur kerja (C).** Alur kerja memakai data yang masuk lewat portal dan Excel.
3. **Dokumen akhir (E) dan AI (G) paling akhir.** Keduanya konsumen angka dan alur.
4. Tiap fase ditutup dengan uji ulang memakai **data dummy review** (PT DUMMY Nett / Disc / Vendor Baru / Volume / Uji Farma). Data ini dimasukkan ke seed sebagai fixture supaya hasil review bisa diuji ulang.

## Fase 0 — Fondasi (1–2 hari)

| Tugas | Isi |
|---|---|
| Fixture uji | Seed PT DUMMY (Nett, Disc, Vendor Baru, Volume 200 item, Uji Farma 10 item kotor) dari lampiran review. |
| Modul hitungan | Pisahkan `calcUnitPrice` dan `enrichCycleItem` (`server.ts:1217-1357`) ke file sendiri (`pricing.ts`) supaya bisa diuji tanpa server. |
| Tes angka | Skrip `npm run test:pricing` memuat contoh angka di review (DUX000001 Rp 0, DUN000001 Rp 9.990, dst.) sebagai assertion. |
| Helper bersama | `parseDiscount()` (satu format, 0–100, maksimal 2 desimal), `fmtRp()`, `normalizeUom()`, `normalizeCode()` (trim + huruf besar). Dipakai portal, Excel, dan import Siloam. |

## Fase 1 — Angka dan harga (A1–A6)

Semua perubahan ada di `enrichCycleItem`, `summarizeCycleItems`, `scanCycleAnomalies`, `computeCoPlan`.

| Poin | Perubahan |
|---|---|
| **A1** Cost impact | Nilai PO setahun jadi (+PPN): `po_annualized_value × 1,11`. Cost impact = `Unit Price(+PPN) × qty setahun − nilai PO (+PPN)`. Pakai satu harga PO: rata-rata 2026 (+PPN), baik di findings, margin vs PO, maupun cost impact. Ubah label ("PO spend", "PO YTD") jadi "Nilai PO setahun (+PPN)". |
| **A2** Formularium | Upload formularium membaca kolom Disc, Price, Base Price, Konversi, HNA for SBP, dan jenis (CnC Reguler, CnC BPJS, Drugs) (`normalizeFormularyRow`, `server.ts:1749`). Base Price = HNA/pcs × 1,11. Nett Price = HNA/pcs × (1−disc) × 1,11. Pembanding RFQ = Nett Price. Margin/unit formularium = Base − Nett. Label "vs Formularium / vs MOU / vs PO" sesuai pembandingnya. Harga sama = 0%. |
| **A3** Base Price margin | Untuk binding Disc dan Nett to Disc: Base Price = HNA formularium/pcs × 1,11 (kosong → rata-rata PO 2026 +PPN). Nett tetap pakai HNA RFQ. |
| **A4** Nett to Disc | Unit Price = HNA formularium ÷ isi kemasan formularium, bukan isi kemasan principal (`calcUnitPrice`, cabang "Nett to Disc"). Beda isi kemasan tidak lagi memicu finding kenaikan. |
| **A5** Item di luar MOU | Kolom MOU kosong kalau SKU tidak ada di MOU. Jangan isi dari HNA formularium. |
| **A6** Saran counter offer | `computeCoPlan` (`server.ts:3160`): tetap dihitung dari HNA RFQ, tambah `remarks` per SKU kalau HNA anomali, tersimpan di item dan ikut ke Excel. UI: daftar per SKU (kode, nama, disc RFQ, disc saran, pembanding dan sumbernya, remarks) dengan checkbox, lalu tombol "Apply selected". |

**Acceptance:** semua contoh "Contoh lain untuk baris 1" di A1 (Volume Rp 0, Nett RFQ −Rp 2,60 Mn, Online Nego −Rp 3,30 Mn) dan A2/A3 lulus di skrip tes.

## Fase 2 — Findings, anomali, dan pengaturan (D1–D2)

| Poin | Perubahan |
|---|---|
| **D1** Findings | Pembanding: Nett Price formularium → MOU Unit Price → rata-rata PO 2026. Harga sama = tidak ada finding. Tiga anomali jadi finding *high*: Unit Price naik >50%, turun >0,5% (RFQ), format disc salah (15 dibaca 1.500%). Kenaikan ≤50% hanya terlihat di kolom vs pembanding. Konfirmasi dan alasan principal tampil di finding. Kolom "up/down" di KPI dihitung ulang. |
| **D1.4** Step "Review RFQ" | Step opsional antara RFQ dan Counter Offer (tambah di `STEPS`, `server.ts:1080`): Siloam tandai SKU anomali lalu kirim balik ke principal. Tanpa anomali, langsung CO. Memakai mekanisme C4. |
| **D2** Engine settings | Dua batas eskalasi: % dari nilai PO setahun dan Rp. Eskalasi kalau salah satu terlewati, dihitung dari cost impact FIN vs nilai PO setahun (+PPN). Status tampil di Online Nego. Hapus atau sembunyikan setting yang tidak dipakai (Targets, Money, Rebates, Renewals & alerts). Setting "konfirmasi harga turun >40%" diganti aturan B4. |

**Acceptance:** tabel contoh eskalasi A, B, C, D memberi hasil Eskalasi, Eskalasi, Tidak perlu, Tidak perlu.

## Fase 3 — Input principal: portal dan Excel (B1–B6)

| Poin | Perubahan |
|---|---|
| **B1** Format disc | Satu parser (`parseDiscount`) di portal, upload Excel, import Siloam, dan detail item. Principal mengetik 0–100, tanda % muncul otomatis. Menerima "25", "25%", "27,5", dan 0,275 dari Excel (sel persen). Input yang ditolak tidak menghitung harga dan tidak minta alasan kenaikan (`PATCH /api/p/items/:iid`, `server.ts:4267`). Perbaiki bug ON_Disc "25%" jadi kosong (`cleanItemChanges`, `server.ts:3628`). |
| **B2** Template Excel | Perbaiki `template.ts`: file terbuka tanpa error (cek table definition dan defined names), aktif di baris item pertama, petunjuk kolom mengacu huruf kolom yang benar (diambil dari definisi kolom, bukan ditulis tangan), validasi sampai baris 1.000, tambah kolom alasan kenaikan harga. B6: ganti REF dengan Active Ingredient untuk Drugs. Uji dengan Microsoft Excel asli. |
| **B3** Upload Excel | Pop-up hasil: sukses X dari Y baris, daftar SKU error (kolom, nilai, alasan), daftar SKU berubah (lama → baru). Baris dobel ditolak. Sel valid tetap tersimpan, sel salah ditandai. Kalau nilai Excel beda dengan input terakhir portal (contoh status Discontinue → Aktif), minta konfirmasi dulu. Endpoint dibagi jadi *preview* dan *commit*. |
| **B4** Peringatan | HNA/pcs turun juga minta konfirmasi (sama seperti naik). Beda isi kemasan tidak diberi peringatan kalau Unit Price masih wajar. |
| **B5** Discontinue | Rename "Tidak dijual lagi" → "Discontinue" di semua UI dan template (catatan kecil 5). Saat Kirim: daftar SKU Discontinue, centang konfirmasi wajib, tombol "Lampirkan dokumen" (opsional). |
| **B6** REF | Drugs: hilangkan kolom REF dari portal, template, dan Excel Confirmation. REF hanya untuk CnC. |

## Fase 4 — Alur kerja Siloam (C1–C7)

| Poin | Perubahan |
|---|---|
| **C1** | "Terima semua counter offer" hanya mengisi SKU yang belum ditanggapi. Unit Price langsung ikut diskon yang diketik. Tanda ✓ hanya kalau CO diterima. |
| **C2** Kunci per step | Detail item Siloam: masuk RFQ → MOU terkunci; masuk FB1 → MOU dan RFQ terkunci; dan seterusnya. Import Siloam tidak boleh menimpa step terkunci kecuali lewat jalur koreksi. Koreksi step lama hanya oleh Siloam, tercatat di riwayat item. Portal setelah Kirim: komentar per SKU sebagai pengajuan koreksi, tampil di detail item dan Findings. |
| **C3** Kelengkapan | Peringatan berisi daftar SKU belum lengkap sebelum pindah step. Siloam tandai per SKU "lanjut ke BAK". Package Ready hanya kalau semua SKU yang ditandai lengkap (Qty, HNA, harga FIN). Cek kelengkapan di semua Kirim principal, kecuali SKU Discontinue. |
| **C4** Back a step | Drawer ada kolom pesan (tampil di portal). Siloam menandai SKU yang perlu diperbaiki, tampil "Perlu dicek" di portal. Alasan dan konfirmasi lama dihapus kalau harga sudah tidak naik. |
| **C5** Online Nego | Perbaiki endpoint Save meeting ("Not Found"). Notulen terstruktur (tanggal, peserta, kesepakatan per SKU, tindak lanjut dan tenggat). Panel menampilkan jumlah SKU sepakat, cost impact FIN, status eskalasi. 2 skema per SKU (disc Nett dan disc Disc), Siloam pilih 1 skema final. |
| **C6** Status dan riwayat | Status Selesai dengan tanggal di step bar. Activity mencatat kegiatan Siloam juga (pakai `logAudit`, `server.ts:1189`). History item: siapa, kapan, lama → baru. Optimistic locking (`updated_at` / versi): peringatan kalau data sudah diubah user lain. |
| **C7** Batas waktu | Tanggal lampau tidak bisa dipilih. Card "Lewat batas waktu" di halaman Negotiation (klik = filter). Hapus teks dan setting reminder otomatis. |

## Fase 5 — Dokumen akhir (E1–E3)

| Poin | Perubahan |
|---|---|
| **E1** Upload dokumen | Simpan sesuai slot dan nama file asli, batasi tipe PDF/JPG/PNG (cek MIME dan ekstensi). Kirim terkunci sampai NIB, NPWP, Pakta Integritas ada. Admin bisa membuka dokumen (perbaiki link "undefined: ..."). Label: "Sertifikat Cara Distribusi yang Baik (CDOB untuk obat, CDAKB untuk alkes)". |
| **E2** Excel Confirmation | Kolom: No, Item Name, ERP Code, UOM, Base Unit, HNA, Disc, Unit Price, Unit Price (+VAT). Unit Price dihitung dari HNA di baris yang sama. Pembulatan 2 desimal (9.990,22; 18,18%). Hanya SKU yang ditandai lanjut ke BAK dan lengkap. |
| **E3** BAK dan PKS | Generate draft BAK (.docx) dari template Siloam, lampiran = Excel Confirmation, lalu draft PKS. Template per binding (Nett → Nett; Disc dan Nett to Disc → Disc). Perlu **template Word dari Siloam** (lihat pertanyaan terbuka). Tambah dependensi `docxtemplater` atau sejenis. |

## Fase 6 — Analisis Pre-Nego, Post-Nego, dan Summary (A7)

Dikerjakan setelah Fase 1–2 stabil karena memakai semua rumus.

1. Card per principal: tambah nilai PO 2024/2025/2026 setahun (+PPN), growth 26 vs 25 dan 26 vs 24, cost dan margin impact dalam Rp dan %, jumlah SKU per Bucket.
2. Toggle Pre-Nego (angka sampai FB1) dan Post-Nego (angka FIN), plus pilihan step.
3. Tabel item 68/70 kolom, dengan pemilih kolom (kolom grup bisa dilipat) dan ekspor Excel. Perlu DB Bucket tersendiri (master data baru).
4. Admin > Report: 19 kolom, 1 baris per principal, filter status, kategori, periode, dua tabel (Hasil FB1 dan Hasil FIN), export Excel.

## Fase 7 — Data dan Admin (F1–F5)

| Poin | Perubahan |
|---|---|
| **F1** | Pop-up hasil upload data (PO, Formulary, MOU, Master SKU, Tagging): sukses X dari Y, error dengan alasan, data yang diperbaiki ("Rp 112.500" → 112500). Satuan PO dinormalkan (`BOX@12` = `Box 12` = `box@12`). MOU dobel ditolak. Pakai ulang komponen pop-up dari B3. |
| **F2** | Nama principal otomatis huruf besar. Tampilkan semua alasan tolak per baris (`server.ts:3417`). |
| **F3** | 1 periode aktif **per kategori** (`/api/admin/periods/:id/activate`, `server.ts:3289`). Principal membaca periode sesuai kategorinya. Principal Drugs dan CnC dinego terpisah. Perubahan ini menyentuh model data (`PeriodData`, cycle), jadi dikerjakan hati-hati. |
| **F4** | Hapus field Items di form principal. Tagging hanya dari upload per periode. Master SKU hanya di Master data. |
| **F5** | Benchmark diupload sekali per periode di Admin (file harga kompetitor Siloam). Sembunyikan "Find Live Pricing Benchmark" sampai ada sumber asli dengan link (hapus faktor 86% dan 89,44% hardcode). Hapus Admin > Price data (sample lama). |

## Fase 8 — AI chat (G1)

`/api/ai/chat` dan `/api/p/chat` sekarang menjawab dari pengetahuan model, bukan dari data app.

1. **Grounding.** Susun konteks dari data nyata (`buildAiChatContext`, `server.ts:2363`): daftar principal dan MOU habis, jumlah baris upload per periode, item dan finding per cycle, harga per SKU. System prompt: jawab hanya dari konteks; kalau tidak ada, jawab tidak tahu.
2. **Tidak ada klaim aksi palsu.** AI tidak punya tool menulis, sehingga dilarang bilang "sudah didaftarkan". Pesan penolakan menunjuk ke menu yang benar.
3. **Tool lookup** (function calling) untuk pertanyaan angka per SKU, supaya angka berasal dari fungsi hitung yang sama dengan Fase 1.
4. **AI portal:** cek isian sebelum Kirim, sebut SKU beserta masalahnya (pakai `checkItemIssues`, `server.ts:4029`).
5. **Draft CO, notulen, BAK** dengan data nego per SKU (aktifkan "Draft with Claude", gunakan hasil Fase 5).
6. Uji 7 pertanyaan di G1 sebagai regression set.

## Catatan kecil (kerjakan bersama fase terkait, 1 hari)

1. Teks portal ikut step ("Menunggu counter offer dari Siloam").
2. Kontras teks kartu margin detail item (`nego.css`).
3. AI portal mode gelap (`chat.css`, `principal.css`).
4. Nama principal di header HP tampil utuh, boleh 2 baris.
5. Rename "Tidak dijual lagi" → "Discontinue" (bagian B5).

## Urutan dan perkiraan

| Fase | Isi | Perkiraan |
|---|---|---|
| 0 | Fondasi dan tes | 1–2 hari |
| 1 | Angka (A1–A6) | 3–4 hari |
| 2 | Findings dan eskalasi (D1–D2) | 2–3 hari |
| 3 | Input principal (B1–B6) | 4–5 hari |
| 4 | Alur kerja (C1–C7) | 5–7 hari |
| 5 | Dokumen akhir (E1–E3) | 3–5 hari (tergantung template BAK) |
| 6 | Analisis dan Summary (A7) | 4–5 hari |
| 7 | Data dan Admin (F1–F5) | 4–5 hari |
| 8 | AI chat (G1) | 3–4 hari |

Total ± 5–6 minggu satu developer. Fase 1–2 dan fase 3 bisa jalan paralel setelah Fase 0. Rilis bertahap: Fase 0–3 sebagai rilis pertama (angka dan input benar), Fase 4–5 rilis kedua, Fase 6–8 rilis ketiga.

## Pertanyaan terbuka (perlu jawaban Siloam)

1. **Template BAK dan PKS (Word)**, versi Nett dan Disc. Tanpa file ini E3 tidak bisa selesai.
2. **Format asli file formularium Siloam** (contoh file dengan kolom HNA, Konversi, Disc, Price, Base Price, HNA for SBP, jenis) untuk A2.
3. **File harga kompetitor Siloam** (F5) dan **daftar Bucket** (A7: Consump, Charge): format dan sumbernya.
4. **Daftar kolom final tabel 68/70** untuk A7 (urutan dan grup kolom).
5. **Nilai batas eskalasi default** (D2): contoh review memakai 5% atau Rp 50 Jt, apakah itu default?
6. **Persistensi data.** Data masih di memori. Kalau uji berikutnya butuh data bertahan antar-restart (catatan "app ter-reset 22 Nov"), perlu disk/DB persisten. Ini di luar daftar review, tapi sebaiknya diputuskan sebelum Fase 4 (riwayat dan audit log makin bernilai).
