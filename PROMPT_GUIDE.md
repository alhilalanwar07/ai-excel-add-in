# Prompt Guide (Copy-Paste)

Panduan ini berisi prompt siap pakai untuk fitur utama Excel AI Gemini.
Setiap kali fitur ditambah/diubah, update file ini dan `src/taskpane/constants/promptTemplates.ts`.

## Level Template

- Basic: Prompt harian yang aman untuk pemula.
- Advanced: Prompt analitis dengan parameter lebih spesifik.
- Automation: Prompt otomasi multi-langkah, cocok untuk data besar.

## Formula

Tulis formula di C2:C200 dengan pola =A2*B2 dan pastikan formula menyesuaikan setiap baris.

Gunakan bulk_write_formulas untuk mengisi formula ke C2:C1000 berdasarkan A dan B, gunakan payload rangeAddress+formulas agar eksekusi cepat.

## Formatting

Format range A1:D1 menjadi bold, background kuning muda, lalu autofit kolom.

## Data Insert

Masukkan tabel contoh di A1 dengan header [Produk, Qty, Harga], isi 5 baris data dummy penjualan.

## Clear / Cleanup

Kosongkan isi range F2:F200 saja tanpa menghapus format kolom lain.

## Sort / Filter

Urutkan data pada range A1:E500 berdasarkan kolom E dari terbesar ke terkecil.

Aktifkan filter otomatis pada range A1:G1000.

## Chart

Buat chart column dari range A1:B12 dengan judul Penjualan Bulanan di sheet aktif.

## Pivot Summary

Buat pivot_summary dari A1:C1000, groupBy kolom 0, value kolom 2, agregasi sum, outputSheetName RingkasanPenjualan.

Buat pivot_summary dari A1:C5000, groupBy kolom 0, value kolom 2, agregasi sum, sortDescending true, topN 10, outputSheetName TopKategori, createChart true, chartType column, chartTitle Top 10 Kategori.

Buat pivot_summary dari A1:C5000, groupBy kolom 0, value kolom 2, agregasi sum, sortDescending true, minValue 100000, topN 12, outputSheetName RingkasanPremium, createChart true, chartType bar, chartPreset executive, chartTitle Kategori Nilai Tinggi.

## Multi-Sheet

Terapkan format header tebal dan warna biru muda untuk Sheet1, Sheet2, dan Sheet3 pada range A1:H1.

## Maintenance Checklist

- Tambah template prompt baru di `src/taskpane/constants/promptTemplates.ts`.
- Tambah contoh prompt copy-paste di file ini.
- Jika action baru ditambahkan, update skema pada `src/taskpane/utils/gemini.ts`.
- Jika execution payload berubah, update validasi/handler di `src/taskpane/services/ExcelService.ts`.
- Pastikan setiap template memiliki level `Basic`, `Advanced`, atau `Automation`.
