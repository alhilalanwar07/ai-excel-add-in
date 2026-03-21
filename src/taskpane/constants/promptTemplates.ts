export interface PromptTemplate {
  id: string;
  title: string;
  prompt: string;
  level: "Basic" | "Advanced" | "Automation";
}

export const PROMPT_TEMPLATES: PromptTemplate[] = [
  {
    id: "formula-single",
    title: "Tulis Formula Satu Kolom",
    level: "Basic",
    prompt:
      "Tulis formula di C2:C200 dengan pola =A2*B2 dan pastikan formula menyesuaikan setiap baris.",
  },
  {
    id: "formula-bulk-fast",
    title: "Formula Massal Cepat (bulk_write_formulas)",
    level: "Advanced",
    prompt:
      "Gunakan bulk_write_formulas untuk mengisi formula ke C2:C1000 berdasarkan A dan B, gunakan payload rangeAddress+formulas agar eksekusi cepat.",
  },
  {
    id: "format-range",
    title: "Format Range",
    level: "Basic",
    prompt:
      "Format range A1:D1 menjadi bold, background kuning muda, lalu autofit kolom.",
  },
  {
    id: "insert-table",
    title: "Insert Data Tabel",
    level: "Basic",
    prompt:
      "Masukkan tabel contoh di A1 dengan header [Produk, Qty, Harga], isi 5 baris data dummy penjualan.",
  },
  {
    id: "clear-safe",
    title: "Bersihkan Range Tertentu",
    level: "Basic",
    prompt:
      "Kosongkan isi range F2:F200 saja tanpa menghapus format kolom lain.",
  },
  {
    id: "sort-data",
    title: "Sort Data",
    level: "Advanced",
    prompt:
      "Urutkan data pada range A1:E500 berdasarkan kolom E dari terbesar ke terkecil.",
  },
  {
    id: "filter-data",
    title: "Aktifkan Filter",
    level: "Advanced",
    prompt:
      "Aktifkan filter otomatis pada range A1:G1000.",
  },
  {
    id: "chart-create",
    title: "Buat Chart",
    level: "Advanced",
    prompt:
      "Buat chart column dari range A1:B12 dengan judul Penjualan Bulanan di sheet aktif.",
  },
  {
    id: "pivot-summary",
    title: "Pivot Summary Dasar",
    level: "Advanced",
    prompt:
      "Buat pivot_summary dari A1:C1000, groupBy kolom 0, value kolom 2, agregasi sum, outputSheetName RingkasanPenjualan.",
  },
  {
    id: "pivot-summary-topn",
    title: "Pivot Summary + Top N + Chart",
    level: "Automation",
    prompt:
      "Buat pivot_summary dari A1:C5000, groupBy kolom 0, value kolom 2, agregasi sum, sortDescending true, topN 10, outputSheetName TopKategori, createChart true, chartType column, chartTitle Top 10 Kategori.",
  },
  {
    id: "pivot-summary-threshold",
    title: "Pivot Summary + Threshold + Preset",
    level: "Automation",
    prompt:
      "Buat pivot_summary dari A1:C5000, groupBy kolom 0, value kolom 2, agregasi sum, sortDescending true, minValue 100000, topN 12, outputSheetName RingkasanPremium, createChart true, chartType bar, chartPreset executive, chartTitle Kategori Nilai Tinggi.",
  },
  {
    id: "multi-sheet",
    title: "Aksi Multi-Sheet",
    level: "Automation",
    prompt:
      "Terapkan format header tebal dan warna biru muda untuk Sheet1, Sheet2, dan Sheet3 pada range A1:H1.",
  },
];
