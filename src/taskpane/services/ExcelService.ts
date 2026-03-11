/**
 * Service singleton untuk mengelola interaksi rumit dengan Excel (Office.js)
 * Enterprise Features: Self-Correction Execution, Undo Snapshot (1 step), 
 * dan Queueing System yang mencegah "Conflict/Race Condition" saat eksekusi beruntun.
 */

export class ExcelService {
    private static instance: ExcelService;
    
    // Menyimpan state dari sel sebelum tindakan destruktif diubah (untuk Undo Layer)
    private __lastSnapshot: { address: string; values: any[][]; formulas: any[][], sheetName: string } | null = null;
    
    // Sistem Antrean (Queue)
    private actionQueue: Array<() => Promise<void>> = [];
    private isExecutingQueue = false;

    private constructor() {}
  
    public static getInstance(): ExcelService {
      if (!ExcelService.instance) {
        ExcelService.instance = new ExcelService();
      }
      return ExcelService.instance;
    }
  
    /**
     * Helper untuk memparsing range string yang mungkin mengandung alias Sheet
     */
    private getSafeRange(context: Excel.RequestContext, addressString: string) {
      if (!addressString || typeof addressString !== 'string') {
          return context.workbook.worksheets.getActiveWorksheet().getRange();
      }
      const parts = addressString.split('!');
      if (parts.length === 2) {
         const sheetName = parts[0].replace(/'/g, ''); 
         const rangeAddress = parts[1];
         return context.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
      }
      return context.workbook.worksheets.getActiveWorksheet().getRange(addressString);
    }
  
    /**
     * Mempersiapkan state untuk Undo. Dipanggil SEBELUM eksekusi action.
     */
    private async captureSnapshot(context: Excel.RequestContext, addressString: string) {
        try {
            const range = this.getSafeRange(context, addressString);
            range.load(["address", "values", "formulas"]);
            const worksheet = range.worksheet;
            worksheet.load("name");
            await context.sync();
            
            this.__lastSnapshot = {
                address: range.address,
                values: range.values,
                formulas: range.formulas,
                sheetName: worksheet.name
            };
        } catch (e) {
            console.warn("Gagal membuat pre-snapshot untuk Undo. Undo mungkin tidak tersedia.", e);
        }
    }
  
    /**
     * Mengeksekusi aksi Excel ke dalam Queue (Antrean)
     */
    public async executeAction(functionName: string, args: any, details?: any): Promise<void> {
        return new Promise((resolve, reject) => {
            // Tambahkan tugas ke antrean
            this.actionQueue.push(async () => {
                try {
                    await this._executeActionCore(functionName, args, details);
                    resolve();
                } catch (error) {
                    reject(error);
                }
            });

            // Picu pemrosesan antrean
            this.processQueue();
        });
    }

    /**
     * Pemroses Antrean internal (Hanya 1 aksi yang berjalan pada satu waktu)
     */
    private async processQueue() {
        if (this.isExecutingQueue || this.actionQueue.length === 0) return;
        
        this.isExecutingQueue = true;
        
        while (this.actionQueue.length > 0) {
            const actionTask = this.actionQueue.shift();
            if (actionTask) {
                try {
                    await actionTask();
                } catch (e) {
                    console.error("Kesalahan saat mengeksekusi aksi dalam sistem queue Excel:", e);
                    // Lanjutkan queue berikutnya meskipun yang satu ini gagal,
                    // atau bisa juga diberhentikan tergantung kebijakan. Kita biarkan berlanjut.
                }
            }
        }
        
        this.isExecutingQueue = false;
    }
  
    private getSafeRangeWithSheet(context: Excel.RequestContext, addressString: string, currentSheet: Excel.Worksheet) {
        if (!addressString || typeof addressString !== 'string') {
             throw new Error("Target address (range) tidak disertakan di dalam execution_payload oleh AI.");
        }
        
        const parts = addressString.split('!');
        if (parts.length === 2) {
             const sheetName = parts[0].replace(/'/g, ''); 
             const rangeAddress = parts[1];
             return context.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
        }
        return currentSheet.getRange(addressString);
    }

    /**
     * Eksekusi spesifik inti dari Office.js (dibungkus untuk dipakai oleh queue)
     */
    private async _executeActionCore(functionName: string, args: any, details?: any): Promise<void> {
      await Excel.run(async (context) => {
        // Ambil snapshot untuk fitur Undo (1 langkah ke belakang) dari sel target utama pertama (fallback safety)
        if (args.address || args.startAddress) {
            await this.captureSnapshot(context, args.address || args.startAddress);
        }

        const worksheets = context.workbook.worksheets;
        worksheets.load("items/name");
        await context.sync();

        // 1. EVALUASI TARGET SCOPE LOGIC (ENTERPRISE MULTI-SHEET)
        let targetSheetNames: string[] = [];
        if (details?.target_scope === "all_sheets") {
             targetSheetNames = worksheets.items.map(s => s.name);
        } else if (details?.target_scope === "specific_sheets" && Array.isArray(details.sheet_names)) {
             targetSheetNames = details.sheet_names;
        } else {
             targetSheetNames = [worksheets.getActiveWorksheet().name]; // fallback: active_sheet
        }

        // 2. ITERASI KE MASING-MASING SHEET TARGET
        const validFunctions = ["write_formula", "format_range", "insert_data", "clear_range", "chart", "data_manipulation"];
        if (!validFunctions.includes(functionName)) {
             throw new Error(`Fungsi eksekusi '${functionName}' belum diimplementasikan atau tidak diperbolehkan. Gunakan aksi lain!`);
        }

        let isTargetSheetFound = false;
        let actionErrors: string[] = [];

        for (const sheetName of targetSheetNames) {
            try {
                // Enterprise Feature: Auto-Create Sheet jika belum ada
                let sheet = worksheets.getItemOrNullObject(sheetName);
                sheet.load("isNullObject");
                await context.sync();
                
                if (sheet.isNullObject) {
                    sheet = worksheets.add(sheetName);
                }
                
                isTargetSheetFound = true;
                switch (functionName) {
                  case "write_formula":
                    const targetCell = this.getSafeRangeWithSheet(context, args.address, sheet);
                    targetCell.formulas = [[args.formula]]; 
                    targetCell.format.autofitColumns();
                    break;
          
                  case "format_range":
                    const formatRange = this.getSafeRangeWithSheet(context, args.address, sheet);
                    if (args.backgroundColor) formatRange.format.fill.color = args.backgroundColor;
                    if (args.bold !== undefined) formatRange.format.font.bold = args.bold;
                    break;
          
                  case "insert_data":
                    const startRange = this.getSafeRangeWithSheet(context, args.startAddress, sheet);
                    const dataRange = startRange.getResizedRange(args.dataValues.length - 1, args.dataValues[0].length - 1);
                    dataRange.values = args.dataValues;
                    dataRange.format.autofitColumns();
                    break;
          
                  case "clear_range":
                    const rangeToClear = this.getSafeRangeWithSheet(context, args.address, sheet);
                    rangeToClear.clear();
                    break;
          
                  case "chart":
                    const chartDataRange = this.getSafeRangeWithSheet(context, args.data_range, sheet);
                    
                    // Menyesuaikan tipe grafik dengan enum Office.js
                    let chartTypeMapping: Excel.ChartType = Excel.ChartType.columnClustered;
                    if (args.chart_type === "line") chartTypeMapping = Excel.ChartType.line;
                    if (args.chart_type === "pie") chartTypeMapping = Excel.ChartType.pie;
                    if (args.chart_type === "bar") chartTypeMapping = Excel.ChartType.barClustered;

                    let targetChartSheet = sheet;
                    if (args.insert_position === "new_sheet") {
                        const newSheetName = "Report_" + Math.random().toString(36).substring(7);
                        targetChartSheet = worksheets.add(newSheetName);
                    }

                    const chart = targetChartSheet.charts.add(chartTypeMapping, chartDataRange, Excel.ChartSeriesBy.auto);
                    if (args.title) {
                        chart.title.text = args.title;
                        chart.title.visible = true;
                    }
                    break;

                  case "data_manipulation":
                    const manipulationRange = this.getSafeRangeWithSheet(context, args.range, sheet);
                    
                    if (args.operation === "remove_duplicates") {
                        // Secara asali menganggap header di baris pertama
                        manipulationRange.removeDuplicates([0], true); 
                    } else if (args.operation === "sort") {
                        const sortFields: Excel.SortField[] = [];
                        if (args.criteria && args.criteria.columns) {
                            for(let c of args.criteria.columns) {
                                // Default sorting ascending jika true / undefined
                                sortFields.push({ key: c.index || 0, ascending: c.ascending !== false }); 
                            }
                        } else {
                            sortFields.push({ key: 0, ascending: true }); // Default kolom pertama asceding
                        }
                        manipulationRange.sort.apply(sortFields, true);
                    } else if (args.operation === "filter") {
                         // Terapkan autofilter simpel di range tersebut melalui objek Worksheet
                         sheet.autoFilter.apply(manipulationRange);
                    } else if (args.operation === "delete") {
                         // Mengekstrak alamat range khusus untuk penghapusan
                         if (args.target === "sheet") {
                              sheet.delete();
                         } else {
                              manipulationRange.delete(Excel.DeleteShiftDirection.up);
                         }
                    }
                    break;
          
                }
            } catch (err: any) {
                 console.error(`Gagal operasi Multi-Sheet pada lembar '${sheetName}':`, err);
                 actionErrors.push(`Sheet '${sheetName}': ${err.message}`);
            }
        }
        
        if (!isTargetSheetFound && targetSheetNames.length > 0) {
            throw new Error(`Lembar dengan nama ${targetSheetNames.join(', ')} tidak ditemukan.`);
        }

        if (actionErrors.length > 0) {
            throw new Error(`Operasi gagal sebagian atau seluruhnya pada sheet: \n${actionErrors.join('\n')}`);
        }
  
        // Eksekusi antrean Office.js
        await context.sync();
      });
    }
  
    /**
     * Mengembalikan sel ke kondisi sebelum eksekusi AI terakhir kali
     */
    public async undoLastAction(): Promise<boolean> {
        if (!this.__lastSnapshot) {
            return false;
        }
  
        try {
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getItem(this.__lastSnapshot!.sheetName);
                const range = sheet.getRange(this.__lastSnapshot!.address);
                
                // Restore values (atau formulas kalau itu adalah rumus)
                // Catatan: Jika cell aslinya rumus, kita harus memulihkan formula, jika bukan kembalikan value.
                if (this.__lastSnapshot!.formulas && this.__lastSnapshot!.formulas[0][0].startsWith('=')) {
                    range.formulas = this.__lastSnapshot!.formulas;
                } else {
                    range.values = this.__lastSnapshot!.values;
                }
                
                await context.sync();
            });
            // Kosongkan snapshot setelah dipakai
            this.__lastSnapshot = null;
            return true;
        } catch (e) {
            console.error("Gagal melakukan undo", e);
            throw new Error("Gagal melakukan Undo Office.js");
        }
    }
  }
  
  export const excelService = ExcelService.getInstance();
