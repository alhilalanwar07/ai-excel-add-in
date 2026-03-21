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

    private normalizeFormulaMatrix(formulas: unknown): string[][] {
        if (!Array.isArray(formulas)) {
            throw new Error("Payload formulas harus berupa array.");
        }

        if (formulas.length === 0) {
            throw new Error("Payload formulas tidak boleh kosong.");
        }

        if (Array.isArray(formulas[0])) {
            return (formulas as unknown[]).map((row) => {
                if (!Array.isArray(row)) {
                    throw new Error("Setiap baris formulas harus berupa array string.");
                }
                return row.map((cell) => String(cell));
            });
        }

        return (formulas as unknown[]).map((value) => [String(value)]);
    }

    private buildPivotSummaryRows(
        values: any[][],
        groupByColumn: number,
        valueColumn: number,
        aggregation: string,
        sortDescending: boolean,
        topN?: number,
        minValue?: number
    ): any[][] {
        if (!Array.isArray(values) || values.length === 0) {
            return [["Category", "Value"], ["(no data)", 0]];
        }

        const headerRow = values[0] ?? [];
        if (!Array.isArray(headerRow) || headerRow.length === 0) {
            return [["Category", "Value"], ["(no data)", 0]];
        }

        const safeGroupByColumn = Math.max(0, Math.min(groupByColumn, headerRow.length - 1));
        const safeValueColumn = Math.max(0, Math.min(valueColumn, headerRow.length - 1));
        const dataRows = values.slice(1);
        const groupHeader = String(headerRow[safeGroupByColumn] ?? `Column_${safeGroupByColumn + 1}`);
        const valueHeader = String(headerRow[safeValueColumn] ?? `Column_${safeValueColumn + 1}`);

        if (dataRows.length === 0) {
            return [[groupHeader, `SUM_${valueHeader}`], ["(no data)", 0]];
        }

        const normalizedAggregation = ["sum", "avg", "count"].includes(String(aggregation).toLowerCase())
            ? String(aggregation).toLowerCase()
            : "sum";

        const map = new Map<string, { sum: number; count: number }>();

        for (const row of dataRows) {
            const key = String(row?.[safeGroupByColumn] ?? "(blank)");
            const rawValue = row?.[safeValueColumn];
            const numericValue = typeof rawValue === "number" ? rawValue : Number(rawValue);

            const state = map.get(key) ?? { sum: 0, count: 0 };
            state.count += 1;
            if (!Number.isNaN(numericValue)) {
                state.sum += numericValue;
            }
            map.set(key, state);
        }

        const result: any[][] = [[groupHeader, `${normalizedAggregation.toUpperCase()}_${valueHeader}`]];
        const entries = Array.from(map.entries());
        const effectiveTopN = typeof topN === "number" && topN > 0
            ? Math.min(Math.floor(topN), 5000)
            : null;

        const computedRows = entries.map(([key, state]) => {
            let aggregated: number;
            if (normalizedAggregation === "count") {
                aggregated = state.count;
            } else if (normalizedAggregation === "avg") {
                aggregated = state.count > 0 ? state.sum / state.count : 0;
            } else {
                aggregated = state.sum;
            }
            return [key, aggregated] as [string, number];
        });

        computedRows.sort((a, b) => sortDescending ? b[1] - a[1] : a[1] - b[1]);
        const threshold = typeof minValue === "number" && Number.isFinite(minValue)
            ? Math.max(minValue, 0)
            : null;
        const thresholdRows = threshold !== null
            ? computedRows.filter((row) => row[1] >= threshold)
            : computedRows;
        const finalRows = effectiveTopN ? thresholdRows.slice(0, effectiveTopN) : thresholdRows;

        for (const row of finalRows) {
            result.push(row);
        }

        if (result.length === 1) {
            result.push(["(no rows matched filter)", 0]);
        }

        return result;
    }

    /**
     * Eksekusi spesifik inti dari Office.js (dibungkus untuk dipakai oleh queue)
     */
    private async _executeActionCore(functionName: string, args: any, details?: any): Promise<void> {
      await Excel.run(async (context) => {
        // Ambil snapshot untuk fitur Undo (1 langkah ke belakang) dari sel target utama pertama (fallback safety)
                const snapshotAddress =
                    args.address ||
                    args.startAddress ||
                    (Array.isArray(args.items) && args.items[0]?.address ? args.items[0].address : null);

                if (snapshotAddress) {
                        await this.captureSnapshot(context, snapshotAddress);
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
        const validFunctions = ["write_formula", "bulk_write_formulas", "format_range", "insert_data", "clear_range", "chart", "pivot_summary", "data_manipulation"];
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

                                    case "bulk_write_formulas":
                                        if (args.rangeAddress && args.formulas) {
                                            const bulkRange = this.getSafeRangeWithSheet(context, args.rangeAddress, sheet);
                                            bulkRange.formulas = this.normalizeFormulaMatrix(args.formulas);
                                            bulkRange.format.autofitColumns();
                                            break;
                                        }

                                        if (!Array.isArray(args.items) || args.items.length === 0) {
                                            throw new Error("bulk_write_formulas membutuhkan rangeAddress+formulas atau items[] berisi address dan formula.");
                                        }

                                        for (const item of args.items) {
                                            if (!item?.address || !item?.formula) {
                                                continue;
                                            }

                                            const formulaCell = this.getSafeRangeWithSheet(context, item.address, sheet);
                                            formulaCell.formulas = [[item.formula]];
                                        }
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

                                    case "pivot_summary": {
                                        const sourceRange = this.getSafeRangeWithSheet(context, args.sourceRange, sheet);
                                        sourceRange.load("values");
                                        await context.sync();

                                        const groupByColumn = typeof args.groupByColumn === "number" ? Math.floor(args.groupByColumn) : 0;
                                        const valueColumn = typeof args.valueColumn === "number" ? Math.floor(args.valueColumn) : 1;
                                        const aggregation = typeof args.aggregation === "string" ? args.aggregation : "sum";
                                        const sortDescending = args.sortDescending !== false;
                                        const topN = typeof args.topN === "number" && Number.isFinite(args.topN)
                                            ? Math.max(1, Math.min(Math.floor(args.topN), 5000))
                                            : undefined;
                                        const minValue = typeof args.minValue === "number" && Number.isFinite(args.minValue)
                                            ? Math.max(args.minValue, 0)
                                            : undefined;
                                        const outputSheetName = typeof args.outputSheetName === "string" && args.outputSheetName.trim()
                                            ? args.outputSheetName.trim()
                                            : "PivotSummary";

                                        const summaryRows = this.buildPivotSummaryRows(
                                            sourceRange.values,
                                            groupByColumn,
                                            valueColumn,
                                            aggregation,
                                            sortDescending,
                                            topN,
                                            minValue
                                        );

                                        let outputSheet = worksheets.getItemOrNullObject(outputSheetName);
                                        outputSheet.load("isNullObject");
                                        await context.sync();
                                        if (outputSheet.isNullObject) {
                                            outputSheet = worksheets.add(outputSheetName);
                                        }

                                        const outputStart = outputSheet.getRange("A1");
                                        const outputRange = outputStart.getResizedRange(summaryRows.length - 1, summaryRows[0].length - 1);
                                        outputRange.values = summaryRows;
                                        outputRange.format.autofitColumns();

                                        if (args.createChart && summaryRows.length > 1) {
                                            const chartDataRange = outputSheet.getRangeByIndexes(0, 0, summaryRows.length, summaryRows[0].length);
                                            const chartTypeRaw = String(args.chartType ?? "column").toLowerCase();
                                            const chartPreset = String(args.chartPreset ?? "compact").toLowerCase();
                                            let chartType: Excel.ChartType = Excel.ChartType.columnClustered;
                                            if (chartTypeRaw === "pie") chartType = Excel.ChartType.pie;
                                            if (chartTypeRaw === "line") chartType = Excel.ChartType.line;
                                            if (chartTypeRaw === "bar") chartType = Excel.ChartType.barClustered;

                                            const summaryChart = outputSheet.charts.add(chartType, chartDataRange, Excel.ChartSeriesBy.columns);
                                            if (chartPreset === "presentation") {
                                                summaryChart.setPosition("D2", "N24");
                                            } else if (chartPreset === "executive") {
                                                summaryChart.setPosition("D2", "M20");
                                            } else {
                                                summaryChart.setPosition("D2", "L18");
                                            }
                                            summaryChart.title.visible = true;
                                            summaryChart.title.text = typeof args.chartTitle === "string" && args.chartTitle.trim()
                                                ? args.chartTitle.trim()
                                                : "Pivot Summary Chart";
                                        }
                                        break;
                                    }

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
