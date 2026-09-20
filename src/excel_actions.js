/**
 * Excel Live Action Engine (Office.js)
 * Interacts directly with the active Excel workbook:
 * reads selections, writes values/formulas, styles ranges,
 * and maintains an undo snapshot stack.
 *
 * Bug fixes applied:
 *  4  - Freeze activeSheet name before any await in executeActions
 *  5  - Atomic backup (delete+rethrow on copy failure)
 *  8  - Hash-based backup name to prevent long-name collisions
 *  9  - Strengthen backup-prefix guard
 * 10/11 - Structural undo stores explicit backupName; repush on failure
 * 12  - find_replace preserves formula cells
 * 13  - addTotalRow uses usedRange.rowIndex / columnIndex
 * 14  - fill_formula fallback adjusts row references
 * 15  - Transactional batch: rollback on partial failure
 * 16  - validateAction() schema layer
 * 20/21/22 - Debounced reads, sequence guard, cancellation counter
 * 26  - Action preview in non-auto-apply mode
 * 27  - create_pivot_table: selective pivot deletion (not all)
 * 28  - Pivot: warn when dest area non-empty
 * 29/30 - Structural snapshot for create_pivot + create_chart
 * 31  - Table undo deletes Table object
 * 32/33/34/35 - Borders/merge/colWidth/rowHeight in snapshot + undo
 * 36  - add_sheet: use getItemOrNullObject; only snapshot if created
 * 37  - find_replace: preserve numeric types
 * 48  - Strip sheet prefix from backup address (local address only)
 * 49/51 - Remove txtGeminiKey + dead quickAction paths
 */

const ExcelBridge = {
    isOfficeInitialized: false,
    undoStack: [],
    redoStack: [],
    _chartRegistry: {},
    _createdChartsThisBatch: [],

    // Bug 20: Cancellation counter for getFullSheetData
    _readGeneration: 0,
    // Bug 22: Monotone sequence to discard stale selection reads
    _selReadSeq: 0,
    // Bug 21: Debounce handle
    _selDebounceTimer: null,

    /**
     * Column helpers — 0-based index <-> A1 letters, correct past Z (AA, AB, ...).
     */
    columnLetter(index) {
        let s = "";
        let n = index + 1;
        while (n > 0) {
            const m = (n - 1) % 26;
            s = String.fromCharCode(65 + m) + s;
            n = Math.floor((n - 1) / 26);
        }
        return s;
    },

    columnIndex(letters) {
        const m = /^([A-Za-z]+)/.exec(String(letters || ""));
        if (!m) return 0;
        let n = 0;
        for (const ch of m[1].toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
        return n - 1;
    },

    /** Accepts 0-based number or "C"-style letters, returns 0-based index. */
    resolveColumn(col) {
        return (typeof col === "number") ? col : this.columnIndex(col);
    },

    /**
     * Bug 48: Strip the sheet-name prefix from a sheet-qualified address.
     * Excel's Range.address returns "Sheet1!A1:B4"; backup.getRange() needs
     * only the local portion "A1:B4".
     */
    localAddress(addr) {
        if (!addr) return addr;
        const bang = String(addr).lastIndexOf("!");
        return bang >= 0 ? addr.slice(bang + 1) : addr;
    },

    /**
     * Bug 8: Hash-based backup name so two long sheet names with the same
     * 31-char prefix don't collide. Uses a simple djb2-style 16-bit hash.
     */
    _shortHash(str) {
        let h = 5381;
        for (let i = 0; i < str.length; i++) h = ((h << 5) + h) ^ str.charCodeAt(i);
        return (h >>> 0).toString(36).slice(0, 4).padStart(4, "0");
    },

    /** Push a restorable snapshot (sheet-qualified) onto the undo stack. */
    pushSnapshot(sheetName, address, values, formulas, numberFormat, kind, extra) {
        const entry = Object.assign({
            kind: kind || "range",
            sheet: sheetName,
            range: address,
            values: values ? JSON.parse(JSON.stringify(values)) : null,
            formulas: formulas ? JSON.parse(JSON.stringify(formulas)) : null,
            numberFormat: numberFormat ? JSON.parse(JSON.stringify(numberFormat)) : null
        }, extra || {});
        this.undoStack.push(entry);
        if (this.undoStack.length > 100) {
            this.undoStack.splice(0, this.undoStack.length - 100);
        }
        this.persistUndoMeta();
        return entry;
    },

    persistUndoMeta() {
        try {
            const recs = Object.values(this._backupRecords || {})
                .filter(r => r && r.sheet)
                .sort((a, b) => (a.time || 0) - (b.time || 0))
                .slice(-10);
            const meta = {
                stack: this.undoStack
                    .filter(e => e && e.kind && e.kind !== "range")
                    .slice(-20),
                lastBackup: this.lastBackup || null,
                backups: recs,
                chartRegistry: this._chartRegistry || {}
            };
            localStorage.setItem("excelAiUndoMeta", JSON.stringify(meta));
        } catch (_) {}
    },

    restoreUndoMeta() {
        try {
            const meta = JSON.parse(localStorage.getItem("excelAiUndoMeta") || "null");
            if (!meta) return;
            const kept = Array.isArray(meta.stack) ? meta.stack : [];
            if (kept.length > 0) {
                this.undoStack = kept.concat(this.undoStack).slice(-100);
            }
            if (!this.lastBackup && meta.lastBackup && meta.lastBackup.sheet) {
                this.lastBackup = meta.lastBackup;
            }
            if (Array.isArray(meta.backups)) {
                for (const r of meta.backups) {
                    if (r && r.sheet && r.backup) this._backupRecords[r.sheet] = r;
                }
            }
            if (meta.chartRegistry && typeof meta.chartRegistry === "object") {
                this._chartRegistry = Object.assign({}, meta.chartRegistry, this._chartRegistry);
            }
        } catch (_) {}
    },

    BACKUP_PREFIX: "AI_Backup__",
    lastBackup: null,
    _backupMap: {},
    _backupNotes: [],
    _backupRecords: {},
    _lastBatchBackups: [],

    sanitizeSheetName(name) {
        return String(name || "Sheet").replace(/[\[\]\*\/\\\?:]/g, "").substring(0, 31) || "Sheet";
    },

    /**
     * Bug 8: Append a 4-char hash so two long sheet names can't collide after
     * truncation to Excel's 31-character worksheet name limit.
     */
    backupNameFor(sheetName) {
        const base = String(sheetName || "Sheet").replace(/[\[\]\*\/\\\?:]/g, "");
        const hash = this._shortHash(String(sheetName));
        const prefix = this.BACKUP_PREFIX; // "AI_Backup__" = 11 chars
        // Reserve 4 chars for hash + 1 separator = 5; prefix = 11; total budget = 31
        const maxBase = 31 - prefix.length - 5;
        return (prefix + base.substring(0, maxBase) + "_" + hash).substring(0, 31);
    },

    isBackupSheet(name) {
        return String(name || "").indexOf(this.BACKUP_PREFIX) === 0;
    },

    /**
     * Copy a sheet's used-range content into a hidden backup sheet.
     * Bug 5: Atomic — backup sheet is deleted and error re-thrown if copy fails.
     * Bug 9: Refuses to back up sheets whose name starts with BACKUP_PREFIX.
     * Bug 48: Uses localAddress() to strip sheet qualifier from used.address.
     * Must be called inside an Excel.run with a live context.
     */
    async ensureBackupSheet(context, sheetName) {
        // Bug 9: Never back up a backup sheet itself.
        if (this.isBackupSheet(sheetName)) {
            throw new Error(`Sheet "${sheetName}" looks like a backup sheet — skipping.`);
        }
        const backupName = this.backupNameFor(sheetName);
        const sheets = context.workbook.worksheets;
        try {
            const old = sheets.getItemOrNullObject(backupName);
            await context.sync();
            if (!old.isNullObject) {
                old.delete();
                await context.sync();
            }
        } catch (_) {}
        const src = sheets.getItem(sheetName);
        const backup = sheets.add(backupName);
        backup.visibility = "Hidden";
        let copySucceeded = false;
        try {
            const used = src.getUsedRange();
            used.load(["address"]);
            await context.sync();
            // Bug 48: Use only the local (sheet-unqualified) address.
            const localAddr = this.localAddress(used.address);
            const dest = backup.getRange(localAddr);
            try {
                dest.copyFrom(used, Excel.RangeCopyType.all);
            } catch (_) {
                // Older builds: fall back to values + formulas.
                used.load(["values", "formulas"]);
                await context.sync();
                dest.values = used.values;
                dest.formulas = used.formulas;
            }
            await context.sync();
            copySucceeded = true;
        } catch (copyErr) {
            // Bug 5: Copy failed — delete the partial backup and re-throw.
            try { backup.delete(); await context.sync(); } catch (_) {}
            throw new Error(`Backup copy failed for "${sheetName}": ${copyErr.message || copyErr}`);
        }
        // Full record: table + chart definitions
        const rec = { sheet: sheetName, backup: backupName, time: Date.now(), charts: [], tables: [] };
        try {
            const chColl = src.charts;
            chColl.load("items/name");
            await context.sync();
            for (const ch of chColl.items) {
                try {
                    const d = await this.captureChartDef(context, src, ch);
                    d.sheet = sheetName;
                    rec.charts.push(d);
                } catch (_) {}
            }
        } catch (_) {}
        try {
            const tColl = src.tables;
            tColl.load("items/name");
            await context.sync();
            const tRanges = [];
            for (const t of tColl.items) {
                try {
                    t.load(["name", "style", "showHeaders", "showTotals"]);
                    const tr = t.getRange();
                    tr.load("address");
                    tRanges.push({ t: t, tr: tr });
                } catch (_) {}
            }
            await context.sync();
            for (const pair of tRanges) {
                try {
                    rec.tables.push({
                        name: pair.t.name, address: pair.tr.address, style: pair.t.style,
                        hasHeaders: pair.t.showHeaders !== false,
                        showTotals: Boolean(pair.t.showTotals)
                    });
                } catch (_) {}
            }
        } catch (_) {}
        // Bug 5: Only record success AFTER copy confirmed.
        this._backupRecords[sheetName] = rec;
        this._lastBatchBackups.push(sheetName);
        this.lastBackup = { sheet: sheetName, backup: backupName, time: Date.now() };
        this.persistUndoMeta();
        return backupName;
    },

    /**
     * Rebuild tables + charts on worksheet from backup-record definitions.
     */
    async rebuildSheetObjects(context, sheetName, rec, preservedChartNames) {
        rec = rec || {};
        const preserved = preservedChartNames || new Set();
        const ws = context.workbook.worksheets.getItem(sheetName);
        for (const t of (rec.tables || [])) {
            try {
                const nt = ws.tables.add(t.address, t.hasHeaders !== false);
                if (t.name) {
                    try { nt.name = String(t.name).replace(/[^a-zA-Z0-9_]/g, ""); } catch (_) {}
                }
                if (t.style) { try { nt.style = t.style; } catch (_) {} }
                if (t.showTotals !== undefined) {
                    try { nt.showTotals = Boolean(t.showTotals); } catch (_) {}
                }
            } catch (_) {}
        }
        try { await context.sync(); } catch (_) {}
        for (const d of (rec.charts || [])) {
            if (!d) continue;
            try {
                d.sheet = sheetName;
                if (d.name && preserved.has(d.name)) {
                    // Pre-existing chart still present — only revert properties if needed, never delete or re-add
                    const ch = ws.charts.getItemOrNullObject(d.name);
                    await context.sync();
                    if (!ch.isNullObject) {
                        if (d.title) {
                            try { ch.title.text = d.title; ch.title.visible = d.titleVisible !== false; } catch (_) {}
                        }
                        if (d.style !== undefined) {
                            try { ch.style = d.style; } catch (_) {}
                        }
                        if (d.chartType) {
                            try { ch.chartType = d.chartType; } catch (_) {}
                        }
                    }
                } else {
                    // Chart is missing from sheet (was deleted during batch) — restore it!
                    await this.restoreChartFromDef(context, sheetName, d);
                }
            } catch (_) {}
        }
    },

    /**
     * Unified rollback of sheet `sheetName` to its hidden backup.
     * IN-PLACE whenever the sheet still exists (preserves the sheet
     * object, so cross-sheet references never break).
     */
    async rollbackSheetToBackup(context, sheetName, rec) {
        rec = rec || {};
        const sheets = context.workbook.worksheets;
        const backupName = rec.backup || this.backupNameFor(sheetName);
        const backup = sheets.getItemOrNullObject(backupName);
        await context.sync();
        if (backup.isNullObject) {
            throw new Error(`No backup found for "${sheetName}" (looked for hidden sheet "${backupName}").`);
        }
        const cur = sheets.getItemOrNullObject(sheetName);
        await context.sync();
        if (cur.isNullObject) {
            backup.visibility = "Visible";
            backup.name = this.sanitizeSheetName(sheetName);
            await context.sync();
            await this.rebuildSheetObjects(context, sheetName, rec);
            return `Restored deleted sheet "${sheetName}" from "${backupName}" (values, formats, tables, charts).`;
        }
        try {
            const liveTables = cur.tables;
            liveTables.load("items/name");
            await context.sync();
            for (const t of liveTables.items) {
                try { t.delete(); } catch (_) {}
            }
            await context.sync();
        } catch (_) {}

        // Selective chart management: DO NOT delete pre-existing charts!
        // Only delete charts that were newly added by the batch being rolled back.
        const backupChartDefs = rec.charts || [];
        const backupChartNames = new Set(backupChartDefs.map(c => c && c.name).filter(Boolean));
        const createdNow = new Set(this._createdChartsThisBatch || []);
        const preservedChartNames = new Set();

        try {
            const liveCharts = cur.charts;
            liveCharts.load("items/name");
            await context.sync();
            for (const c of liveCharts.items) {
                // If created in this batch, delete it:
                const wasCreatedThisBatch = createdNow.has(c.name) ||
                    (backupChartNames.size > 0 && !backupChartNames.has(c.name));
                if (wasCreatedThisBatch) {
                    try { c.delete(); } catch (_) {}
                } else {
                    // Pre-existing chart: KEEP IT!
                    preservedChartNames.add(c.name);
                }
            }
            await context.sync();
        } catch (_) {}

        try {
            cur.getUsedRange().clear("All");
            await context.sync();
        } catch (_) {}
        let bkr = 0, bkc = 0;
        try {
            const bu = backup.getUsedRange();
            bu.load(["address", "rowCount", "columnCount"]);
            await context.sync();
            bkr = bu.rowCount; bkc = bu.columnCount;
            if (bkr > 0 && bkc > 0) {
                // Bug 48: Use local address for copy destination.
                const localAddr = this.localAddress(bu.address);
                cur.getRange(localAddr).copyFrom(bu, Excel.RangeCopyType.all);
                await context.sync();
            }
        } catch (e) {
            throw new Error(`Backup "${backupName}" is unreadable: ${e.message || e}`);
        }
        try {
            const cu = cur.getUsedRange();
            cu.load(["rowCount", "columnCount"]);
            await context.sync();
            if (cu.rowCount > bkr) {
                cur.getRangeByIndexes(bkr, 0, cu.rowCount - bkr, 1)
                    .getEntireRow().delete(Excel.DeleteShiftDirection.up);
            }
            if (cu.columnCount > bkc) {
                cur.getRangeByIndexes(0, bkc, 1, cu.columnCount - bkc)
                    .getEntireColumn().delete(Excel.DeleteShiftDirection.left);
            }
            await context.sync();
        } catch (_) {}
        await this.rebuildSheetObjects(context, sheetName, rec, preservedChartNames);
        const parts = [];
        parts.push(`${(rec.tables || []).length} table(s)`);
        const totalCharts = preservedChartNames.size + (rec.charts || []).filter(c => !preservedChartNames.has(c.name)).length;
        parts.push(`${totalCharts} chart(s)`);
        return `Rolled "${sheetName}" fully back to its backup (values, formats, ${parts.join(", ")}).`;
    },

    /**
     * Capture a chart's full definition (plain JSON-safe data).
     */
    async captureChartDef(context, ws, chart) {
        const def = { series: [], axes: {} };
        try {
            chart.load(["name", "chartType", "top", "left", "height", "width"]);
            await context.sync();
            def.name = chart.name;
            def.chartType = chart.chartType;
            def.top = chart.top; def.left = chart.left;
            def.height = chart.height; def.width = chart.width;
        } catch (_) {}

        // Check registry for previously recorded creation info (sourceRange, position, title, etc.)
        const reg = (this._chartRegistry && def.name && this._chartRegistry[def.name]) || null;
        if (reg) {
            if (reg.sourceRange) def.sourceRange = reg.sourceRange;
            if (reg.position) def.position = reg.position;
            if (reg.seriesBy) def.seriesBy = reg.seriesBy;
            if (reg.title) def.title = reg.title;
            if (reg.style !== undefined) def.style = reg.style;
        }

        try {
            chart.load("style");
            await context.sync();
            def.style = chart.style;
        } catch (_) {}

        try {
            chart.title.load(["visible", "text"]);
            await context.sync();
            if (chart.title.visible) {
                def.title = chart.title.text;
                def.titleVisible = true;
            } else {
                def.titleVisible = false;
            }
        } catch (_) {}

        try {
            chart.legend.load(["visible", "position"]);
            await context.sync();
            def.legendVisible = chart.legend.visible;
            def.legendPosition = chart.legend.position;
        } catch (_) {}

        try {
            chart.dataLabels.load("showValue");
            await context.sync();
            def.showValue = chart.dataLabels.showValue;
        } catch (_) {}

        // Safely inspect series data sources if ExcelApi 1.15+ is available
        try {
            chart.series.load("items/name");
            await context.sync();
            for (const s of chart.series.items) {
                const sd = { name: s.name };
                try {
                    s.load("chartType");
                    await context.sync();
                    sd.chartType = s.chartType;
                } catch (_) {}

                if (typeof s.getDimensionDataSourceString === "function") {
                    try {
                        const vsDim = (typeof Excel !== "undefined" && Excel.ChartSeriesDimension && Excel.ChartSeriesDimension.values)
                            ? Excel.ChartSeriesDimension.values : "Values";
                        const vs = s.getDimensionDataSourceString(vsDim);
                        await context.sync();
                        if (vs && vs.value) sd.valuesSource = vs.value;
                    } catch (_) {
                        try {
                            const vs = s.getDimensionDataSourceString("values");
                            await context.sync();
                            if (vs && vs.value) sd.valuesSource = vs.value;
                        } catch (_) {}
                    }
                    try {
                        const csDim = (typeof Excel !== "undefined" && Excel.ChartSeriesDimension && Excel.ChartSeriesDimension.categories)
                            ? Excel.ChartSeriesDimension.categories : "Categories";
                        const cs = s.getDimensionDataSourceString(csDim);
                        await context.sync();
                        if (cs && cs.value) sd.catSource = cs.value;
                    } catch (_) {
                        try {
                            const cs = s.getDimensionDataSourceString("categories");
                            await context.sync();
                            if (cs && cs.value) sd.catSource = cs.value;
                        } catch (_) {}
                    }
                }
                def.series.push(sd);
            }
        } catch (_) {}

        // Only query axes for chart types that support them (not pie, doughnut)
        const isPieOrDoughnut = /pie|doughnut/i.test(def.chartType || "");
        if (!isPieOrDoughnut) {
            for (const pair of [["category", "categoryAxis"], ["value", "valueAxis"]]) {
                try {
                    const ax = chart.axes[pair[1]];
                    ax.load(["majorUnit", "minorUnit", "numberFormat"]);
                    await context.sync();
                    const a = { majorUnit: ax.majorUnit, minorUnit: ax.minorUnit, numberFormat: ax.numberFormat };
                    try {
                        ax.title.load(["visible", "text"]);
                        await context.sync();
                        if (ax.title.visible) {
                            a.title = ax.title.text;
                            a.titleVisible = true;
                        }
                    } catch (_) {}
                    try {
                        ax.majorGridlines.load("visible");
                        await context.sync();
                        a.majorGrid = ax.majorGridlines.visible;
                    } catch (_) {}
                    try {
                        ax.minorGridlines.load("visible");
                        await context.sync();
                        a.minorGrid = ax.minorGridlines.visible;
                    } catch (_) {}
                    def.axes[pair[0]] = a;
                } catch (_) {}
            }
        }

        return def;
    },

    parseChartSource(str) {
        const m = /^=?'?(.*?)'?!([\s\S]+)$/.exec(String(str || "").trim());
        return m ? { sheet: m[1], address: m[2] } : null;
    },

    resolveChartRange(context, str, fallbackSheet) {
        const p = this.parseChartSource(str);
        if (!p || !p.address) {
            try {
                return fallbackSheet.getRange(String(str).trim());
            } catch (_) {
                return null;
            }
        }
        try {
            const ws = p.sheet ? context.workbook.worksheets.getItem(p.sheet) : fallbackSheet;
            return ws.getRange(p.address);
        } catch (_) {
            return null;
        }
    },

    applyAxisState(chart, which, a) {
        const ax = which.indexOf("val") === 0 ? chart.axes.valueAxis
            : which.indexOf("ser") === 0 ? chart.axes.seriesAxis
            : chart.axes.categoryAxis;
        if (a.majorUnit !== undefined && a.majorUnit !== null) {
            try { ax.majorUnit = a.majorUnit; } catch (_) {}
        }
        if (a.minorUnit !== undefined && a.minorUnit !== null) {
            try { ax.minorUnit = a.minorUnit; } catch (_) {}
        }
        if (a.numberFormat) { try { ax.numberFormat = a.numberFormat; } catch (_) {} }
        if (a.title) {
            try { ax.title.text = a.title; ax.title.visible = a.titleVisible !== false; } catch (_) {}
        }
        if (a.majorGrid !== undefined) {
            try { ax.majorGridlines.visible = Boolean(a.majorGrid); } catch (_) {}
        }
        if (a.minorGrid !== undefined) {
            try { ax.minorGridlines.visible = Boolean(a.minorGrid); } catch (_) {}
        }
    },

    async restoreChartFromDef(context, sheetName, def) {
        if (!def) return null;
        const ws = context.workbook.worksheets.getItem(sheetName);

        // 1. Resolve source range
        let srcRange = null;
        const reg = (this._chartRegistry || {})[def.name] || {};
        const sourceAddr = def.sourceRange || reg.sourceRange || def.dataRangeAddress;
        if (sourceAddr) {
            srcRange = this.resolveChartRange(context, sourceAddr, ws);
        }
        if (!srcRange && def.series && def.series.length > 0) {
            const first = def.series[0];
            if (first && first.valuesSource) {
                srcRange = this.resolveChartRange(context, first.valuesSource, ws);
            }
        }
        if (!srcRange) {
            try {
                const tables = ws.tables;
                tables.load("items");
                await context.sync();
                if (tables.items.length > 0) {
                    srcRange = tables.items[0].getRange();
                }
            } catch (_) {}
        }
        if (!srcRange) {
            try {
                const ur = ws.getUsedRange();
                ur.load("address");
                await context.sync();
                srcRange = ws.getRange(this.localAddress(ur.address));
            } catch (_) {
                srcRange = ws.getRange("A1");
            }
        }

        const chartType = def.chartType || reg.chartType || "ColumnClustered";
        const seriesByStr = def.seriesBy || reg.seriesBy || "columns";
        const seriesBy = String(seriesByStr).toLowerCase() === "rows"
            ? Excel.ChartSeriesBy.rows : Excel.ChartSeriesBy.columns;

        const chart = ws.charts.add(chartType, srcRange, seriesBy);
        await context.sync();

        // Position: prefer position string (e.g. "E2:K20"), then top/left
        const posStr = def.position || reg.position;
        if (posStr) {
            try {
                const pos = String(posStr).split(":");
                chart.setPosition(pos[0] || "E2", pos[1] || null);
            } catch (_) {}
        } else if (def.top !== undefined && def.left !== undefined) {
            try {
                chart.top = def.top;
                chart.left = def.left;
                if (def.height) chart.height = def.height;
                if (def.width) chart.width = def.width;
            } catch (_) {}
        }

        // Title
        if (def.title || reg.title) {
            try {
                chart.title.text = String(def.title || reg.title);
                chart.title.visible = def.titleVisible !== false;
            } catch (_) {}
        }

        // Legend
        if (def.legendVisible !== undefined || reg.legend !== undefined) {
            try {
                const legVis = def.legendVisible !== undefined ? def.legendVisible : reg.legend;
                chart.legend.visible = Boolean(legVis);
                const legPos = def.legendPosition || reg.legendPosition;
                if (legPos) {
                    const lpMap = { top: "Top", bottom: "Bottom", left: "Left", right: "Right", topright: "TopRight" };
                    chart.legend.position = lpMap[String(legPos).toLowerCase()] || legPos;
                }
            } catch (_) {}
        }

        // Data labels
        if (def.showValue || reg.dataLabels) {
            try { chart.dataLabels.showValue = true; } catch (_) {}
        }

        // Style
        const st = def.style !== undefined ? def.style : reg.style;
        if (st !== undefined) {
            try { chart.style = st; } catch (_) {}
        }

        // Axes (non-pie/doughnut)
        const isPieOrDoughnut = /pie|doughnut/i.test(chartType);
        if (!isPieOrDoughnut) {
            if (def.axes) {
                try {
                    for (const k of ["category", "value"]) {
                        if (def.axes[k]) this.applyAxisState(chart, k, def.axes[k]);
                    }
                } catch (_) {}
            }
            if (reg.valueAxisTitle) {
                try { chart.axes.valueAxis.title.text = String(reg.valueAxisTitle); chart.axes.valueAxis.title.visible = true; } catch (_) {}
            }
            if (reg.categoryAxisTitle) {
                try { chart.axes.categoryAxis.title.text = String(reg.categoryAxisTitle); chart.axes.categoryAxis.title.visible = true; } catch (_) {}
            }
        }

        // Series names / custom sources — NEVER DELETE SERIES that Excel created from srcRange!
        const seriesList = (Array.isArray(def.series) && def.series.length > 0)
            ? def.series
            : (Array.isArray(reg.seriesNames) ? reg.seriesNames.map(n => ({ name: n })) : []);

        if (seriesList.length > 0) {
            try {
                chart.series.load("items/name");
                await context.sync();
                for (let i = 0; i < seriesList.length && i < chart.series.items.length; i++) {
                    const sd = seriesList[i];
                    if (sd && sd.name) {
                        try { chart.series.items[i].name = String(sd.name); } catch (_) {}
                    }
                    if (sd && sd.valuesSource) {
                        try {
                            const r = this.resolveChartRange(context, sd.valuesSource, ws);
                            if (r) chart.series.items[i].setValues(r);
                        } catch (_) {}
                    }
                    if (sd && sd.catSource) {
                        try {
                            const r = this.resolveChartRange(context, sd.catSource, ws);
                            if (r) chart.series.items[i].setXAxisValues(r);
                        } catch (_) {}
                    }
                }
                await context.sync();
            } catch (_) {}
        }

        if (def.name) {
            try { chart.name = def.name; await context.sync(); } catch (_) {}
        }
        try { chart.load("name"); await context.sync(); return chart.name; }
        catch (_) { return def.name; }
    },

    /** UI entry point: restore the most recently backed-up sheet. */
    async restoreLastBackup() {
        if (!this.isOfficeInitialized) {
            console.log("[Dev Mode] restoreLastBackup called");
            return { success: true, message: "[Dev Mode] Backup restored." };
        }
        const target = (this.lastBackup && this.lastBackup.sheet) || null;
        if (!target) return {
            success: false,
            message: "No backup available yet. Backups are created automatically before risky edits."
        };
        // Bug 11: Pop to temp; repush on failure.
        try {
            let msg = "";
            await Excel.run(async (context) => {
                const rec = (this._backupRecords || {})[target] || null;
                msg = await this.rollbackSheetToBackup(context, target, rec);
            });
            this.lastBackup = null;
            this.persistUndoMeta();
            return { success: true, message: msg };
        } catch (err) {
            return { success: false, message: `Restore failed: ${err.message}` };
        }
    },

    sanitizeFileName(name) {
        let safe = String(name || "workbook-copy").replace(/[\\/:*?"<>|]/g, "").trim().substring(0, 80) || "workbook-copy";
        if (!safe.toLowerCase().match(/\.(xlsx|xlsm|xlsb|xls)$/)) safe += ".xlsx";
        return safe;
    },

    bytesToBase64(bytes) {
        let bin = "";
        const CHUNK = 32768;
        for (let i = 0; i < bytes.length; i += CHUNK) {
            bin += String.fromCharCode.apply(null, bytes.slice(i, i + CHUNK));
        }
        return btoa(bin);
    },

    readWorkbookFile() {
        return new Promise((resolve, reject) => {
            try {
                if (typeof Office === "undefined" || !Office.context || !Office.context.document) {
                    reject(new Error("Office is not available."));
                    return;
                }
                const okStatus = (Office.AsyncResultStatus && Office.AsyncResultStatus.Succeeded) || "succeeded";
                let fileType = "compressed";
                try { fileType = Office.FileType.Compressed; } catch (_) {}
                Office.context.document.getFileAsync(fileType, { sliceSize: 1048576 }, (result) => {
                    if (!result || result.status !== okStatus) {
                        const msg = (result && result.error && result.error.message) ||
                            "Could not read the workbook file. Press Ctrl+S to save, then retry.";
                        reject(new Error(msg));
                        return;
                    }
                    const file = result.value;
                    const parts = [];
                    const finish = (err, data) => {
                        try { file.closeAsync(); } catch (_) {}
                        if (err) reject(err); else resolve(data);
                    };
                    const next = (i) => {
                        if (i >= file.sliceCount) { finish(null, [].concat(...parts)); return; }
                        file.getSliceAsync(i, (sr) => {
                            if (!sr || sr.status !== okStatus) {
                                finish(new Error((sr && sr.error && sr.error.message) || "Slice read failed."), null);
                                return;
                            }
                            parts.push(Array.from(sr.value.data));
                            next(i + 1);
                        });
                    };
                    next(0);
                });
            } catch (e) {
                reject(e);
            }
        });
    },

    async saveCopyAs(fileName) {
        if (!this.isOfficeInitialized) {
            return { success: true, message: `[Dev Mode] Would save copy as "${this.sanitizeFileName(fileName)}".` };
        }
        const safe = this.sanitizeFileName(fileName);
        const bytes = await this.readWorkbookFile();
        const resp = await fetch("http://localhost:3000/api/save-copy", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ filename: safe, data: this.bytesToBase64(bytes) })
        });
        if (!resp.ok) throw new Error(`Bridge refused the file (HTTP ${resp.status}).`);
        const j = await resp.json();
        if (!j.ok) throw new Error(j.error || "Bridge failed to save the copy.");
        return {
            success: true,
            message: `Saved a full copy as "${j.filename}" → ${j.path}. Open it from that path; the original file is untouched.`,
            path: j.path
        };
    },

    /**
     * Pre-batch pass: FULL backup of EVERY sheet the batch touches.
     */
    async prepareBackups(actionsList) {
        this._backupMap = {};
        this._backupNotes = [];
        this._lastBatchBackups = [];
        const SKIP = {
            save_copy_as: 1, restore_backup: 1, delete_backup: 1,
            add_sheet: 1
        };
        const need = {};
        let activeName = null;
        try {
            await Excel.run(async (ctx) => {
                const a = ctx.workbook.worksheets.getActiveWorksheet();
                a.load("name");
                await ctx.sync();
                activeName = a.name;
            });
        } catch (_) {}
        for (const act of actionsList || []) {
            if (!act || SKIP[act.type]) continue;
            let s = null;
            if (act.type === "delete_sheet") s = act.name;
            else if (act.type === "rename_sheet") s = act.from || activeName;
            else s = act.sheet || activeName;
            // Bug 9: Strengthen guard — never back up backup sheets.
            if (!s || this.isBackupSheet(s)) continue;
            need[s] = true;
        }
        const names = Object.keys(need);
        if (names.length === 0) return;
        try {
            await Excel.run(async (ctx) => {
                for (const s of names) {
                    try {
                        this._backupMap[s] = await this.ensureBackupSheet(ctx, s);
                    } catch (e) {
                        this._backupNotes.push(`backup failed for "${s}": ${e.message || e}`);
                    }
                }
            });
        } catch (e) {
            this._backupNotes.push(`backup pass failed: ${e.message || e}`);
        }
    },

    async init() {
        this.restoreUndoMeta();
        if (typeof Office !== "undefined") {
            try {
                const info = await Office.onReady();
                if (info.host === Office.HostType.Excel) {
                    this.isOfficeInitialized = true;
                    console.log("Office.js initialized successfully in Excel host.");
                    return true;
                }
            } catch (err) {
                console.warn("Office.onReady error:", err);
            }
        }
        console.warn("Running outside of Excel. Mock/Dev mode active.");
        return false;
    },

    /**
     * Bug 21: Debounced selection handler — fires at most once per 500ms.
     * Bug 22: Sequence guard — discards stale reads that arrive out of order.
     * Bug 20: Cancellation counter — aborts superseded getFullSheetData calls.
     */
    async onSelectionChange(callback) {
        if (!this.isOfficeInitialized) return;
        try {
            Office.context.document.addHandlerAsync(
                Office.EventType.DocumentSelectionChanged,
                () => {
                    // Bug 21: Debounce — wait 500ms of silence before reading.
                    if (this._selDebounceTimer) clearTimeout(this._selDebounceTimer);
                    this._selDebounceTimer = setTimeout(async () => {
                        // Bug 22: Tag this read with a sequence number.
                        const mySeq = ++this._selReadSeq;
                        try {
                            // Bug 20: Pass generation so abandoned reads bail early.
                            const data = await this.getActiveSheetData();
                            // Discard if a newer read already finished.
                            if (mySeq !== this._selReadSeq) return;
                            callback(data);
                        } catch (err) {
                            if (mySeq !== this._selReadSeq) return;
                            callback({ readError: String((err && err.message) || err) });
                        }
                    }, 500);
                }
            );
        } catch (err) {
            console.warn("Failed to attach onSelectionChanged listener:", err);
        }
    },

    /**
     * Lightweight read of just the ACTIVE sheet + current selection.
     * Used by the selection-change debounce handler to avoid reading all sheets
     * on every click. Bug 21: Much cheaper than getFullSheetData().
     */
    async getActiveSheetData() {
        if (!this.isOfficeInitialized) {
            return this._devMockData();
        }
        const MAX_ROWS = 500, MAX_COLS = 100;
        const SEL_MAX_ROWS = 200, SEL_MAX_COLS = 50;

        let selInfo = null;
        try {
            await Excel.run(async (ctx) => {
                const sel = ctx.workbook.getSelectedRange();
                sel.load(["address", "rowCount", "columnCount"]);
                await ctx.sync();
                const rows = Math.max(1, Math.min(sel.rowCount, SEL_MAX_ROWS));
                const cols = Math.max(1, Math.min(sel.columnCount, SEL_MAX_COLS));
                let vals = [];
                try {
                    if (sel.rowCount <= SEL_MAX_ROWS && sel.columnCount <= SEL_MAX_COLS) {
                        sel.load("values");
                        await ctx.sync();
                        vals = sel.values;
                    } else {
                        const win = sel.getCell(0, 0).getResizedRange(rows - 1, cols - 1);
                        win.load("values");
                        await ctx.sync();
                        vals = win.values;
                    }
                } catch (_) { vals = []; }
                selInfo = {
                    address: sel.address,
                    rowCount: sel.rowCount,
                    columnCount: sel.columnCount,
                    values: vals,
                    truncated: sel.rowCount > SEL_MAX_ROWS || sel.columnCount > SEL_MAX_COLS
                };
            });
        } catch (_) { selInfo = null; }

        return await Excel.run(async (context) => {
            const activeSheet = context.workbook.worksheets.getActiveWorksheet();
            activeSheet.load("name");
            await context.sync();
            const sheetName = activeSheet.name;
            let address = null, rowCount = 0, columnCount = 0;
            let values = [], formulas = [], truncated = false, totalRows = 0, totalCols = 0;
            let tables = [], chartCount = 0, charts = [];
            try {
                const usedRange = activeSheet.getUsedRange();
                usedRange.load(["address", "rowCount", "columnCount", "rowIndex", "columnIndex"]);
                await context.sync();
                totalRows = usedRange.rowCount;
                totalCols = usedRange.columnCount;
                truncated = usedRange.rowCount > MAX_ROWS || usedRange.columnCount > MAX_COLS;
                const rows = Math.min(usedRange.rowCount, MAX_ROWS);
                const cols = Math.min(usedRange.columnCount, MAX_COLS);
                address = usedRange.address;
                rowCount = rows; columnCount = cols;
                if (rows > 0 && cols > 0) {
                    const view = activeSheet.getRangeByIndexes(
                        usedRange.rowIndex, usedRange.columnIndex, rows, cols);
                    view.load(["values", "formulas"]);
                    await context.sync();
                    values = view.values; formulas = view.formulas;
                }
            } catch (_) {}
            try {
                const tbl = activeSheet.tables;
                tbl.load("items/name");
                await context.sync();
                for (const t of tbl.items) {
                    try { const tr = t.getRange(); tr.load("address"); await context.sync();
                        tables.push({ name: t.name, address: tr.address }); } catch (_) {}
                }
            } catch (_) {}
            try {
                const ch = activeSheet.charts;
                ch.load("items/name");
                const cnt = ch.getCount();
                await context.sync();
                chartCount = cnt.value || 0;
                charts = ch.items.map(c => c.name);
            } catch (_) {}
            return {
                activeSheet: sheetName, sheetName,
                address, rowCount, columnCount, values, formulas,
                truncated, totalRows, totalCols, tables,
                visible: true, chartCount, charts,
                readWarnings: null, selection: selInfo, namedRanges: [], sheets: []
            };
        });
    },

    async getSelectedRangeData() {
        return this.getFullSheetData();
    },

    _devMockData() {
        return {
            activeSheet: "Sheet1 (Dev Mock)",
            address: "A1:C5",
            sheetName: "Sheet1 (Dev Mock)",
            rowCount: 5,
            columnCount: 3,
            values: [["Item","Quantity","Price"],["Apples",10,2.5],["Bananas",25,1.2],["Oranges",15,3.0],["Grapes",8,4.5]],
            formulas: [["Item","Quantity","Price"],["Apples",10,2.5],["Bananas",25,1.2],["Oranges",15,3.0],["Grapes",8,4.5]],
            truncated: false, totalRows: 5, totalCols: 3,
            tables: [{ name: "MockTable", address: "Sheet1 (Dev Mock)!A1:C5" }],
            visible: true, chartCount: 0, charts: [], readWarnings: null,
            selection: { address: "A1:C5", values: [["Item","Quantity","Price"],["Apples",10,2.5]] },
            namedRanges: [], sheets: []
        };
    },

    async getFullSheetData() {
        if (!this.isOfficeInitialized) {
            return this._devMockData();
        }

        // Bug 20: Increment generation; getActiveSheetData calls check this.
        const myGen = ++this._readGeneration;

        const MAX_ROWS = 500;
        const MAX_COLS = 100;
        const SEL_MAX_ROWS = 200;
        const SEL_MAX_COLS = 50;

        let selInfo = null;
        try {
            await Excel.run(async (ctx) => {
                const sel = ctx.workbook.getSelectedRange();
                sel.load(["address", "rowCount", "columnCount"]);
                await ctx.sync();
                if (myGen !== this._readGeneration) return; // Bug 20: cancelled
                const rows = Math.max(1, Math.min(sel.rowCount, SEL_MAX_ROWS));
                const cols = Math.max(1, Math.min(sel.columnCount, SEL_MAX_COLS));
                let vals = [];
                try {
                    if (sel.rowCount <= SEL_MAX_ROWS && sel.columnCount <= SEL_MAX_COLS) {
                        sel.load("values");
                        await ctx.sync();
                        vals = sel.values;
                    } else {
                        const win = sel.getCell(0, 0).getResizedRange(rows - 1, cols - 1);
                        win.load("values");
                        await ctx.sync();
                        vals = win.values;
                    }
                } catch (_) { vals = []; }
                selInfo = {
                    address: sel.address,
                    rowCount: sel.rowCount,
                    columnCount: sel.columnCount,
                    values: vals,
                    truncated: sel.rowCount > SEL_MAX_ROWS || sel.columnCount > SEL_MAX_COLS
                };
            });
        } catch (_) { selInfo = null; }

        if (myGen !== this._readGeneration) return null; // Bug 20: cancelled

        let namedRanges = [];
        try {
            await Excel.run(async (ctx) => {
                const names = ctx.workbook.names;
                names.load("items/name, items/value");
                await ctx.sync();
                for (const ni of names.items) {
                    try { namedRanges.push({ name: ni.name, refersTo: ni.value }); }
                    catch (_) {}
                }
            });
        } catch (_) { namedRanges = []; }

        if (myGen !== this._readGeneration) return null; // Bug 20: cancelled

        return await Excel.run(async (context) => {
            const worksheets = context.workbook.worksheets;
            worksheets.load("items/name, items/visibility");
            const activeSheet = context.workbook.worksheets.getActiveWorksheet();
            activeSheet.load("name");
            await context.sync();

            if (myGen !== this._readGeneration) return null; // Bug 20: cancelled

            const activeSheetName = activeSheet.name;
            const sheetsData = [];
            const selection = selInfo;

            for (const ws of worksheets.items) {
                if (String(ws.name).indexOf("AI_Backup__") === 0) {
                    sheetsData.push({
                        name: ws.name, backup: true, visible: false,
                        address: null, rowCount: 0, columnCount: 0,
                        values: [], formulas: [], truncated: false,
                        totalRows: 0, totalCols: 0, tables: [], commentCount: 0
                    });
                    continue;
                }
                let visibility = "Visible";
                try { visibility = String(ws.visibility); } catch (_) {}
                const entry = {
                    name: ws.name,
                    visible: visibility === "Visible",
                    address: null, rowCount: 0, columnCount: 0,
                    values: [], formulas: [],
                    truncated: false, totalRows: 0, totalCols: 0,
                    tables: [], commentCount: 0, chartCount: 0, charts: []
                };
                try {
                    const usedRange = ws.getUsedRange();
                    usedRange.load(["address", "rowCount", "columnCount", "rowIndex", "columnIndex"]);
                    await context.sync();
                    entry.totalRows = usedRange.rowCount;
                    entry.totalCols = usedRange.columnCount;
                    entry.truncated = usedRange.rowCount > MAX_ROWS || usedRange.columnCount > MAX_COLS;
                    const rows = Math.min(usedRange.rowCount, MAX_ROWS);
                    const cols = Math.min(usedRange.columnCount, MAX_COLS);
                    entry.address = usedRange.address;
                    entry.rowCount = rows;
                    entry.columnCount = cols;
                    if (rows > 0 && cols > 0) {
                        const view = ws.getRangeByIndexes(
                            usedRange.rowIndex, usedRange.columnIndex, rows, cols);
                        view.load(["values", "formulas"]);
                        await context.sync();
                        entry.values = view.values;
                        entry.formulas = view.formulas;
                    }
                } catch (e) {
                    const msg = String((e && e.message) || e);
                    if (!/doesn'?t exist|not found/i.test(msg)) {
                        entry.error = msg;
                    }
                }
                try {
                    const tables = ws.tables;
                    tables.load("items/name");
                    await context.sync();
                    for (const t of tables.items) {
                        try {
                            const tr = t.getRange();
                            tr.load("address");
                            await context.sync();
                            entry.tables.push({ name: t.name, address: tr.address });
                        } catch (_) {
                            entry.tables.push({ name: t.name, address: null });
                        }
                    }
                } catch (_) {}
                try {
                    const countResult = ws.comments.getCount();
                    await context.sync();
                    entry.commentCount = countResult.value || 0;
                } catch (_) {}
                try {
                    const chartColl = ws.charts;
                    chartColl.load("items/name");
                    const chartCountResult = chartColl.getCount();
                    await context.sync();
                    entry.chartCount = chartCountResult.value || 0;
                    entry.charts = chartColl.items.map(ch => ch.name);
                } catch (_) {}
                sheetsData.push(entry);
            }

            const active = sheetsData.find(s => s.name === activeSheetName) || sheetsData[0] || {};
            const unreadable = sheetsData
                .filter(s => s.error && !s.backup)
                .map(s => `${s.name} (${s.error})`);
            return {
                activeSheet: activeSheetName,
                sheetName: activeSheetName,
                address: active.address,
                rowCount: active.rowCount,
                columnCount: active.columnCount,
                values: active.values,
                formulas: active.formulas,
                truncated: !!active.truncated,
                totalRows: active.totalRows,
                totalCols: active.totalCols,
                tables: active.tables || [],
                visible: active.visible !== false,
                chartCount: active.chartCount || 0,
                charts: active.charts || [],
                readWarnings: unreadable.length > 0
                    ? `Could not read ${unreadable.length} sheet(s): ${unreadable.slice(0, 3).join("; ")}`
                    : null,
                selection: selection,
                namedRanges: namedRanges,
                sheets: sheetsData
            };
        });
    },

    async addRangeValues(sourceRange, targetCell) {
        const actions = [{
            type: "set_formulas",
            range: targetCell,
            formulas: [[`=SUM(${sourceRange})`]]
        }];
        return this.executeActions(actions);
    },

    /**
     * Adds a Total row with SUM formulas at the bottom of each numeric column.
     * Bug 13: Uses usedRange.rowIndex + usedRange.columnIndex for correct
     * addresses when the data doesn't start at A1.
     */
    async addTotalRow(labelCell) {
        if (!this.isOfficeInitialized) {
            console.log("[Dev Mode] addTotalRow called");
            return { success: true, message: "[Dev Mode] Total row added." };
        }
        try {
            return await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                sheet.load("name");
                const usedRange = sheet.getUsedRange();
                usedRange.load(["address", "values", "rowCount", "columnCount", "rowIndex", "columnIndex"]);
                await context.sync();

                const rowCount = usedRange.rowCount;
                const colCount = usedRange.columnCount;
                const values = usedRange.values;
                // Bug 13: Account for non-A1 starting position.
                const startRow = usedRange.rowIndex;     // 0-based row of top-left
                const startCol = usedRange.columnIndex;  // 0-based col of top-left
                const totalRowIdx = startRow + rowCount + 1; // 1-based Excel row of total

                const actions = [];
                if (labelCell) {
                    actions.push({ type: "set_values", range: labelCell, values: [["Total"]] });
                }
                for (let c = 0; c < colCount; c++) {
                    let isNumeric = false;
                    for (let r = 0; r < rowCount; r++) {
                        const val = values[r][c];
                        if (typeof val === "number" && !isNaN(val)) { isNumeric = true; break; }
                    }
                    if (isNumeric) {
                        const colLetter = this.columnLetter(startCol + c);
                        const firstDataRow = startRow + 1;  // 1-based
                        const lastDataRow  = startRow + rowCount; // 1-based
                        const rangeAddr = `${colLetter}${firstDataRow}:${colLetter}${lastDataRow}`;
                        const targetCell = `${colLetter}${totalRowIdx}`;
                        actions.push({
                            type: "set_formulas",
                            range: targetCell,
                            formulas: [[`=SUM(${rangeAddr})`]]
                        });
                    }
                }
                if (actions.length === 0) {
                    return { success: true, message: "No numeric columns found to sum." };
                }
                const bridge = this;
                for (const act of actions) {
                    const rangeAddr = act.range;
                    if (!rangeAddr) continue;
                    const targetRange = sheet.getRange(rangeAddr);
                    targetRange.load(["address", "values", "formulas", "numberFormat"]);
                    await context.sync();
                    bridge.pushSnapshot(sheet.name, targetRange.address,
                        targetRange.values, targetRange.formulas, targetRange.numberFormat);
                    if (act.type === "set_values" && Array.isArray(act.values)) {
                        targetRange.values = act.values;
                    }
                    if (act.type === "set_formulas" && Array.isArray(act.formulas)) {
                        targetRange.formulas = act.formulas;
                    }
                }
                await context.sync();
                return { success: true, message: `Total row added at row ${totalRowIdx}.` };
            });
        } catch (err) {
            return { success: false, message: `Excel error: ${err.message}` };
        }
    },

    // ── Action validation layer (Bug 16) ────────────────────────────────────

    /** Known action types that require a range field. */
    _RANGE_ACTIONS: new Set([
        "set_values", "set_formulas", "format_range", "set_number_format",
        "borders", "clear_range", "fill_formula", "conditional_format",
        "clear_conditional_format", "sort_range", "find_replace",
        "add_comment", "create_table", "autofit", "merge_range"
    ]),

    /** Known action types that require a sheet name field. */
    _SHEET_NAME_ACTIONS: new Set([
        "delete_sheet", "hide_sheet", "unhide_sheet"
    ]),

    _A1_RE: /^[A-Za-z]{1,3}[1-9]\d*(:[A-Za-z]{1,3}[1-9]\d*)?$/,

    /**
     * Validate a single action object. Returns null on success or an error string.
     */
    validateAction(act) {
        if (!act || typeof act !== "object") return "action is not an object";
        const type = act.type;
        if (!type) return "missing type field";

        // Range required?
        if (this._RANGE_ACTIONS.has(type)) {
            const r = act.range || act.cell;
            if (!r) return `${type} requires a range field`;
            const localR = this.localAddress(String(r));
            if (!this._A1_RE.test(localR)) return `${type} has invalid range "${r}"`;
        }

        // Sheet name required?
        if (this._SHEET_NAME_ACTIONS.has(type)) {
            if (!act.name) return `${type} requires a name field`;
        }

        // 2D array validation for set_values / set_formulas
        if (type === "set_values" && act.values !== undefined) {
            if (!Array.isArray(act.values) || !Array.isArray(act.values[0]))
                return "set_values: values must be a 2D array";
        }
        if (type === "set_formulas" && act.formulas !== undefined) {
            if (!Array.isArray(act.formulas) || !Array.isArray(act.formulas[0]))
                return "set_formulas: formulas must be a 2D array";
        }

        return null; // valid
    },

    /**
     * Bug 14: Adjust row references in a formula string by rowDelta.
     * E.g. adjustFormula("=A2*B2", 1) => "=A3*B3"
     * Simple regex-based substitution for absolute A1 row references.
     */
    adjustFormula(formula, rowDelta) {
        if (!rowDelta) return formula;
        return String(formula).replace(/([A-Za-z]+)(\d+)/g, (m, col, row) => {
            return col + (parseInt(row, 10) + rowDelta);
        });
    },

    /**
     * Build a summary of actions for the preview panel (Bug 26).
     * Returns an HTML string.
     */
    buildActionPreviewHtml(actions) {
        if (!actions || actions.length === 0) return "";
        const lines = actions.map((act, i) => {
            const loc = act.sheet ? `<em>${act.sheet}</em>!` : "";
            const range = act.range || act.cell || act.name || "";
            return `<li>${i + 1}. <strong>${act.type}</strong> ${loc}${range}</li>`;
        });
        return `<details class="action-preview"><summary>📋 ${actions.length} action(s) prepared — click to preview</summary><ul>${lines.join("")}</ul></details>`;
    },

    /**
     * Executes a batch of actions proposed by the AI model.
     * Bug 4: Freezes activeSheet.name before any await.
     * Bug 15: Transactional — rolls back backed-up sheets on partial failure.
     */
    async executeActions(actionsList) {
        if (!actionsList || actionsList.length === 0) return { success: true, message: "No actions to perform." };

        if (!this.isOfficeInitialized) {
            console.log("[Dev Mode] Applied actions:", actionsList);
            return { success: true, message: `[Dev Mode] Applied ${actionsList.length} action(s) successfully.` };
        }

        // Bug 16: Validate all actions before touching Excel.
        const validationErrors = [];
        for (const act of actionsList) {
            const err = this.validateAction(act);
            if (err) validationErrors.push(`${act.type || "unknown"}: ${err}`);
        }
        if (validationErrors.length > 0) {
            return {
                success: false,
                message: `Action validation failed: ${validationErrors.slice(0, 3).join("; ")}`,
                details: validationErrors.map(e => ({ type: "validation", ok: false, error: e }))
            };
        }

        // New actions invalidate the redo history.
        this.redoStack = [];
        this._createdChartsThisBatch = [];

        try { await this.prepareBackups(actionsList); } catch (_) {}

        const outcomes = [];
        let batchFailed = false;

        try {
            await Excel.run(async (context) => {
                // Bug 4: Capture active sheet name ONCE before any await.
                const activeSheet = context.workbook.worksheets.getActiveWorksheet();
                activeSheet.load("name");
                await context.sync();
                const frozenActiveSheetName = activeSheet.name;

                // Helper: resolve target sheet using frozen name.
                const resolveSheet = (act) => {
                    if (act.sheet) return context.workbook.worksheets.getItem(act.sheet);
                    return context.workbook.worksheets.getItem(frozenActiveSheetName);
                };

                const bridge = this;
                for (const act of actionsList) {
                  try {

                    // ── SHEET-LEVEL ACTIONS ──────────────────────────────────

                    // Bug 36: add_sheet — use getItemOrNullObject; only create+snapshot if new.
                    if (act.type === "add_sheet") {
                        const name = (act.name || "NewSheet").replace(/[\[\]\*\/\\\?:]/g, "").substring(0, 31);
                        const existing = context.workbook.worksheets.getItemOrNullObject(name);
                        await context.sync();
                        let newSheet;
                        let created = false;
                        if (existing.isNullObject) {
                            newSheet = context.workbook.worksheets.add(name);
                            created = true;
                        } else {
                            newSheet = existing;
                        }
                        if (act.activate !== false) newSheet.activate();
                        await context.sync();
                        if (created) {
                            bridge.pushSnapshot(name, null, null, null, null, "add", {});
                        }
                        outcomes.push({ type: act.type, ok: true, detail: "sheet=" + name + (created ? " (created)" : " (already existed)") });
                        continue;
                    }

                    if (act.type === "rename_sheet") {
                        const target = act.from
                            ? context.workbook.worksheets.getItem(act.from)
                            : context.workbook.worksheets.getItem(frozenActiveSheetName);
                        const fromName = act.from || frozenActiveSheetName;
                        const newName = (act.to || "Sheet").replace(/[\[\]\*\/\\\?:]/g, "").substring(0, 31);
                        target.name = newName;
                        await context.sync();
                        bridge.pushSnapshot(newName, null, null, null, null, "rename", { from: fromName });
                        outcomes.push({ type: act.type, ok: true, detail: fromName + " -> " + newName });
                        continue;
                    }

                    if (act.type === "delete_sheet" && act.name) {
                        context.workbook.worksheets.getItem(act.name).delete();
                        await context.sync();
                        // Bug 10: Store explicit backupName in the undo entry.
                        bridge.pushSnapshot(act.name, null, null, null, null, "structural",
                            { backup: (bridge._backupMap || {})[act.name] || bridge.backupNameFor(act.name) });
                        outcomes.push({ type: act.type, ok: true, detail: "sheet=" + act.name });
                        continue;
                    }

                    if ((act.type === "hide_sheet" || act.type === "unhide_sheet") && act.name) {
                        const ws = context.workbook.worksheets.getItem(act.name);
                        let prev = "Visible";
                        try {
                            ws.load("visibility");
                            await context.sync();
                            prev = String(ws.visibility);
                        } catch (_) {}
                        ws.visibility = (act.type === "hide_sheet") ? "Hidden" : "Visible";
                        await context.sync();
                        bridge.pushSnapshot(act.name, null, null, null, null, "visibility", { from: prev });
                        outcomes.push({ type: act.type, ok: true, detail: "sheet=" + act.name });
                        continue;
                    }

                    if (act.type === "insert_rows" || act.type === "delete_rows") {
                        const target = act.sheet
                            ? context.workbook.worksheets.getItem(act.sheet)
                            : context.workbook.worksheets.getItem(frozenActiveSheetName);
                        const sheetName = act.sheet || frozenActiveSheetName;
                        const count = Math.max(1, act.count || 1);
                        const rowIdx = Math.max(0, (act.row || 1) - 1);
                        const anchor = target.getRangeByIndexes(rowIdx, 0, count, 1).getEntireRow();
                        if (act.type === "insert_rows") anchor.insert(Excel.InsertShiftDirection.down);
                        else anchor.delete(Excel.DeleteShiftDirection.up);
                        await context.sync();
                        bridge.pushSnapshot(sheetName, null, null, null, null, "structural",
                            { backup: (bridge._backupMap || {})[sheetName] || bridge.backupNameFor(sheetName) });
                        outcomes.push({ type: act.type, ok: true, detail: "row=" + (act.row || 1) + " count=" + count });
                        continue;
                    }

                    if (act.type === "insert_columns" || act.type === "delete_columns") {
                        const target = act.sheet
                            ? context.workbook.worksheets.getItem(act.sheet)
                            : context.workbook.worksheets.getItem(frozenActiveSheetName);
                        const sheetName = act.sheet || frozenActiveSheetName;
                        const count = Math.max(1, act.count || 1);
                        const colIdx = Math.max(0, bridge.resolveColumn(act.column !== undefined ? act.column : 0));
                        const anchor = target.getRangeByIndexes(0, colIdx, 1, count).getEntireColumn();
                        if (act.type === "insert_columns") anchor.insert(Excel.InsertShiftDirection.right);
                        else anchor.delete(Excel.DeleteShiftDirection.left);
                        await context.sync();
                        bridge.pushSnapshot(sheetName, null, null, null, null, "structural",
                            { backup: (bridge._backupMap || {})[sheetName] || bridge.backupNameFor(sheetName) });
                        outcomes.push({ type: act.type, ok: true, detail: "column=" + bridge.columnLetter(colIdx) + " count=" + count });
                        continue;
                    }

                    if (act.type === "hide_rows" || act.type === "unhide_rows" ||
                        act.type === "hide_columns" || act.type === "unhide_columns") {
                        const target = act.sheet
                            ? context.workbook.worksheets.getItem(act.sheet)
                            : context.workbook.worksheets.getItem(frozenActiveSheetName);
                        const count = Math.max(1, act.count || 1);
                        const isRow = act.type.indexOf("_rows") >= 0;
                        const hide = act.type.indexOf("hide_") === 0;
                        let detail = "";
                        if (isRow) {
                            const rowIdx = Math.max(0, (act.row || 1) - 1);
                            const r = target.getRangeByIndexes(rowIdx, 0, count, 1).getEntireRow();
                            r.rowHidden = hide;
                            detail = `row=${act.row || 1} count=${count} hidden=${hide}`;
                        } else {
                            const colIdx = Math.max(0, bridge.resolveColumn(act.column !== undefined ? act.column : 0));
                            const c = target.getRangeByIndexes(0, colIdx, 1, count).getEntireColumn();
                            c.columnHidden = hide;
                            detail = `column=${bridge.columnLetter(colIdx)} count=${count} hidden=${hide}`;
                        }
                        await context.sync();
                        outcomes.push({ type: act.type, ok: true, detail: detail });
                        continue;
                    }

                    if (act.type === "freeze_panes") {
                        const target = act.sheet
                            ? context.workbook.worksheets.getItem(act.sheet)
                            : context.workbook.worksheets.getItem(frozenActiveSheetName);
                        const sheetName = act.sheet || frozenActiveSheetName;
                        let prevFrozen = null;
                        try {
                            const loc = target.freezePanes.getLocationOrNullObject();
                            loc.load("address");
                            await context.sync();
                            if (!loc.isNullObject) prevFrozen = loc.address;
                        } catch (_) {}
                        let rows = act.rows;
                        let cols = act.columns;
                        if (rows === undefined && cols === undefined && act.range) {
                            const cell = String(act.range).split("!").pop().split(":")[0];
                            const m = /^([A-Za-z]+)(\d+)$/.exec(cell.trim());
                            if (m) {
                                rows = parseInt(m[2], 10) - 1;
                                cols = bridge.columnIndex(m[1]);
                            } else { rows = 0; cols = 0; }
                        }
                        rows = Math.max(0, rows || 0);
                        cols = Math.max(0, cols || 0);
                        target.freezePanes.unfreeze();
                        if (rows > 0) target.freezePanes.freezeRows(rows);
                        if (cols > 0) target.freezePanes.freezeColumns(cols);
                        await context.sync();
                        bridge.pushSnapshot(sheetName, null, null, null, null, "freeze", { from: prevFrozen });
                        outcomes.push({ type: act.type, ok: true, detail: `rows=${rows} cols=${cols}` });
                        continue;
                    }

                    if (act.type === "unfreeze_panes") {
                        const target = act.sheet
                            ? context.workbook.worksheets.getItem(act.sheet)
                            : context.workbook.worksheets.getItem(frozenActiveSheetName);
                        const sheetName = act.sheet || frozenActiveSheetName;
                        let prevFrozen = null;
                        try {
                            const loc = target.freezePanes.getLocationOrNullObject();
                            loc.load("address");
                            await context.sync();
                            if (!loc.isNullObject) prevFrozen = loc.address;
                        } catch (_) {}
                        target.freezePanes.unfreeze();
                        await context.sync();
                        bridge.pushSnapshot(sheetName, null, null, null, null, "freeze", { from: prevFrozen });
                        outcomes.push({ type: act.type, ok: true });
                        continue;
                    }

                    // Bug 29: create_chart — push structural snapshot.
                    if (act.type === "create_chart") {
                        const chartSheet = act.sheet
                            ? context.workbook.worksheets.getItem(act.sheet)
                            : context.workbook.worksheets.getItem(frozenActiveSheetName);
                        const sheetName = act.sheet || frozenActiveSheetName;
                        let srcAddr = String(act.sourceRange || act.range || "");
                        if (!srcAddr) throw new Error("create_chart needs a sourceRange.");
                        let srcRange;
                        const bang = srcAddr.indexOf("!");
                        if (bang >= 0) {
                            const sn = srcAddr.slice(0, bang).replace(/^'/, "").replace(/'$/, "");
                            srcRange = context.workbook.worksheets.getItem(sn)
                                .getRange(srcAddr.slice(bang + 1));
                        } else {
                            const srcSheet = act.sourceSheet
                                ? context.workbook.worksheets.getItem(act.sourceSheet)
                                : chartSheet;
                            srcRange = srcSheet.getRange(srcAddr);
                        }
                        const typeMap = {
                            bar: "BarClustered", column: "ColumnClustered",
                            stackedbar: "BarStacked", stackedcolumn: "ColumnStacked",
                            line: "Line", linemarkers: "LineMarkers", line_markers: "LineMarkers",
                            stackedline: "LineStacked", pie: "Pie", doughnut: "Doughnut",
                            area: "Area", stackedarea: "AreaStacked"
                        };
                        const chartType = typeMap[String(act.chartType || "column").toLowerCase()] || "ColumnClustered";
                        const seriesBy = String(act.seriesBy || "columns").toLowerCase() === "rows"
                            ? Excel.ChartSeriesBy.rows : Excel.ChartSeriesBy.columns;
                        const chart = chartSheet.charts.add(chartType, srcRange, seriesBy);
                        const pos = String(act.position || "E2:K20").split(":");
                        chart.setPosition(pos[0] || "E2", pos[1] || null);
                        if (act.title) {
                            chart.title.text = String(act.title);
                            try { chart.title.visible = true; } catch (_) {}
                        }
                        if (act.legend === false) { try { chart.legend.visible = false; } catch (_) {} }
                        if (act.legendPosition) {
                            const lpMap = { top: "Top", bottom: "Bottom", left: "Left", right: "Right", topright: "TopRight" };
                            try { chart.legend.position = lpMap[String(act.legendPosition).toLowerCase()] || "Right"; } catch (_) {}
                        }
                        if (act.dataLabels) { try { chart.dataLabels.showValue = true; } catch (_) {} }
                        if (act.style !== undefined) { try { chart.style = act.style; } catch (_) {} }
                        if (act.valueAxisTitle) {
                            try { chart.axes.valueAxis.title.text = String(act.valueAxisTitle); chart.axes.valueAxis.title.visible = true; } catch (_) {}
                        }
                        if (act.categoryAxisTitle) {
                            try { chart.axes.categoryAxis.title.text = String(act.categoryAxisTitle); chart.axes.categoryAxis.title.visible = true; } catch (_) {}
                        }
                        chart.load("name");
                        await context.sync();
                        if (Array.isArray(act.seriesNames) && act.seriesNames.length > 0) {
                            try {
                                const series = chart.series;
                                series.load("items/name");
                                await context.sync();
                                act.seriesNames.forEach((n, i) => {
                                    if (n && series.items.length > i) {
                                        try { series.items[i].name = String(n); } catch (_) {}
                                    }
                                });
                                await context.sync();
                            } catch (_) {}
                        }
                        // Record chart creation in registry and in batch list
                        bridge._chartRegistry = bridge._chartRegistry || {};
                        const chartDef = {
                            name: chart.name,
                            chartType: chartType,
                            sourceRange: srcAddr,
                            sourceSheet: act.sourceSheet || null,
                            position: act.position || "E2:K20",
                            seriesBy: act.seriesBy || "columns",
                            title: act.title || null,
                            legend: act.legend !== false,
                            legendPosition: act.legendPosition || "Right",
                            dataLabels: Boolean(act.dataLabels),
                            style: act.style,
                            valueAxisTitle: act.valueAxisTitle,
                            categoryAxisTitle: act.categoryAxisTitle,
                            seriesNames: act.seriesNames || []
                        };
                        bridge._chartRegistry[chart.name] = chartDef;
                        bridge._createdChartsThisBatch = bridge._createdChartsThisBatch || [];
                        bridge._createdChartsThisBatch.push(chart.name);
                        bridge.persistUndoMeta();

                        bridge.pushSnapshot(sheetName, null, null, null, null, "chart_create", {
                            chartName: chart.name,
                            sheet: sheetName,
                            chartDef: chartDef
                        });
                        outcomes.push({ type: act.type, ok: true, detail: "chart=" + chart.name });
                        continue;
                    }

                    if (act.type === "delete_chart") {
                        const ws = act.sheet
                            ? context.workbook.worksheets.getItem(act.sheet)
                            : context.workbook.worksheets.getItem(frozenActiveSheetName);
                        const wsName = act.sheet || frozenActiveSheetName;
                        const charts = ws.charts;
                        charts.load("items/name");
                        await context.sync();
                        const available = charts.items.map(ch => ch.name);
                        if (act.all) {
                            const deletedDefs = [];
                            for (const ch of charts.items) {
                                try {
                                    const d = await bridge.captureChartDef(context, ws, ch);
                                    deletedDefs.push(d);
                                } catch (_) {}
                                ch.delete();
                            }
                            await context.sync();
                            bridge.pushSnapshot(wsName, null, null, null, null, "chart_delete_all", {
                                sheet: wsName,
                                chartDefs: deletedDefs
                            });
                            outcomes.push({ type: act.type, ok: true, detail: `removed all (${available.length}) charts on "${wsName}"` });
                            continue;
                        }
                        let hit = null;
                        if (act.name) {
                            hit = charts.items.find(
                                ch => ch.name === act.name ||
                                      ch.name.toLowerCase() === String(act.name).toLowerCase());
                        } else if (act.index !== undefined && charts.items.length > act.index) {
                            hit = charts.items[act.index];
                        } else {
                            throw new Error(`delete_chart needs a chart name, index, or all:true. Available on "${wsName}": ${available.join(", ") || "none"}.`);
                        }
                        if (!hit) throw new Error(`chart "${act.name || act.index}" not found on "${wsName}". Available: ${available.join(", ") || "none"}.`);
                        let def = null;
                        try {
                            def = await bridge.captureChartDef(context, ws, hit);
                        } catch (_) {}
                        const chartName = hit.name;
                        hit.delete();
                        await context.sync();
                        bridge.pushSnapshot(wsName, null, null, null, null, "chart_delete", {
                            chartName: chartName,
                            sheet: wsName,
                            chartDef: def || { name: chartName, sheet: wsName }
                        });
                        outcomes.push({ type: act.type, ok: true, detail: act.name || ("index " + act.index) });
                        continue;
                    }

                    if (act.type === "save_copy_as") {
                        const r = await bridge.saveCopyAs(act.filename || act.name || "workbook-copy");
                        outcomes.push({ type: act.type, ok: true, detail: r.path || r.message });
                        continue;
                    }

                    if (act.type === "format_chart_axis") {
                        const ws = act.sheet
                            ? context.workbook.worksheets.getItem(act.sheet)
                            : context.workbook.worksheets.getItem(frozenActiveSheetName);
                        const wsName = act.sheet || frozenActiveSheetName;
                        const charts = ws.charts;
                        charts.load("items/name");
                        await context.sync();
                        const available = charts.items.map(ch => ch.name);
                        let chart = null;
                        if (act.chart !== undefined && act.chart !== null) {
                            if (typeof act.chart === "number") {
                                if (act.chart < charts.items.length) chart = charts.items[act.chart];
                            } else {
                                chart = charts.items.find(
                                    ch => ch.name === act.chart ||
                                          ch.name.toLowerCase() === String(act.chart).toLowerCase());
                            }
                        } else if (charts.items.length === 1) {
                            chart = charts.items[0];
                        }
                        if (!chart) throw new Error(`chart "${act.chart !== undefined ? act.chart : "(pick one)"}" not found on "${wsName}". Available: ${available.join(", ") || "none"}.`);
                        const which = String(act.axis || "category").toLowerCase();
                        const ax = which.indexOf("val") === 0 ? chart.axes.valueAxis
                            : which.indexOf("ser") === 0 ? chart.axes.seriesAxis
                            : chart.axes.categoryAxis;
                        if (act.majorUnit !== undefined) ax.majorUnit = act.majorUnit;
                        if (act.minorUnit !== undefined) ax.minorUnit = act.minorUnit;
                        if (act.numberFormat) ax.numberFormat = String(act.numberFormat);
                        if (act.title) { ax.title.text = String(act.title); try { ax.title.visible = true; } catch (_) {} }
                        if (act.majorGridlines !== undefined) { try { ax.majorGridlines.visible = Boolean(act.majorGridlines); } catch (_) {} }
                        if (act.minorGridlines !== undefined) { try { ax.minorGridlines.visible = Boolean(act.minorGridlines); } catch (_) {} }
                        await context.sync();
                        outcomes.push({ type: act.type, ok: true, detail: `axis=${which} on "${chart.name}"` });
                        continue;
                    }

                    if (act.type === "restore_backup") {
                        const s = act.sheet || (bridge.lastBackup && bridge.lastBackup.sheet);
                        if (!s) throw new Error("No sheet specified and no recent backup exists.");
                        const rec0 = (bridge._backupRecords || {})[s] || null;
                        const msg0 = await bridge.rollbackSheetToBackup(context, s, rec0);
                        outcomes.push({ type: act.type, ok: true, detail: msg0 });
                        continue;
                    }

                    if (act.type === "delete_backup") {
                        const s = act.sheet || (bridge.lastBackup && bridge.lastBackup.sheet);
                        if (!s) throw new Error("No sheet specified and no recent backup exists.");
                        const bn = bridge.backupNameFor(s);
                        const b = context.workbook.worksheets.getItemOrNullObject(bn);
                        await context.sync();
                        if (!b.isNullObject) { b.delete(); await context.sync(); }
                        if (bridge.lastBackup && bridge.lastBackup.sheet === s) bridge.lastBackup = null;
                        if (bridge._backupRecords) delete bridge._backupRecords[s];
                        bridge.persistUndoMeta();
                        outcomes.push({ type: act.type, ok: true, detail: `removed "${bn}"` });
                        continue;
                    }

                    // Bug 27: create_pivot_table — only delete conflicting pivot, not all.
                    if (act.type === "create_pivot_table") {
                        const sourceAddr = act.sourceRange || (frozenActiveSheetName + "!" + (act.range || "A1"));
                        const pivotSheetName = (act.destSheet || "PivotTable").replace(/[\[\]\*\/\\\?:]/g, "").substring(0, 31);
                        const destCell = act.destCell || "A1";
                        const allSheets = context.workbook.worksheets;
                        allSheets.load("items/name");
                        await context.sync();
                        let pivotSheet;
                        const existingSheet = allSheets.items.find(s => s.name === pivotSheetName);
                        if (existingSheet) {
                            pivotSheet = existingSheet;
                            // Bug 27: Only delete the NAMED pivot that conflicts, not all pivots.
                            const tableName = ((act.tableName || "PT") + "_").replace(/[^a-zA-Z0-9_]/g, "");
                            pivotSheet.pivotTables.load("items/name");
                            await context.sync();
                            for (const oldPt of pivotSheet.pivotTables.items) {
                                if (String(oldPt.name).indexOf(tableName) === 0) {
                                    oldPt.delete();
                                }
                            }
                            await context.sync();
                            // Bug 28: Warn if destination cell area appears non-empty.
                            try {
                                const destR = pivotSheet.getRange(destCell);
                                destR.load("values");
                                await context.sync();
                                const v = destR.values && destR.values[0] && destR.values[0][0];
                                if (v !== null && v !== undefined && v !== "") {
                                    bridge._backupNotes.push(`Pivot destination "${destCell}" on "${pivotSheetName}" may not be empty — existing data may be overwritten.`);
                                }
                            } catch (_) {}
                        } else {
                            pivotSheet = context.workbook.worksheets.add(pivotSheetName);
                            await context.sync();
                        }
                        const tableName = ((act.tableName || "PT") + "_" + Date.now().toString(36)).replace(/[^a-zA-Z0-9_]/g, "");
                        const pt = pivotSheet.pivotTables.add(tableName, sourceAddr, pivotSheet.getRange(destCell));
                        await context.sync();
                        pt.hierarchies.load("items/name");
                        await context.sync();
                        const findHierarchy = (name) => {
                            const h = pt.hierarchies.items.find(i => i.name.toLowerCase() === name.toLowerCase());
                            if (!h) throw new Error(`Field "${name}" not found in PivotTable. Available: ${pt.hierarchies.items.map(i=>i.name).join(", ")}`);
                            return h;
                        };
                        if (Array.isArray(act.rows)) {
                            for (const f of act.rows) pt.rowHierarchies.add(findHierarchy(f));
                            await context.sync();
                        }
                        if (Array.isArray(act.columns)) {
                            for (const f of act.columns) pt.columnHierarchies.add(findHierarchy(f));
                            await context.sync();
                        }
                        if (Array.isArray(act.values)) {
                            for (const v of act.values) {
                                const fieldName = typeof v === "string" ? v : v.field;
                                const dh = pt.dataHierarchies.add(findHierarchy(fieldName));
                                if (v.aggregation) {
                                    const aggMap = {
                                        "sum": Excel.AggregationFunction.sum,
                                        "count": Excel.AggregationFunction.count,
                                        "average": Excel.AggregationFunction.average,
                                        "max": Excel.AggregationFunction.max,
                                        "min": Excel.AggregationFunction.min
                                    };
                                    dh.summarizeBy = aggMap[v.aggregation.toLowerCase()] || Excel.AggregationFunction.sum;
                                }
                            }
                            await context.sync();
                        }
                        if (Array.isArray(act.filters)) {
                            for (const f of act.filters) pt.filterHierarchies.add(findHierarchy(f));
                            await context.sync();
                        }
                        pivotSheet.activate();
                        await context.sync();
                        // Bug 29: Push structural undo entry.
                        bridge.pushSnapshot(pivotSheetName, null, null, null, null, "structural",
                            { backup: (bridge._backupMap || {})[frozenActiveSheetName] || bridge.backupNameFor(frozenActiveSheetName) });
                        outcomes.push({ type: act.type, ok: true, detail: "pivot=" + tableName + " on " + pivotSheetName });
                        continue;
                    }

                    // ── RANGE-LEVEL ACTIONS ──────────────────────────────────
                    let rangeAddr = act.range || act.cell || null;
                    const targetSheet = resolveSheet(act);
                    const sheetName = act.sheet || frozenActiveSheetName;
                    if (!rangeAddr && act.type === "find_replace") {
                        const ur = targetSheet.getUsedRange();
                        ur.load("address");
                        await context.sync();
                        rangeAddr = bridge.localAddress(ur.address);
                    }
                    if (!rangeAddr) {
                        console.warn("Action skipped — no range and not a sheet-level action:", act.type);
                        outcomes.push({ type: act.type, ok: false, error: "missing range" });
                        continue;
                    }

                    const targetRange = targetSheet.getRange(rangeAddr);
                    targetRange.load(["address", "values", "formulas", "numberFormat",
                                      "rowCount", "columnCount", "columnIndex"]);
                    try {
                        targetRange.format.font.load(["bold", "italic", "color", "size", "name"]);
                        targetRange.format.fill.load("color");
                        targetRange.format.load(["horizontalAlignment", "verticalAlignment", "wrapText"]);
                    } catch (_) {}
                    await context.sync();

                    // Bug 32/33/34/35: Extended snapshot captures borders, merge state,
                    // column widths, row heights.
                    let formatSnap = null;
                    try {
                        formatSnap = {
                            font: {
                                bold: targetRange.format.font.bold,
                                italic: targetRange.format.font.italic,
                                color: targetRange.format.font.color,
                                size: targetRange.format.font.size,
                                name: targetRange.format.font.name
                            },
                            fillColor: targetRange.format.fill.color,
                            horizontalAlignment: targetRange.format.horizontalAlignment,
                            verticalAlignment: targetRange.format.verticalAlignment,
                            wrapText: targetRange.format.wrapText
                        };
                    } catch (_) { formatSnap = null; }

                    // Capture border state for all 6 edges (Bug 33).
                    let borderSnap = null;
                    try {
                        const edgeNames = ["EdgeTop","EdgeBottom","EdgeLeft","EdgeRight",
                                           "InsideHorizontal","InsideVertical"];
                        borderSnap = {};
                        for (const en of edgeNames) {
                            try {
                                const b = targetRange.format.borders.getItem(en);
                                b.load(["style", "weight", "color"]);
                                borderSnap[en] = b; // proxy; values loaded below
                            } catch (_) {}
                        }
                        await context.sync();
                        const resolved = {};
                        for (const en of edgeNames) {
                            if (borderSnap[en]) {
                                resolved[en] = {
                                    style: borderSnap[en].style,
                                    weight: borderSnap[en].weight,
                                    color: borderSnap[en].color
                                };
                            }
                        }
                        borderSnap = resolved;
                    } catch (_) { borderSnap = null; }

                    // Capture merge state (Bug 34).
                    let mergeSnap = null;
                    try {
                        const mc = targetRange.getMergedAreasOrNullObject();
                        mc.load("address");
                        await context.sync();
                        mergeSnap = mc.isNullObject ? null : mc.address;
                    } catch (_) { mergeSnap = null; }

                    // Capture column widths + row heights (Bug 35).
                    let colWidths = null, rowHeights = null;
                    try {
                        targetRange.format.load(["columnWidth", "rowHeight"]);
                        await context.sync();
                        colWidths = targetRange.format.columnWidth;
                        rowHeights = targetRange.format.rowHeight;
                    } catch (_) {}

                    bridge.pushSnapshot(sheetName, targetRange.address,
                        targetRange.values, targetRange.formulas, targetRange.numberFormat,
                        "range", {
                            format: formatSnap,
                            borders: borderSnap,
                            mergeAddress: mergeSnap,
                            colWidths, rowHeights
                        });
                    const rangeOutcomesBefore = outcomes.length;

                    // Action: set_values
                    if (act.type === "set_values" && Array.isArray(act.values)) {
                        targetRange.values = act.values;
                        outcomes.push({ type: act.type, range: targetRange.address, ok: true });
                    }

                    // Action: set_formulas
                    if (act.type === "set_formulas" && Array.isArray(act.formulas)) {
                        targetRange.formulas = act.formulas;
                        outcomes.push({ type: act.type, range: targetRange.address, ok: true });
                    }

                    // Action: format_range
                    if (act.type === "format_range") {
                        if (act.bold !== undefined) targetRange.format.font.bold = Boolean(act.bold);
                        if (act.italic !== undefined) targetRange.format.font.italic = Boolean(act.italic);
                        if (act.fontColor) targetRange.format.font.color = act.fontColor;
                        if (act.fill) targetRange.format.fill.color = act.fill;
                        if (act.fontSize) targetRange.format.font.size = act.fontSize;
                        if (act.wrapText !== undefined) targetRange.format.wrapText = Boolean(act.wrapText);
                        if (act.horizontalAlignment) targetRange.format.horizontalAlignment = act.horizontalAlignment;
                        if (act.verticalAlignment) targetRange.format.verticalAlignment = act.verticalAlignment;
                        outcomes.push({ type: act.type, range: targetRange.address, ok: true });
                    }

                    // Action: set_number_format
                    if (act.type === "set_number_format" && act.format !== undefined) {
                        targetRange.numberFormat = act.format;
                        outcomes.push({ type: act.type, range: targetRange.address, ok: true });
                    }

                    // Action: borders
                    if (act.type === "borders") {
                        const edgeMap = {
                            all: ["EdgeTop","EdgeBottom","EdgeLeft","EdgeRight","InsideHorizontal","InsideVertical"],
                            outer: ["EdgeTop","EdgeBottom","EdgeLeft","EdgeRight"],
                            inner: ["InsideHorizontal","InsideVertical"],
                            top: ["EdgeTop"], bottom: ["EdgeBottom"],
                            left: ["EdgeLeft"], right: ["EdgeRight"]
                        };
                        const edges = edgeMap[String(act.edges || "all").toLowerCase()] || edgeMap.all;
                        const styleMap = { none:"None", continuous:"Continuous", dash:"Dash", dashed:"Dash",
                                           dotted:"Dot", dot:"Dot", dashdot:"DashDot", double:"Double" };
                        const weightMap = { hairline:"Hairline", thin:"Thin", medium:"Medium", thick:"Thick" };
                        for (const e of edges) {
                            const b = targetRange.format.borders.getItem(e);
                            b.style = styleMap[String(act.style || "continuous").toLowerCase()] || "Continuous";
                            b.weight = weightMap[String(act.weight || "thin").toLowerCase()] || "Thin";
                            if (act.color) b.color = act.color;
                        }
                        outcomes.push({ type: act.type, range: targetRange.address, ok: true });
                    }

                    // Action: clear_range
                    if (act.type === "clear_range") {
                        const whatMap = { all:"All", contents:"Contents", formats:"Formats" };
                        targetRange.clear(whatMap[String(act.what || "all").toLowerCase()] || "All");
                        outcomes.push({ type: act.type, range: targetRange.address, ok: true });
                    }

                    // Bug 14: fill_formula fallback adjusts row references.
                    if (act.type === "fill_formula" && act.formula) {
                        const firstCell = targetRange.getCell(0, 0);
                        firstCell.formulas = [[act.formula]];
                        await context.sync();
                        try {
                            firstCell.autoFill(targetRange, Excel.AutoFillType.fillDefault);
                        } catch (_) {
                            // Fallback: build correct 2-D formula grid with adjusted row references.
                            const grid = [];
                            for (let i = 0; i < targetRange.rowCount; i++) {
                                const row = [];
                                for (let j = 0; j < targetRange.columnCount; j++) {
                                    row.push(bridge.adjustFormula(act.formula, i));
                                }
                                grid.push(row);
                            }
                            targetRange.formulas = grid;
                        }
                        outcomes.push({ type: act.type, range: targetRange.address, ok: true });
                    }

                    // Action: conditional_format
                    if (act.type === "conditional_format") {
                        const rule = String(act.rule || "cellValue");
                        if (rule === "clear") {
                            targetRange.conditionalFormats.clearAll();
                        } else if (rule === "cellValue") {
                            const opMap = {
                                between:"Between", notbetween:"NotBetween",
                                equal:"EqualTo", notequal:"NotEqualTo",
                                greater:"GreaterThan", greaterthan:"GreaterThan",
                                less:"LessThan", lessthan:"LessThan",
                                greaterorequal:"GreaterThanOrEqualTo",
                                lessorequal:"LessThanOrEqualTo"
                            };
                            const cf = targetRange.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
                            cf.cellValue.rule = {
                                formula1: String(act.formula1 !== undefined ? act.formula1 : ""),
                                formula2: act.formula2 !== undefined ? String(act.formula2) : undefined,
                                operator: opMap[String(act.operator || "greater").toLowerCase()] || "GreaterThan"
                            };
                            if (act.fill) cf.cellValue.format.fill.color = act.fill;
                            if (act.fontColor) cf.cellValue.format.font.color = act.fontColor;
                            if (act.bold !== undefined) cf.cellValue.format.font.bold = Boolean(act.bold);
                        } else if (rule === "colorScale") {
                            const cf = targetRange.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
                            cf.colorScale.criteria = {
                                minimum: { format: { color: act.minColor || "#63BE7B" }, type: "lowestValue" },
                                midpoint: { format: { color: act.midColor || "#FFEB84" }, type: "percentile", value: "50" },
                                maximum: { format: { color: act.maxColor || "#F8696B" }, type: "highestValue" }
                            };
                        } else if (rule === "dataBar") {
                            const cf = targetRange.conditionalFormats.add(Excel.ConditionalFormatType.dataBar);
                            if (act.fill) cf.dataBar.positiveFormat.fillColor = act.fill;
                            if (act.axisColor) cf.dataBar.axisColor = act.axisColor;
                        }
                        outcomes.push({ type: act.type, range: targetRange.address, ok: true });
                    }
                    if (act.type === "clear_conditional_format") {
                        targetRange.conditionalFormats.clearAll();
                        outcomes.push({ type: act.type, range: targetRange.address, ok: true });
                    }

                    // Action: sort_range
                    if (act.type === "sort_range" && Array.isArray(act.sortBy) && act.sortBy.length) {
                        const fields = act.sortBy.map(s => {
                            const key = (typeof s.column === "number")
                                ? s.column
                                : (bridge.resolveColumn(s.column) - targetRange.columnIndex);
                            return { key: key, ascending: s.ascending !== false };
                        });
                        targetRange.sort.apply(fields, false, act.hasHeaders !== false, Excel.SortOrientation.rows);
                        outcomes.push({ type: act.type, range: targetRange.address, ok: true });
                    }

                    // Bug 12/37: find_replace preserves formulas and numeric types.
                    if (act.type === "find_replace" && act.find !== undefined) {
                        const esc = (s) => String(s).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
                        const findText = String(act.find);
                        const replText = act.replace !== undefined ? String(act.replace) : "";
                        const matchCase = Boolean(act.matchCase);
                        const entire = Boolean(act.matchEntireCell);
                        const vals = targetRange.values;
                        const fmls = targetRange.formulas;
                        let replaced = 0;
                        const rx = new RegExp(esc(findText), matchCase ? "g" : "gi");
                        // We will write back two arrays separately: formulas for formula cells,
                        // values for non-formula cells. We track which cells changed.
                        const newFmls = fmls.map(r => r.slice());
                        const newVals = vals.map(r => r.slice());
                        const changed = []; // [{r,c}]
                        for (let r = 0; r < vals.length; r++) {
                            for (let c = 0; c < vals[r].length; c++) {
                                const formula = fmls[r] && fmls[r][c];
                                // Bug 12: Skip cells that contain formulas (string starting with '=').
                                if (typeof formula === "string" && formula.startsWith("=")) continue;
                                const v = vals[r][c];
                                if (v === null || v === undefined) continue;
                                const s = String(v);
                                let didReplace = false;
                                if (entire) {
                                    const hit = matchCase ? (s === findText)
                                        : (s.toLowerCase() === findText.toLowerCase());
                                    if (hit) {
                                        // Bug 37: Preserve numeric type if possible.
                                        const numReplacement = Number(replText);
                                        newVals[r][c] = (!isNaN(numReplacement) && replText.trim() !== "" &&
                                                          typeof v === "number")
                                            ? numReplacement : replText;
                                        newFmls[r][c] = newVals[r][c];
                                        didReplace = true;
                                    }
                                } else {
                                    rx.lastIndex = 0;
                                    if (rx.test(s)) {
                                        rx.lastIndex = 0;
                                        const replaced_str = s.replace(rx, replText);
                                        newVals[r][c] = replaced_str;
                                        newFmls[r][c] = replaced_str;
                                        didReplace = true;
                                    }
                                }
                                if (didReplace) { replaced++; changed.push({r, c}); }
                            }
                        }
                        if (replaced > 0) {
                            // Write back via formulas array (non-formula cells; plain values are safe here).
                            targetRange.formulas = newFmls;
                        }
                        outcomes.push({ type: act.type, range: targetRange.address, ok: true, detail: replaced + " cell(s)" });
                    }

                    // Action: add_comment
                    if (act.type === "add_comment" && act.text) {
                        const cellAddr = String(act.cell || act.range).split("!").pop();
                        targetSheet.comments.add(cellAddr, String(act.text));
                        outcomes.push({ type: act.type, range: targetRange.address, ok: true });
                    }

                    // Bug 31: create_table — store tableName for undo.
                    if (act.type === "create_table") {
                        const hasHeaders = act.hasHeaders !== false;
                        const table = targetSheet.tables.add(rangeAddr, hasHeaders);
                        const tableName = act.tableName ? act.tableName.replace(/[^a-zA-Z0-9_]/g, "") : null;
                        if (tableName) table.name = tableName;
                        table.style = act.tableStyle || "TableStyleMedium9";
                        // Bug 31: Attach tableCreated to the snapshot so undo can delete it.
                        const lastSnap = bridge.undoStack[bridge.undoStack.length - 1];
                        if (lastSnap) lastSnap.tableCreated = tableName || true;
                        outcomes.push({ type: act.type, range: targetRange.address, ok: true });
                    }

                    // Action: autofit
                    if (act.type === "autofit") {
                        targetRange.format.autofitColumns();
                        if (act.rows) targetRange.format.autofitRows();
                        outcomes.push({ type: act.type, range: targetRange.address, ok: true });
                    }

                    // Action: merge_range
                    if (act.type === "merge_range") {
                        if (act.unmerge) targetRange.unmerge();
                        else targetRange.merge(Boolean(act.across));
                        outcomes.push({ type: act.type, range: targetRange.address, ok: true });
                    }

                    if (outcomes.length === rangeOutcomesBefore) {
                        outcomes.push({
                            type: act.type, range: targetRange.address,
                            ok: false, error: `unknown action type or missing required fields for "${act.type}" - ignored`
                        });
                    }
                  } catch (actionErr) {
                    console.warn("Action failed:", act.type, actionErr);
                    batchFailed = true; // Bug 15: Track partial failure.
                    outcomes.push({
                        type: act.type,
                        range: act.range || act.cell || null,
                        ok: false,
                        error: String((actionErr && actionErr.message) || actionErr)
                    });
                  }
                }

                await context.sync();
            });

        } catch (err) {
            console.error("Error executing Excel actions:", err);
            return { success: false, message: `Excel error: ${err.message}` };
        }

        // Bug 15: If any action failed, roll back all backed-up sheets.
        const failed = outcomes.filter(o => !o.ok);
        if (batchFailed && failed.length > 0 && this._lastBatchBackups && this._lastBatchBackups.length > 0) {
            const rolledBack = [];
            try {
                await Excel.run(async (context) => {
                    for (const sheetName of this._lastBatchBackups) {
                        try {
                            const rec = (this._backupRecords || {})[sheetName] || null;
                            await this.rollbackSheetToBackup(context, sheetName, rec);
                            rolledBack.push(sheetName);
                        } catch (rbErr) {
                            console.warn(`Rollback failed for "${sheetName}":`, rbErr);
                        }
                    }
                });
            } catch (rbErr) {
                console.warn("Batch rollback failed:", rbErr);
            }
            const rbMsg = rolledBack.length > 0
                ? ` Auto-rolled back: ${rolledBack.join(", ")}.`
                : " Rollback attempt failed — use Restore button.";
            return {
                success: false,
                message: `Batch partially failed (${failed.length} error(s)) — changes were reversed.${rbMsg}`,
                details: outcomes
            };
        }

        const appliedCount = outcomes.filter(o => o.ok).length;
        let message = `Applied ${appliedCount} action(s) to spreadsheet.`;
        if (failed.length > 0) {
            message += ` ${failed.length} failed: ` +
                failed.slice(0, 3).map(f =>
                    `${f.type}${f.range ? " @ " + f.range : ""} (${f.error})`).join("; ");
        }
        const backedUp = this._lastBatchBackups || [];
        if (backedUp.length > 0) {
            const recs = this._backupRecords || {};
            message += ` Full backup kept: ` +
                backedUp.map(s => {
                    const r = recs[s] || {};
                    const nt = (r.tables || []).length;
                    const nc = (r.charts || []).length;
                    return `AI_Backup__${s} (values+formats, ${nt} table(s), ${nc} chart(s))`;
                }).join("; ") +
                `. Undo/Restore rolls everything back.`;
        }
        if ((this._backupNotes || []).length > 0) {
            message += ` Backup warnings: ` + this._backupNotes.slice(0, 2).join("; ");
        }
        return { success: failed.length === 0, message: message, details: outcomes };
    },

    /**
     * Undoes the last modification from the undo stack.
     * Bug 11: Pop to temp; re-push on failure so the entry is not lost.
     * Bug 32/33/34/35: Restores borders, merge state, col widths, row heights.
     * Bug 31: Deletes created Table object before restoring values.
     */
    async undoLast() {
        if (this.undoStack.length === 0) {
            return { success: false, message: "Nothing to undo." };
        }

        // Bug 11: Pop to temp; re-push if the operation fails.
        const snapshot = this.undoStack.pop();
        const kind = snapshot.kind || "range";

        if (!this.isOfficeInitialized) {
            console.log("[Dev Mode] Undo snapshot restored:", snapshot);
            const where = snapshot.sheet ? snapshot.sheet + "!" : "";
            return { success: true, message: `[Dev Mode] Reverted ${kind} ${where}${snapshot.range || ""}`.trim() };
        }

        try {
            // Capture the current ("after") state for redo before we revert.
            // Only meaningful for range-kind undos (structural redo is too risky to automate).
            let redoEntry = null;

            const msg = await Excel.run(async (context) => {
                const sheets = context.workbook.worksheets;

                // Capture redo state for range kind BEFORE reverting.
                if (kind === "range" && snapshot.range) {
                    try {
                        const sh = snapshot.sheet
                            ? sheets.getItem(snapshot.sheet)
                            : context.workbook.worksheets.getActiveWorksheet();
                        sh.load("name");
                        const rng = sh.getRange(snapshot.range);
                        rng.load(["values", "formulas", "numberFormat"]);
                        try {
                            rng.format.font.load(["bold", "italic", "color", "size", "name"]);
                            rng.format.fill.load("color");
                            rng.format.load(["horizontalAlignment", "verticalAlignment", "wrapText"]);
                        } catch (_) {}
                        await context.sync();
                        redoEntry = {
                            kind: "range",
                            sheet: sh.name,
                            range: snapshot.range,
                            values: JSON.parse(JSON.stringify(rng.values)),
                            formulas: JSON.parse(JSON.stringify(rng.formulas)),
                            numberFormat: rng.numberFormat ? JSON.parse(JSON.stringify(rng.numberFormat)) : null
                        };
                        try {
                            redoEntry.format = {
                                font: {
                                    bold: rng.format.font.bold,
                                    italic: rng.format.font.italic,
                                    color: rng.format.font.color,
                                    size: rng.format.font.size,
                                    name: rng.format.font.name
                                },
                                fillColor: rng.format.fill.color,
                                horizontalAlignment: rng.format.horizontalAlignment,
                                verticalAlignment: rng.format.verticalAlignment,
                                wrapText: rng.format.wrapText
                            };
                        } catch (_) {}
                    } catch (_) { redoEntry = null; }
                }

                if (kind === "chart_create") {
                    const ws = snapshot.sheet ? sheets.getItem(snapshot.sheet) : sheets.getActiveWorksheet();
                    const ch = ws.charts.getItemOrNullObject(snapshot.chartName);
                    await context.sync();
                    if (!ch.isNullObject) {
                        ch.delete();
                        await context.sync();
                    }
                    redoEntry = {
                        kind: "chart_create",
                        sheet: snapshot.sheet,
                        chartName: snapshot.chartName,
                        chartDef: snapshot.chartDef
                    };
                    return `Removed chart "${snapshot.chartName}".`;
                }

                if (kind === "chart_delete") {
                    if (snapshot.chartDef) {
                        await this.restoreChartFromDef(context, snapshot.sheet, snapshot.chartDef);
                    }
                    redoEntry = {
                        kind: "chart_delete",
                        sheet: snapshot.sheet,
                        chartName: snapshot.chartName
                    };
                    return `Restored chart "${snapshot.chartName}".`;
                }

                if (kind === "chart_delete_all") {
                    for (const cd of (snapshot.chartDefs || [])) {
                        await this.restoreChartFromDef(context, snapshot.sheet, cd);
                    }
                    return `Restored ${(snapshot.chartDefs || []).length} chart(s).`;
                }

                if (kind === "rename" && snapshot.from) {
                    sheets.getItem(snapshot.sheet).name = this.sanitizeSheetName(snapshot.from);
                    await context.sync();
                    return `Renamed "${snapshot.sheet}" back to "${snapshot.from}".`;
                }

                if (kind === "visibility") {
                    const ws = sheets.getItem(snapshot.sheet);
                    ws.visibility = snapshot.from || "Visible";
                    await context.sync();
                    return `Restored visibility of "${snapshot.sheet}".`;
                }

                if (kind === "add") {
                    sheets.getItem(snapshot.sheet).delete();
                    await context.sync();
                    return `Removed auto-created sheet "${snapshot.sheet}".`;
                }

                if (kind === "freeze") {
                    const ws = snapshot.sheet ? sheets.getItem(snapshot.sheet)
                        : context.workbook.worksheets.getActiveWorksheet();
                    ws.freezePanes.unfreeze();
                    if (snapshot.from) {
                        try {
                            const cell = String(snapshot.from).split("!").pop().split(":")[0];
                            const m = /^([A-Za-z]+)(\d+)$/.exec(cell.trim());
                            if (m) {
                                const r = parseInt(m[2], 10) - 1;
                                const c = this.columnIndex(m[1]);
                                if (r > 0) ws.freezePanes.freezeRows(r);
                                if (c > 0) ws.freezePanes.freezeColumns(c);
                            }
                        } catch (_) {}
                    }
                    await context.sync();
                    return `Restored freeze panes on "${snapshot.sheet || "active sheet"}".`;
                }

                if (kind === "structural") {
                    const target = snapshot.sheet;
                    // Bug 10: Use the explicit backupName stored in the undo entry.
                    const rec = Object.assign({},
                        (this._backupRecords || {})[target] || {},
                        snapshot.backup ? { backup: snapshot.backup } : {}
                    );
                    return await this.rollbackSheetToBackup(context, target, rec);
                }

                // kind === "range": restore content + format + borders + merge + widths.
                if (!snapshot.formulas || !snapshot.range) {
                    throw new Error("Undo entry has no snapshot data.");
                }
                const sheet = snapshot.sheet
                    ? sheets.getItem(snapshot.sheet)
                    : context.workbook.worksheets.getActiveWorksheet();
                const targetRange = sheet.getRange(snapshot.range);

                // Bug 31: Delete any table that was created by this action.
                if (snapshot.tableCreated) {
                    try {
                        const tName = typeof snapshot.tableCreated === "string"
                            ? snapshot.tableCreated : null;
                        if (tName) {
                            const tbl = sheet.tables.getItemOrNullObject(tName);
                            await context.sync();
                            if (!tbl.isNullObject) { tbl.delete(); await context.sync(); }
                        }
                    } catch (_) {}
                }

                targetRange.formulas = snapshot.formulas;
                if (snapshot.numberFormat) {
                    try { targetRange.numberFormat = snapshot.numberFormat; } catch (_) {}
                }
                const f = snapshot.format || {};
                try {
                    const ff = f.font || {};
                    if (ff.bold !== undefined && ff.bold !== null) targetRange.format.font.bold = ff.bold;
                    if (ff.italic !== undefined && ff.italic !== null) targetRange.format.font.italic = ff.italic;
                    if (ff.color !== undefined && ff.color !== null) targetRange.format.font.color = ff.color;
                    if (ff.size !== undefined && ff.size !== null) targetRange.format.font.size = ff.size;
                    if (ff.name !== undefined && ff.name !== null) targetRange.format.font.name = ff.name;
                    if (f.fillColor !== undefined && f.fillColor !== null) targetRange.format.fill.color = f.fillColor;
                    if (f.horizontalAlignment) targetRange.format.horizontalAlignment = f.horizontalAlignment;
                    if (f.verticalAlignment) targetRange.format.verticalAlignment = f.verticalAlignment;
                    if (f.wrapText !== undefined && f.wrapText !== null) targetRange.format.wrapText = f.wrapText;
                } catch (_) {}

                // Bug 33: Restore border state.
                if (snapshot.borders) {
                    try {
                        for (const [en, bd] of Object.entries(snapshot.borders)) {
                            if (!bd) continue;
                            try {
                                const b = targetRange.format.borders.getItem(en);
                                if (bd.style) b.style = bd.style;
                                if (bd.weight) b.weight = bd.weight;
                                if (bd.color) b.color = bd.color;
                            } catch (_) {}
                        }
                    } catch (_) {}
                }

                // Bug 34: Restore merge state.
                if (snapshot.mergeAddress !== undefined) {
                    try {
                        targetRange.unmerge();
                        if (snapshot.mergeAddress) {
                            const localMerge = this.localAddress(String(snapshot.mergeAddress));
                            sheet.getRange(localMerge).merge(false);
                        }
                    } catch (_) {}
                }

                // Bug 35: Restore column widths and row heights.
                if (snapshot.colWidths !== null && snapshot.colWidths !== undefined) {
                    try { targetRange.format.columnWidth = snapshot.colWidths; } catch (_) {}
                }
                if (snapshot.rowHeights !== null && snapshot.rowHeights !== undefined) {
                    try { targetRange.format.rowHeight = snapshot.rowHeights; } catch (_) {}
                }

                await context.sync();
                const where = snapshot.sheet ? snapshot.sheet + "!" : "";
                return `Reverted ${where}${snapshot.range} to previous state.`;
            });

            // Push redo entry (for range and chart undos).
            if (redoEntry) {
                this.redoStack.push(redoEntry);
                if (this.redoStack.length > 50) this.redoStack.shift();
            }
            this.persistUndoMeta();
            return { success: true, message: msg, remaining: this.undoStack.length, redoAvailable: this.redoStack.length > 0 };
        } catch (err) {
            // Bug 11: Re-push the snapshot so the user can retry.
            this.undoStack.push(snapshot);
            return { success: false, message: `Undo failed: ${err.message}`, remaining: this.undoStack.length };
        }
    },

    /**
     * Redo the last undone change (supports cell values/formats and chart creations/deletions).
     */
    async redoLast() {
        if (this.redoStack.length === 0) {
            return { success: false, message: "Nothing to redo. Redo is only available after an Undo, for cell-value/format changes or charts." };
        }

        const snapshot = this.redoStack.pop();
        const kind = snapshot.kind || "range";

        if (!this.isOfficeInitialized) {
            const where = snapshot.sheet ? snapshot.sheet + "!" : "";
            return { success: true, message: `[Dev Mode] Redo: ${kind} ${where}${snapshot.range || snapshot.chartName || ""}`.trim() };
        }

        if (kind === "chart_create") {
            try {
                const msg = await Excel.run(async (context) => {
                    await this.restoreChartFromDef(context, snapshot.sheet, snapshot.chartDef);
                    this.undoStack.push({
                        kind: "chart_create",
                        sheet: snapshot.sheet,
                        chartName: snapshot.chartName,
                        chartDef: snapshot.chartDef
                    });
                    this.persistUndoMeta();
                    return `Re-created chart "${snapshot.chartName}".`;
                });
                return { success: true, message: msg, remaining: this.redoStack.length };
            } catch (err) {
                this.redoStack.push(snapshot);
                return { success: false, message: `Redo failed: ${err.message}`, remaining: this.redoStack.length };
            }
        }

        if (kind === "chart_delete") {
            try {
                const msg = await Excel.run(async (context) => {
                    const sheets = context.workbook.worksheets;
                    const ws = snapshot.sheet ? sheets.getItem(snapshot.sheet) : sheets.getActiveWorksheet();
                    const ch = ws.charts.getItemOrNullObject(snapshot.chartName);
                    await context.sync();
                    if (!ch.isNullObject) {
                        ch.delete();
                        await context.sync();
                    }
                    this.undoStack.push({
                        kind: "chart_delete",
                        sheet: snapshot.sheet,
                        chartName: snapshot.chartName
                    });
                    this.persistUndoMeta();
                    return `Re-deleted chart "${snapshot.chartName}".`;
                });
                return { success: true, message: msg, remaining: this.redoStack.length };
            } catch (err) {
                this.redoStack.push(snapshot);
                return { success: false, message: `Redo failed: ${err.message}`, remaining: this.redoStack.length };
            }
        }

        if (kind !== "range") {
            this.redoStack.push(snapshot);
            return {
                success: false,
                message: `Redo not supported for "${kind}" changes. Ask the AI to re-apply if needed.`
            };
        }

        try {
            const msg = await Excel.run(async (context) => {
                const sheets = context.workbook.worksheets;
                const sheet = snapshot.sheet
                    ? sheets.getItem(snapshot.sheet)
                    : context.workbook.worksheets.getActiveWorksheet();
                const targetRange = sheet.getRange(snapshot.range);

                // Capture current state as a new undo entry so this redo can itself be undone.
                targetRange.load(["values", "formulas", "numberFormat"]);
                try {
                    targetRange.format.font.load(["bold", "italic", "color", "size", "name"]);
                    targetRange.format.fill.load("color");
                    targetRange.format.load(["horizontalAlignment", "verticalAlignment", "wrapText"]);
                } catch (_) {}
                await context.sync();

                const newUndoEntry = {
                    kind: "range",
                    sheet: snapshot.sheet,
                    range: snapshot.range,
                    values: JSON.parse(JSON.stringify(targetRange.values)),
                    formulas: JSON.parse(JSON.stringify(targetRange.formulas)),
                    numberFormat: targetRange.numberFormat ? JSON.parse(JSON.stringify(targetRange.numberFormat)) : null
                };
                try {
                    newUndoEntry.format = {
                        font: {
                            bold: targetRange.format.font.bold,
                            italic: targetRange.format.font.italic,
                            color: targetRange.format.font.color,
                            size: targetRange.format.font.size,
                            name: targetRange.format.font.name
                        },
                        fillColor: targetRange.format.fill.color,
                        horizontalAlignment: targetRange.format.horizontalAlignment,
                        verticalAlignment: targetRange.format.verticalAlignment,
                        wrapText: targetRange.format.wrapText
                    };
                } catch (_) {}

                // Apply the redo (re-apply the "after" state that was undone).
                targetRange.formulas = snapshot.formulas;
                if (snapshot.numberFormat) {
                    try { targetRange.numberFormat = snapshot.numberFormat; } catch (_) {}
                }
                const f = snapshot.format || {};
                try {
                    const ff = f.font || {};
                    if (ff.bold !== undefined && ff.bold !== null) targetRange.format.font.bold = ff.bold;
                    if (ff.italic !== undefined && ff.italic !== null) targetRange.format.font.italic = ff.italic;
                    if (ff.color !== undefined && ff.color !== null) targetRange.format.font.color = ff.color;
                    if (ff.size !== undefined && ff.size !== null) targetRange.format.font.size = ff.size;
                    if (ff.name !== undefined && ff.name !== null) targetRange.format.font.name = ff.name;
                    if (f.fillColor !== undefined && f.fillColor !== null) targetRange.format.fill.color = f.fillColor;
                    if (f.horizontalAlignment) targetRange.format.horizontalAlignment = f.horizontalAlignment;
                    if (f.verticalAlignment) targetRange.format.verticalAlignment = f.verticalAlignment;
                    if (f.wrapText !== undefined && f.wrapText !== null) targetRange.format.wrapText = f.wrapText;
                } catch (_) {}
                await context.sync();

                // Push back to undo stack so this redo can itself be undone.
                this.undoStack.push(newUndoEntry);
                if (this.undoStack.length > 100) this.undoStack.splice(0, this.undoStack.length - 100);
                this.persistUndoMeta();

                const where = snapshot.sheet ? snapshot.sheet + "!" : "";
                return `Re-applied ${where}${snapshot.range}.`;
            });
            return { success: true, message: msg, remaining: this.redoStack.length };
        } catch (err) {
            // Re-push on failure so user can retry.
            this.redoStack.push(snapshot);
            return { success: false, message: `Redo failed: ${err.message}`, remaining: this.redoStack.length };
        }
    }
};

window.ExcelBridge = ExcelBridge;
