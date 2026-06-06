import { useTemplateInspectorStore } from '../../stores/templateInspectorStore.js';
import { ref, computed } from 'vue';

export default {
    name: 'TemplateGrid',
    template: `
        <div class="template-grid-section">
            <!-- Sheet Selector -->
            <div class="sheet-tabs mb-4 flex gap-2">
                <button v-for="(sheetData, sheetName) in store.templateLayout" :key="sheetName"
                        class="px-4 py-2 rounded-lg text-sm font-medium transition-colors border border-transparent" 
                        :class="store.currentSheetName === sheetName ? 'bg-blue-600 text-white' : 'bg-slate-700 text-slate-300 hover:bg-slate-600'"
                        @click="store.currentSheetName = sheetName">
                    {{ sheetName }}
                </button>
            </div>
            
            <!-- Zoom & View Controls -->
            <div class="mb-2 flex gap-2 items-center">
                <button class="px-3 py-1 bg-slate-700 hover:bg-slate-600 text-slate-200 rounded transition-colors" @click="zoomOut" title="Zoom Out">-</button>
                <span class="text-sm text-center min-w-12 text-slate-300">{{ zoomPercentage }}%</span>
                <button class="px-3 py-1 bg-slate-700 hover:bg-slate-600 text-slate-200 rounded transition-colors" @click="zoomIn" title="Zoom In">+</button>
                <button class="px-3 py-1 bg-slate-700 hover:bg-slate-600 text-slate-200 rounded transition-colors" @click="resetZoom" title="Reset Zoom">Reset</button>

                <div class="w-px h-6 bg-slate-300 mx-2"></div>

                <label class="flex items-center gap-1 text-sm select-none cursor-pointer text-slate-300">
                    <input type="checkbox" v-model="showFullText"> Wrap Text
                </label>
            </div>
            
            <!-- Excel Grid -->
            <div class="excel-grid-container overflow-auto max-h-75vh relative border border-slate-700 rounded-xl bg-slate-900/50 p-2">
                <div class="excel-grid" :style="gridStyle">
                    <!-- Render Cells -->
                    <div v-for="cell in gridCells" :key="cell.id"
                         class="excel-cell"
                         :style="cell.style"
                         :title="'[' + cell.address + '] ' + cell.content"
                         @click="openCellEditor(cell)">
                         <span v-if="cell.hasOverride" class="absolute bg-blue-500 rounded-full" style="width: 6px; height: 6px; top: 2px; right: 2px;" title="Has mode override"></span>
                         <span v-if="cell.isFormula" class="text-blue-600 italic">{{ cell.content }}</span>
                         <span v-else>{{ cell.content }}</span>
                    </div>
                </div>
            </div>

            <!-- Cell Override Editor Popup -->
            <teleport to="body">
                <div v-if="editingCell" class="fixed inset-0 bg-black/60 z-[100] flex items-center justify-center backdrop-blur-sm" @click.self="closeEditor">
                    <div class="bg-slate-900 border border-slate-700 rounded-2xl p-6 shadow-2xl min-w-[400px] text-slate-100">
                        <h3 class="m-0 mb-3 text-base text-blue-400 font-bold">Cell {{ editingCell.address }}</h3>

                        <div class="mb-4 p-4 bg-slate-800 rounded-lg text-sm border border-slate-700/50">
                            <span class="text-slate-400">Current (default):</span>
                            <span class="text-slate-200 ml-2 font-mono">
                                {{ (typeof editingCell.rawContent === 'object' && editingCell.rawContent !== null) ? (editingCell.rawContent.default ?? "") : (editingCell.rawContent || '(empty)') }}
                            </span>
                        </div>

                        <div class="mb-3">
                            <label class="block text-slate-300 text-sm mb-1 font-medium">Base Value <span class="text-blue-400 text-xs">(applies to ALL modes)</span></label>
                            <input type="text" v-model="editStandardValue" class="w-full bg-slate-950 border border-slate-700 rounded-lg px-4 py-2 text-slate-100 focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 transition-all text-sm" placeholder="Enter value for standard, custom, DAF..." @keyup.enter="saveCellOverrides" />
                            <p class="text-slate-500 text-xs mt-1 mb-0">This value will be used in Standard, Custom, DAF, and any other mode.</p>
                        </div>

                        <div class="mb-3">
                            <label class="block text-slate-300 text-sm mb-1 font-medium">DAF Override <span class="text-amber-400 text-xs">(takes priority in DAF mode)</span></label>
                            <input type="text" v-model="editDafValue" class="w-full bg-slate-950 border border-slate-700 rounded-lg px-4 py-2 text-slate-100 focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 transition-all text-sm" placeholder="Leave empty to use base value" @keyup.enter="saveCellOverrides" />
                            <p class="text-slate-500 text-xs mt-1 mb-0">Only used when generating in DAF mode. If empty, the base value is used.</p>
                        </div>

                        <div v-if="editingCell.currentOverrides" class="mb-3 px-3 py-2 bg-blue-950/50 border border-blue-900/50 rounded text-sm">
                            <div class="text-blue-400 mb-1 font-medium">Existing overrides:</div>
                            <div v-for="(v, k) in editingCell.currentOverrides" :key="k" class="text-slate-300 font-mono text-xs">
                                <strong>{{ k }}:</strong> {{ v }}
                            </div>
                        </div>

                        <div class="flex gap-3 justify-end mt-6">
                            <button class="px-4 py-2 bg-slate-700 hover:bg-slate-600 text-white rounded-lg transition-colors text-sm font-medium" @click="closeEditor">Cancel</button>
                            <button class="px-4 py-2 bg-blue-600 hover:bg-blue-500 text-white rounded-lg transition-colors shadow-lg shadow-blue-500/20 text-sm font-medium" @click="saveCellOverrides" :disabled="isSavingCell">
                                {{ isSavingCell ? 'Saving...' : 'Save Overrides' }}
                            </button>
                        </div>
                        <div v-if="editorMessage" class="mt-2 text-sm text-center font-medium" :class="editorMessageType === 'error' ? 'text-red-400' : 'text-emerald-400'">
                            {{ editorMessage }}
                        </div>
                        <div class="mt-4 p-3 bg-amber-500/10 border border-amber-500/30 rounded-lg text-xs text-amber-400 leading-snug">
                            ⚠ If footer overrides appear shifted after re-generating, the Excel template structure likely changed (rows added/removed). Re-apply overrides after verifying cell positions or delete the template to create a new one.
                        </div>
                    </div>
                </div>
            </teleport>
        </div>
    `,
    setup() {
        const store = useTemplateInspectorStore();
        const zoomLevel = ref(0.6);
        const showFullText = ref(false);

        // Cell override editor state
        const editingCell = ref(null);
        const editStandardValue = ref("");
        const editDafValue = ref("");
        const isSavingCell = ref(false);
        const editorMessage = ref("");
        const editorMessageType = ref("success");

        const zoomIn = () => {
            zoomLevel.value = Math.min(zoomLevel.value + 0.1, 3.0);
        };
        const zoomOut = () => {
            zoomLevel.value = Math.max(zoomLevel.value - 0.1, 0.2);
        };
        const resetZoom = () => {
            zoomLevel.value = 1.0;
        };

        const zoomPercentage = computed(() => Math.round(zoomLevel.value * 100));

        const currentSheetData = computed(() => {
            if (!store.currentSheetName || !store.templateLayout) return null;
            return store.templateLayout[store.currentSheetName];
        });

        // --- Helper functions for Grid ---

        const parseAddress = (addr) => {
            const match = addr.match(/([A-Z]+)([0-9]+)/);
            if (!match) return { row: 0, col: 0 };
            const colStr = match[1];
            const rowStr = match[2];

            let col = 0;
            for (let i = 0; i < colStr.length; i++) {
                col = col * 26 + (colStr.charCodeAt(i) - 64);
            }
            return { row: parseInt(rowStr) - 1, col: col - 1 };
        };

        const colToLetter = (c) => {
            let colLetter = "";
            let tempCol = c + 1;
            while (tempCol > 0) {
                let rem = (tempCol - 1) % 26;
                colLetter = String.fromCharCode(65 + rem) + colLetter;
                tempCol = Math.floor((tempCol - 1) / 26);
            }
            return colLetter;
        };

        const excelColorToCss = (colorStr) => {
            if (!colorStr || typeof colorStr !== 'string') return null;
            let hex = colorStr.replace(/^#/, '');
            if (hex.length === 8) hex = hex.substring(2); // Strip alpha
            if (hex.length === 6 && /^[0-9A-Fa-f]{6}$/.test(hex)) return `#${hex}`;
            return null;
        };

        const excelBorderToCss = (style) => {
            const map = {
                'thin': '1px solid #000',
                'medium': '2px solid #000',
                'thick': '3px solid #000',
                'double': '3px double #000',
                'dashed': '1px dashed #555',
                'dotted': '1px dotted #555',
                'hair': '1px solid #bbb',
                'mediumDashed': '2px dashed #000',
                'dashDot': '1px dashed #555',
                'mediumDashDot': '2px dashed #000'
            };
            return map[style] || '1px solid #ccc';
        };

        const flattenStyles = (stylesRaw, stylePalette) => {
            const result = {};
            for (const [key, value] of Object.entries(stylesRaw)) {
                if (Array.isArray(value)) {
                    const resolved = stylePalette[key] || {};
                    for (const coord of value) {
                        result[coord] = resolved;
                    }
                } else if (typeof value === 'object' && value !== null) {
                    result[key] = value;
                } else if (typeof value === 'string') {
                    result[key] = stylePalette[value] || {};
                }
            }
            return result;
        };

        const buildCellCss = (r, c, cellStyle, mergeInfo, isEmpty) => {
            const defaultBorder = '1px solid #e2e8f0';
            let borderTop = defaultBorder;
            let borderRight = defaultBorder;
            let borderBottom = defaultBorder;
            let borderLeft = defaultBorder;

            if (cellStyle.border) {
                if (cellStyle.border.top) borderTop = excelBorderToCss(cellStyle.border.top);
                if (cellStyle.border.right) borderRight = excelBorderToCss(cellStyle.border.right);
                if (cellStyle.border.bottom) borderBottom = excelBorderToCss(cellStyle.border.bottom);
                if (cellStyle.border.left) borderLeft = excelBorderToCss(cellStyle.border.left);
            }

            let bgColor = (isEmpty && !cellStyle.fill?.color) ? 'transparent' : '#fff';
            if (cellStyle.fill?.color) {
                const parsed = excelColorToCss(cellStyle.fill.color);
                if (parsed) bgColor = parsed;
            }

            let fontColor = '#000';
            if (cellStyle.font?.color) {
                const parsed = excelColorToCss(cellStyle.font.color);
                if (parsed) fontColor = parsed;
            }

            const hAlignMap = { left: 'flex-start', center: 'center', right: 'flex-end', general: 'flex-start', justify: 'flex-start' };
            const vAlignMap = { top: 'flex-start', center: 'center', bottom: 'flex-end' };
            const hAlign = cellStyle.alignment?.horizontal || 'left';
            const vAlign = cellStyle.alignment?.vertical || 'bottom';

            return {
                gridColumnStart: c + 1,
                gridColumnEnd: mergeInfo ? c + 1 + mergeInfo.colspan : c + 2,
                gridRowStart: r + 1,
                gridRowEnd: mergeInfo ? r + 1 + mergeInfo.rowspan : r + 2,
                display: 'flex',
                justifyContent: hAlignMap[hAlign] || 'flex-start',
                alignItems: vAlignMap[vAlign] || 'flex-end',
                textAlign: hAlign === 'center' ? 'center' : (hAlign === 'right' ? 'right' : 'left'),
                borderTop,
                borderRight,
                borderBottom,
                borderLeft,
                padding: '1px 2px',
                fontSize: (cellStyle.font?.size || 11) + 'pt',
                fontWeight: cellStyle.font?.bold ? 'bold' : 'normal',
                fontStyle: cellStyle.font?.italic ? 'italic' : 'normal',
                fontFamily: cellStyle.font?.name || 'Arial, sans-serif',
                backgroundColor: bgColor,
                whiteSpace: (showFullText.value || cellStyle.alignment?.wrap_text) ? 'normal' : 'nowrap',
                wordBreak: (showFullText.value || cellStyle.alignment?.wrap_text) ? 'break-word' : 'normal',
                color: fontColor,
                lineHeight: '1.2'
            };
        };

        const sheetBounds = computed(() => {
            const sheet = currentSheetData.value;
            if (!sheet) return { maxRow: 0, maxCol: 0, headerMaxRow: 0, footerBaseRow: 0 };

            const content = sheet.template_header_content || sheet.header_content || {};
            const stylePalette = sheet.style_palette || {};
            const styles = flattenStyles(sheet.template_header_styles || sheet.header_styles || {}, stylePalette);
            const mergesRaw = sheet.template_header_merges || sheet.header_merges || {};
            const merges = Array.isArray(mergesRaw) ? mergesRaw : Object.keys(mergesRaw);
            const footerRows = sheet.template_footer_rows || sheet.footer_rows || [];

            let headerMaxRow = 0;
            let maxCol = 0;
            Object.keys(content).forEach(addr => {
                const { row, col } = parseAddress(addr);
                if (row > headerMaxRow) headerMaxRow = row;
                if (col > maxCol) maxCol = col;
            });
            Object.keys(styles).forEach(addr => {
                const { row, col } = parseAddress(addr);
                if (row > headerMaxRow) headerMaxRow = row;
                if (col > maxCol) maxCol = col;
            });
            merges.forEach(range => {
                const parts = range.split(":");
                if (parts.length === 2) {
                    const e = parseAddress(parts[1]);
                    if (e.row > headerMaxRow) headerMaxRow = e.row;
                    if (e.col > maxCol) maxCol = e.col;
                }
            });

            const footerBaseRow = headerMaxRow + 1;
            let maxRow = headerMaxRow;

            footerRows.forEach(rowDict => {
                const absRow = footerBaseRow + (rowDict.relative_index ?? 0);
                if (absRow > maxRow) maxRow = absRow;
                for (const cellDict of (rowDict.cells || [])) {
                    const ci = (cellDict.col_index || 1) - 1;
                    if (ci > maxCol) maxCol = ci;
                }
                for (const m of (rowDict.merges || [])) {
                    const mc = (m.max_col || 1) - 1;
                    if (mc > maxCol) maxCol = mc;
                }
            });

            return { maxRow: maxRow + 2, maxCol: maxCol + 2, headerMaxRow, footerBaseRow };
        });

        const gridCells = computed(() => {
            if (!currentSheetData.value) return [];

            const sheet = currentSheetData.value;
            const content = sheet.template_header_content || sheet.header_content || {};
            const stylePalette = sheet.style_palette || {};
            const styles = flattenStyles(sheet.template_header_styles || sheet.header_styles || {}, stylePalette);
            const mergesRaw = sheet.template_header_merges || sheet.header_merges || {};
            const merges = Array.isArray(mergesRaw) ? mergesRaw : Object.keys(mergesRaw);

            const footerContent = {};
            const footerStyles = {};
            const footerMergeRanges = [];
            const footerRows = sheet.template_footer_rows || sheet.footer_rows || [];

            const { maxRow, maxCol, footerBaseRow } = sheetBounds.value;

            footerRows.forEach(rowDict => {
                const relIdx = rowDict.relative_index ?? 0;
                const absRow = footerBaseRow + relIdx;

                for (const cellDict of (rowDict.cells || [])) {
                    const colIdx = cellDict.col_index;
                    const addr = `${colToLetter(colIdx - 1)}${absRow + 1}`;

                    if (cellDict.value !== undefined && cellDict.value !== null) {
                        footerContent[addr] = cellDict.value;
                    }
                    if (cellDict.style_id) {
                        footerStyles[addr] = stylePalette[cellDict.style_id] || {};
                    }
                }

                for (const mDict of (rowDict.merges || [])) {
                    const minCol = mDict.min_col;
                    const maxColM = mDict.max_col;
                    const rowSpan = mDict.row_span || 1;
                    const startAddr = `${colToLetter(minCol - 1)}${absRow + 1}`;
                    const endAddr = `${colToLetter(maxColM - 1)}${absRow + rowSpan}`;
                    footerMergeRanges.push(`${startAddr}:${endAddr}`);
                }
            });

            const allContent = { ...content, ...footerContent };
            const allStyles = { ...styles, ...footerStyles };
            const allMerges = [...merges, ...footerMergeRanges];

            const cells = [];
            const occupied = new Set();
            const mergedRanges = {};

            allMerges.forEach(range => {
                const parts = range.split(":");
                if (parts.length !== 2) return;
                const [start, end] = parts;
                const s = parseAddress(start);
                const e = parseAddress(end);
                mergedRanges[start] = { rowspan: e.row - s.row + 1, colspan: e.col - s.col + 1 };

                for (let r = s.row; r <= e.row; r++) {
                    for (let c = s.col; c <= e.col; c++) {
                        if (r !== s.row || c !== s.col) {
                            occupied.add(`${r},${c}`);
                        }
                    }
                }
            });

            for (let r = 0; r <= maxRow; r++) {
                for (let c = 0; c <= maxCol; c++) {
                    if (occupied.has(`${r},${c}`)) continue;

                    const address = `${colToLetter(c)}${r + 1}`;
                    const cellContent = allContent[address] || "";
                    const cellStyle = allStyles[address] || {};
                    const mergeInfo = mergedRanges[address];

                    cells.push({
                        id: address,
                        address: address,
                        content: typeof cellContent === 'object' && cellContent !== null
                            ? (cellContent.default !== undefined && cellContent.default !== null ? cellContent.default : (Object.keys(cellContent).length > 0 ? JSON.stringify(cellContent) : ""))
                            : (cellContent || ""),
                        rawContent: cellContent,
                        hasOverride: typeof cellContent === 'object' && cellContent !== null,
                        currentOverrides: (typeof cellContent === 'object' && cellContent !== null) ? cellContent : null,
                        style: { ...buildCellCss(r, c, cellStyle, mergeInfo, !cellContent), position: 'relative', cursor: 'pointer' },
                        isFormula: typeof cellContent === 'string' && cellContent.startsWith('=')
                    });
                }
            }
            return cells;
        });

        const gridStyle = computed(() => {
            const sheet = currentSheetData.value;
            const base = {
                display: 'grid',
                gap: '0',
                backgroundColor: '#f1f5f9',
                transform: `scale(${zoomLevel.value})`,
                transformOrigin: 'top left',
                width: 'fit-content'
            };
            if (!sheet) return base;

            const { maxRow, maxCol, footerBaseRow } = sheetBounds.value;
            const colWidthsMap = sheet.col_widths || {};
            const rowHeightsMap = sheet.template_header_row_heights || sheet.header_row_heights || {};
            const footerRows = sheet.template_footer_rows || sheet.footer_rows || [];

            const cols = [];
            for (let c = 0; c <= maxCol; c++) {
                const letter = colToLetter(c);
                const w = colWidthsMap[letter];
                cols.push(w ? Math.max(Math.round(w * 7.5), 20) + 'px' : '64px');
            }
            base.gridTemplateColumns = cols.join(' ');

            const footerHeightLookup = {};
            footerRows.forEach(rowDict => {
                if (rowDict.height != null) {
                    footerHeightLookup[footerBaseRow + (rowDict.relative_index ?? 0)] = rowDict.height;
                }
            });
            const rows = [];
            for (let r = 0; r <= maxRow; r++) {
                const hdrH = rowHeightsMap[String(r + 1)];
                const ftrH = footerHeightLookup[r];
                const h = hdrH || ftrH;
                rows.push(h ? Math.max(Math.round(h * 1.333), 14) + 'px' : '20px');
            }
            base.gridTemplateRows = rows.join(' ');

            return base;
        });

        // --- Cell override editor triggers ---

        const openCellEditor = (cell) => {
            editingCell.value = cell;
            editStandardValue.value = (cell.currentOverrides && cell.currentOverrides.standard) || "";
            editDafValue.value = (cell.currentOverrides && cell.currentOverrides.daf) || "";
            editorMessage.value = "";
        };

        const closeEditor = () => {
            editingCell.value = null;
            editStandardValue.value = "";
            editDafValue.value = "";
            editorMessage.value = "";
        };

        const saveCellOverrides = async () => {
            if (!editingCell.value) return;
            isSavingCell.value = true;
            editorMessage.value = "";

            const res = await store.saveCellOverrides(
                editingCell.value.address,
                editStandardValue.value,
                editDafValue.value
            );

            if (res.success) {
                editorMessage.value = "Saved!";
                editorMessageType.value = "success";
                setTimeout(() => {
                    closeEditor();
                }, 500);
            } else {
                editorMessage.value = res.error || "Save failed";
                editorMessageType.value = "error";
            }
            isSavingCell.value = false;
        };

        return {
            store,
            zoomPercentage,
            zoomIn,
            zoomOut,
            resetZoom,
            showFullText,
            gridStyle,
            gridCells,
            editingCell,
            editStandardValue,
            editDafValue,
            isSavingCell,
            editorMessage,
            editorMessageType,
            openCellEditor,
            closeEditor,
            saveCellOverrides
        };
    }
};
