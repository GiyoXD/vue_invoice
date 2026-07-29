import { useTemplateInspectorStore } from '../../stores/templateInspectorStore.js';
import { ref, computed } from 'vue';

export default {
    name: 'TemplateGrid',
    template: `
        <div class="template-grid-section relative">
            <!-- Grid Selection Mode Banner -->
            <div v-if="selectionTarget" class="fixed top-4 left-1/2 -translate-x-1/2 bg-blue-600 border border-blue-500 text-white px-6 py-3 rounded-full shadow-2xl z-[200] flex items-center gap-4 animate-bounce">
                <span class="font-semibold flex items-center gap-2">
                    <svg class="animate-spin h-5 w-5 text-white" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24">
                        <circle class="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" stroke-width="4"></circle>
                        <path class="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path>
                    </svg>
                    Selecting reference for {{ selectionTarget === 'standard' ? 'Base Value' : 'DAF Override' }}...
                </span>
                <span class="text-xs text-blue-200 bg-blue-700/50 px-2 py-0.5 rounded">Click any cell to select</span>
                <button @click="cancelCellSelection" class="px-3 py-1 bg-slate-800 hover:bg-slate-700 text-slate-200 text-xs font-bold rounded-full transition-colors border border-slate-700">Cancel</button>
            </div>

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

                <div class="w-px h-6 bg-slate-700 mx-2"></div>

                <label class="flex items-center gap-1 text-sm select-none cursor-pointer text-slate-300">
                    <input type="checkbox" v-model="showDummyData"> Show Dummy Rows
                </label>
            </div>
            
            <!-- Flex Container for Grid and Sidebar -->
            <div class="flex gap-6 mt-4">
                <!-- Left: Excel Grid -->
                <div class="flex-1 min-w-0" style="align-self: flex-start;">
                    <div class="excel-grid-container overflow-auto max-h-75vh relative border border-slate-700 rounded-xl bg-slate-900/50 p-2">
                        <div class="excel-grid" :style="gridStyle">
                            <!-- Render Cells -->
                            <div v-for="cell in gridCells" :key="cell.id"
                                 class="excel-cell transition-all"
                                 :class="{ 'hover:ring-2 hover:ring-blue-500 hover:z-10 cursor-crosshair': selectionTarget }"
                                 :style="cell.style"
                                 :title="cell.title"
                                 @click="handleCellClick(cell)">
                                 <span v-if="cell.hasOverride" class="absolute bg-blue-500 rounded-full" style="width: 6px; height: 6px; top: 2px; right: 2px;" title="Has mode override"></span>
                                 <template v-if="cell.hasOverride">
                                     <span>{{ cell.content }}</span>
                                     <span v-if="cell.hasStandardOverride" class="text-emerald-500 font-bold ml-1" style="font-size: 10px; word-break: break-all;" title="Base Override">→ {{ cell.standardOverride }}</span>
                                     <span v-if="cell.hasDafOverride" class="text-red-500 font-bold ml-1" style="font-size: 10px; word-break: break-all;" title="DAF Override">→ {{ cell.dafOverride }}</span>
                                 </template>
                                 <span v-else-if="cell.isFormula" class="text-blue-600 italic">{{ cell.content }}</span>
                                 <span v-else>{{ cell.content }}</span>
                            </div>
                        </div>
                    </div>
                </div>

                <!-- Right: Cell Editor Dock Panel -->
                <div class="bg-slate-900 border border-slate-700 rounded-xl p-4 text-slate-100 flex flex-col gap-4" style="width: 320px; min-width: 320px; flex-shrink: 0; align-self: flex-start;">
                    <div v-if="editingCell">
                        <div class="flex justify-between items-center mb-3">
                            <h3 class="m-0 text-base text-blue-400 font-bold">Cell {{ editingCell.address }}</h3>
                            <button @click="closeEditor" class="text-slate-400 hover:text-white transition-colors">✕</button>
                        </div>

                        <div class="mb-3 p-3 bg-slate-800 rounded-lg text-xs border border-slate-700/50">
                            <span class="text-slate-400">Current (default):</span>
                            <span class="text-slate-200 ml-2 font-mono break-words">
                                {{ (typeof editingCell.rawContent === 'object' && editingCell.rawContent !== null) ? (editingCell.rawContent.default ?? "") : (editingCell.rawContent || '(empty)') }}
                            </span>
                        </div>

                        <div class="mb-3">
                            <label class="block text-slate-300 text-xs mb-1 font-medium">Base Value <span class="text-blue-400 font-medium" style="font-size: 10px;">(applies to ALL modes)</span></label>
                            <div class="flex gap-1.5">
                                <input type="text" v-model="editStandardValue" class="flex-1 bg-slate-950 border border-slate-700 rounded-lg px-3 py-1.5 text-red-400 font-medium focus:outline-none focus:border-red-500 focus:ring-1 focus:ring-red-500 transition-all text-xs" placeholder="Value or formula..." @keyup.enter="saveCellOverrides" />
                                <button type="button" @click="startCellSelection('standard')" class="p-1.5 bg-slate-800 hover:bg-slate-700 border border-slate-700 rounded-lg text-blue-400 hover:text-blue-300 transition-colors flex-shrink-0" title="Pick cell from sheet">
                                    <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" class="lucide lucide-crosshair"><circle cx="12" cy="12" r="10"/><line x1="22" y1="12" x2="18" y2="12"/><line x1="6" y1="12" x2="2" y2="12"/><line x1="12" y1="6" x2="12" y2="2"/><line x1="12" y1="22" x2="12" y2="18"/></svg>
                                </button>
                            </div>
                            <!-- Chips -->
                            <div class="mt-1.5 flex flex-wrap gap-1 items-center">
                                <span class="text-slate-500 font-medium" style="font-size: 10px;">Insert:</span>
                                <button type="button" @click="insertRef('standard', '=DeepSheet!B1')" class="px-1.5 py-0.5 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded transition-colors border border-slate-700/50" style="font-size: 10px;">=Ref (B1)</button>
                                <button type="button" @click="insertRef('standard', '=DeepSheet!B2')" class="px-1.5 py-0.5 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded transition-colors border border-slate-700/50" style="font-size: 10px;">=Inv (B2)</button>
                                <button type="button" @click="insertRef('standard', '=DeepSheet!B3')" class="px-1.5 py-0.5 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded transition-colors border border-slate-700/50" style="font-size: 10px;">=Date (B3)</button>
                                <button type="button" @click="insertRef('standard', '=DeepSheet!B4')" class="px-1.5 py-0.5 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded transition-colors border border-slate-700/50" style="font-size: 10px;">=Net (B4)</button>
                                <button type="button" @click="insertRef('standard', '=DeepSheet!B5')" class="px-1.5 py-0.5 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded transition-colors border border-slate-700/50" style="font-size: 10px;">=Gross (B5)</button>
                            </div>
                            <div class="mt-1 flex flex-wrap gap-1 items-center">
                                <span class="text-slate-500 font-medium" style="font-size: 10px;">Text:</span>
                                <button type="button" @click="insertRef('standard', 'DAF')" class="px-1.5 py-0.5 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded transition-colors border border-slate-700/50" style="font-size: 10px;">DAF:</button>
                                <button type="button" @click="insertRef('standard', 'BAVET')" class="px-1.5 py-0.5 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded transition-colors border border-slate-700/50" style="font-size: 10px;">BAVET</button>
                            </div>
                        </div>

                        <div class="mb-3">
                            <label class="block text-slate-300 text-xs mb-1 font-medium">DAF Override <span class="text-amber-400 font-medium" style="font-size: 10px;">(takes priority in DAF mode)</span></label>
                            <div class="flex gap-1.5">
                                <input type="text" v-model="editDafValue" class="flex-1 bg-slate-950 border border-slate-700 rounded-lg px-3 py-1.5 text-red-400 font-medium focus:outline-none focus:border-red-500 focus:ring-1 focus:ring-red-500 transition-all text-xs" placeholder="Leave empty to use base..." @keyup.enter="saveCellOverrides" />
                                <button type="button" @click="startCellSelection('daf')" class="p-1.5 bg-slate-800 hover:bg-slate-700 border border-slate-700 rounded-lg text-blue-400 hover:text-blue-300 transition-colors flex-shrink-0" title="Pick cell from sheet">
                                    <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" class="lucide lucide-crosshair"><circle cx="12" cy="12" r="10"/><line x1="22" y1="12" x2="18" y2="12"/><line x1="6" y1="12" x2="2" y2="12"/><line x1="12" y1="6" x2="12" y2="2"/><line x1="12" y1="22" x2="12" y2="18"/></svg>
                                </button>
                            </div>
                            <!-- Chips -->
                            <div class="mt-1.5 flex flex-wrap gap-1 items-center">
                                <span class="text-slate-500 font-medium" style="font-size: 10px;">Insert:</span>
                                <button type="button" @click="insertRef('daf', '=DeepSheet!B1')" class="px-1.5 py-0.5 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded transition-colors border border-slate-700/50" style="font-size: 10px;">=Ref (B1)</button>
                                <button type="button" @click="insertRef('daf', '=DeepSheet!B2')" class="px-1.5 py-0.5 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded transition-colors border border-slate-700/50" style="font-size: 10px;">=Inv (B2)</button>
                                <button type="button" @click="insertRef('daf', '=DeepSheet!B3')" class="px-1.5 py-0.5 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded transition-colors border border-slate-700/50" style="font-size: 10px;">=Date (B3)</button>
                                <button type="button" @click="insertRef('daf', '=DeepSheet!B4')" class="px-1.5 py-0.5 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded transition-colors border border-slate-700/50" style="font-size: 10px;">=Net (B4)</button>
                                <button type="button" @click="insertRef('daf', '=DeepSheet!B5')" class="px-1.5 py-0.5 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded transition-colors border border-slate-700/50" style="font-size: 10px;">=Gross (B5)</button>
                            </div>
                            <div class="mt-1 flex flex-wrap gap-1 items-center">
                                <span class="text-slate-500 font-medium" style="font-size: 10px;">Text:</span>
                                <button type="button" @click="insertRef('daf', 'DAF')" class="px-1.5 py-0.5 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded transition-colors border border-slate-700/50" style="font-size: 10px;">DAF:</button>
                                <button type="button" @click="insertRef('daf', 'BAVET')" class="px-1.5 py-0.5 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded transition-colors border border-slate-700/50" style="font-size: 10px;">BAVET</button>
                            </div>
                        </div>

                        <div class="flex gap-2 justify-end mt-4">
                            <button class="px-3 py-1.5 bg-slate-700 hover:bg-slate-600 text-white rounded-lg transition-colors text-xs font-medium" @click="closeEditor">Cancel</button>
                            <button class="px-3 py-1.5 bg-blue-600 hover:bg-blue-500 text-white rounded-lg transition-colors shadow-lg shadow-blue-500/20 text-xs font-medium" @click="saveCellOverrides" :disabled="isSavingCell">
                                {{ isSavingCell ? 'Saving...' : 'Save Overrides' }}
                            </button>
                        </div>
                        
                        <div v-if="editorMessage" class="mt-2 text-xs text-center font-medium" :class="editorMessageType === 'error' ? 'text-red-400' : 'text-emerald-400'">
                            {{ editorMessage }}
                        </div>
                    </div>
                    
                    <div v-else class="py-8 flex items-center justify-center text-slate-500 italic text-center p-8">
                        Click any cell in the sheet grid to edit overrides.
                    </div>

                    <div class="mt-4 p-2 bg-amber-500/10 border border-amber-500/30 rounded-lg text-[10px] text-amber-400 leading-snug">
                        ⚠ If footer overrides appear shifted after re-generating, the template structure likely changed. Re-apply overrides after verifying cell positions.
                    </div>
                </div>
            </div>
        </div>
    `,
    setup() {
        const store = useTemplateInspectorStore();
        const zoomLevel = ref(0.6);
        const showFullText = ref(false);
        const showDummyData = ref(false);

        // Cell override editor state
        const editingCell = ref(null);
        const editStandardValue = ref("");
        const editDafValue = ref("");
        const isSavingCell = ref(false);
        const editorMessage = ref("");
        const editorMessageType = ref("success");

        // Cell selection state
        const selectionTarget = ref(null); // 'standard' | 'daf' | null
        const editingCellClosedTemporarily = ref(null);

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

        const currentSheetDataRaw = computed(() => {
            if (!store.currentSheetName || !store.templateLayout) return null;
            return store.templateLayout[store.currentSheetName];
        });

        const currentSheetData = computed(() => {
            return currentSheetDataRaw.value;
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

            let headerMaxRow = 0;
            let maxCol = 0;

            const headerRows = sheet.header_rows || [];
            headerRows.forEach(rowDict => {
                const absRow = rowDict.relative_index ?? 0;
                if (absRow > headerMaxRow) headerMaxRow = absRow;
                for (const cellDict of (rowDict.cells || [])) {
                    const ci = (cellDict.col_index || 1) - 1;
                    if (ci > maxCol) maxCol = ci;

                    if (cellDict.merge) {
                        const mc = (cellDict.merge.max_col || 1) - 1;
                        if (mc > maxCol) maxCol = mc;
                    }
                }
            });

            const numDummyRows = showDummyData.value ? 3 : 0;
            const footerBaseRow = headerMaxRow + 1 + numDummyRows;
            let maxRow = headerMaxRow;

            const footerRows = sheet.template_footer_rows || sheet.footer_rows || [];
            footerRows.forEach(rowDict => {
                const absRow = footerBaseRow + (rowDict.relative_index ?? 0);
                if (absRow > maxRow) maxRow = absRow;
                for (const cellDict of (rowDict.cells || [])) {
                    const ci = (cellDict.col_index || 1) - 1;
                    if (ci > maxCol) maxCol = ci;

                    if (cellDict.merge) {
                        const mc = (cellDict.merge.max_col || 1) - 1;
                        if (mc > maxCol) maxCol = mc;
                    }
                }
                for (const m of (rowDict.merges || [])) {
                    const mc = (m.max_col || 1) - 1;
                    if (mc > maxCol) maxCol = mc;
                }
            });

            return { maxRow: maxRow, maxCol: maxCol, headerMaxRow, footerBaseRow };
        });

        const gridCells = computed(() => {
            const sheet = currentSheetData.value;
            if (!sheet) return [];

            const stylePalette = sheet.style_palette || {};
            const { maxRow, maxCol, footerBaseRow } = sheetBounds.value;

            const cells = [];
            const occupied = new Set();
            const mergedRanges = {};

            const gridCellsMap = {};

            // --- 1. PROCESS HEADER ---
            const headerRows = sheet.header_rows || [];
            headerRows.forEach(rowDict => {
                const absRow = rowDict.relative_index ?? 0;

                (rowDict.cells || []).forEach(cellDict => {
                    const colIdx = cellDict.col_index - 1;
                    gridCellsMap[`${absRow},${colIdx}`] = cellDict;

                    if (cellDict.merge) {
                        const minCol = cellDict.merge.min_col - 1;
                        const maxColM = cellDict.merge.max_col - 1;
                        const rowSpan = cellDict.merge.row_span || 1;
                        mergedRanges[`${absRow},${minCol}`] = { rowspan: rowSpan, colspan: maxColM - minCol + 1 };

                        for (let r = absRow; r < absRow + rowSpan; r++) {
                            for (let c = minCol; c <= maxColM; c++) {
                                if (r !== absRow || c !== minCol) {
                                    occupied.add(`${r},${c}`);
                                }
                            }
                        }
                    }
                });
            });

            // --- 2. PROCESS FOOTER ---
            const footerRows = sheet.template_footer_rows || sheet.footer_rows || [];
            footerRows.forEach(rowDict => {
                const relIdx = rowDict.relative_index ?? 0;
                const absRow = footerBaseRow + relIdx;

                (rowDict.cells || []).forEach(cellDict => {
                    const colIdx = cellDict.col_index - 1;
                    gridCellsMap[`${absRow},${colIdx}`] = cellDict;

                    if (cellDict.merge) {
                        const minCol = cellDict.merge.min_col - 1;
                        const maxColM = cellDict.merge.max_col - 1;
                        const rowSpan = cellDict.merge.row_span || 1;
                        mergedRanges[`${absRow},${minCol}`] = { rowspan: rowSpan, colspan: maxColM - minCol + 1 };

                        for (let r = absRow; r < absRow + rowSpan; r++) {
                            for (let c = minCol; c <= maxColM; c++) {
                                if (r !== absRow || c !== minCol) {
                                    occupied.add(`${r},${c}`);
                                }
                            }
                        }
                    }
                });

                (rowDict.merges || []).forEach(mDict => {
                    const minCol = mDict.min_col - 1;
                    const maxColM = mDict.max_col - 1;
                    const rowSpan = mDict.row_span || 1;
                    mergedRanges[`${absRow},${minCol}`] = { rowspan: rowSpan, colspan: maxColM - minCol + 1 };

                    for (let r = absRow; r < absRow + rowSpan; r++) {
                        for (let c = minCol; c <= maxColM; c++) {
                            if (r !== absRow || c !== minCol) {
                                occupied.add(`${r},${c}`);
                            }
                        }
                    }
                });
            });

            // --- 3. CLASSIFY DUMMY COLUMNS ---
            const dummyColData = {};
            if (showDummyData.value) {
                for (let c = 0; c <= maxCol; c++) {
                    let valLower = "";
                    for (const row of headerRows) {
                        const cellDict = (row.cells || []).find(cell => (cell.col_index - 1) === c);
                        if (cellDict) {
                            const rawVal = cellDict.value;
                            const cellVal = typeof rawVal === 'object' && rawVal !== null
                                ? (rawVal.default !== undefined ? rawVal.default : '')
                                : (rawVal || '');
                            valLower = String(cellVal).trim().toLowerCase();
                            if (valLower) break;
                        }
                    }
                    if (valLower) {
                        if (valLower.includes('desc') || valLower.includes('item') || valLower.includes('particular')) {
                            dummyColData[c] = 'desc';
                        } else if (valLower.includes('qty') || valLower.includes('pcs') || valLower.includes('quantity')) {
                            dummyColData[c] = 'qty';
                        } else if (valLower.includes('price') || valLower.includes('rate') || valLower.includes('unit')) {
                            dummyColData[c] = 'price';
                        } else if (valLower.includes('amount') || valLower.includes('total') || valLower.includes('value')) {
                            dummyColData[c] = 'amount';
                        } else if (valLower.includes('po ') || valLower.includes('po_') || valLower === 'po' || valLower.includes('order')) {
                            dummyColData[c] = 'po';
                        } else if (valLower.includes('cbm')) {
                            dummyColData[c] = 'cbm';
                        } else if (valLower.includes('ctn') || valLower.includes('box') || valLower.includes('carton') || valLower.includes('pkg')) {
                            dummyColData[c] = 'ctn';
                        } else if (valLower.includes('net')) {
                            dummyColData[c] = 'net';
                        } else if (valLower.includes('gross')) {
                            dummyColData[c] = 'gross';
                        } else if (valLower === 'no' || valLower === 'no.' || valLower === 'seq' || valLower === 'index') {
                            dummyColData[c] = 'no';
                        }
                    }
                }
            }

            // --- 4. BUILD GRID CELLS ARRAY ---
            for (let r = 0; r <= maxRow; r++) {
                if (showDummyData.value && r > headerMaxRow && r < footerBaseRow) {
                    const dummyRowIdx = r - headerMaxRow;
                    for (let c = 0; c <= maxCol; c++) {
                        const address = `${colToLetter(c)}${r + 1}`;
                        const dummyType = dummyColData[c];
                        let dummyContent = "";
                        if (dummyType === 'desc') dummyContent = `Dummy Item ${String.fromCharCode(64 + dummyRowIdx)}`;
                        else if (dummyType === 'qty') dummyContent = String(10 * dummyRowIdx);
                        else if (dummyType === 'price') dummyContent = "15.00";
                        else if (dummyType === 'amount') dummyContent = (10 * dummyRowIdx * 15).toFixed(2);
                        else if (dummyType === 'po') dummyContent = "PO-998877";
                        else if (dummyType === 'cbm') dummyContent = "0.25";
                        else if (dummyType === 'ctn') dummyContent = "5";
                        else if (dummyType === 'net') dummyContent = "100.0";
                        else if (dummyType === 'gross') dummyContent = "110.0";
                        else if (dummyType === 'no') dummyContent = String(dummyRowIdx);

                        cells.push({
                            id: address,
                            address: address,
                            content: dummyContent,
                            rawContent: dummyContent,
                            hasOverride: false,
                            currentOverrides: null,
                            title: `[Mock Data]`,
                            style: {
                                gridColumnStart: c + 1,
                                gridColumnEnd: c + 2,
                                gridRowStart: r + 1,
                                gridRowEnd: r + 2,
                                display: 'flex',
                                justifyContent: (dummyType === 'qty' || dummyType === 'price' || dummyType === 'amount' || dummyType === 'no') ? 'flex-end' : 'flex-start',
                                alignItems: 'center',
                                borderTop: '1px solid #cbd5e1',
                                borderRight: '1px solid #cbd5e1',
                                borderBottom: '1px solid #cbd5e1',
                                borderLeft: '1px solid #cbd5e1',
                                padding: '1px 2px',
                                fontSize: '11pt',
                                fontFamily: 'Arial, sans-serif',
                                backgroundColor: '#f8fafc',
                                color: '#64748b',
                                fontStyle: 'italic',
                                lineHeight: '1.2',
                                position: 'relative',
                                cursor: 'default'
                            },
                            isFormula: false,
                            isDummy: true
                        });
                    }
                    continue;
                }

                for (let c = 0; c <= maxCol; c++) {
                    if (occupied.has(`${r},${c}`)) continue;

                    const address = `${colToLetter(c)}${r + 1}`;
                    let cellContent = "";
                    let cellStyle = {};
                    const mergeInfo = mergedRanges[`${r},${c}`];

                    const cellDict = gridCellsMap[`${r},${c}`];
                    if (cellDict) {
                        cellContent = cellDict.value !== undefined && cellDict.value !== null ? cellDict.value : "";
                        if (cellDict.style) {
                            cellStyle = cellDict.style;
                        } else if (cellDict.style_id) {
                            cellStyle = stylePalette[cellDict.style_id] || {};
                        }
                    }

                    const isObj = typeof cellContent === 'object' && cellContent !== null;
                    const defaultVal = isObj ? (cellContent.default !== undefined && cellContent.default !== null ? cellContent.default : "") : (cellContent || "");

                    let cellTitle = `[${address}] ${defaultVal}`;
                    let overrideDisplay = "";
                    let hasStandardOverride = false;
                    let hasDafOverride = false;
                    let standardOverride = "";
                    let dafOverride = "";

                    if (isObj) {
                        const standard = cellContent.standard;
                        const daf = cellContent.daf;
                        const ov = [];
                        if (standard) {
                            ov.push(`Base: ${standard}`);
                            standardOverride = standard;
                            hasStandardOverride = true;
                        }
                        if (daf) {
                            ov.push(`DAF: ${daf}`);
                            if (daf !== standard) {
                                dafOverride = daf;
                                hasDafOverride = true;
                            }
                        }
                        if (ov.length > 0) {
                            cellTitle += `\nOverrides:\n• ${ov.join('\n• ')}`;
                        }

                        if (standard && daf && standard !== daf) {
                            overrideDisplay = `${standard} [DAF: ${daf}]`;
                        } else {
                            overrideDisplay = standard || daf || "";
                        }
                    }

                    cells.push({
                        id: address,
                        address: address,
                        content: defaultVal,
                        rawContent: cellContent,
                        hasOverride: isObj,
                        currentOverrides: isObj ? cellContent : null,
                        overrideDisplay: overrideDisplay,
                        hasStandardOverride,
                        hasDafOverride,
                        standardOverride,
                        dafOverride,
                        title: cellTitle,
                        style: { ...buildCellCss(r, c, cellStyle, mergeInfo, !cellContent), position: 'relative', cursor: 'pointer' },
                        isFormula: (typeof cellContent === 'string' && cellContent.startsWith('=')) || (isObj && overrideDisplay.startsWith('='))
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
                zoom: zoomLevel.value,
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

            const headerHeightLookup = {};
            const headerRows = sheet.header_rows || [];
            headerRows.forEach(rowDict => {
                if (rowDict.height != null) {
                    headerHeightLookup[rowDict.relative_index] = rowDict.height;
                }
            });

            const footerHeightLookup = {};
            footerRows.forEach(rowDict => {
                if (rowDict.height != null) {
                    footerHeightLookup[footerBaseRow + (rowDict.relative_index ?? 0)] = rowDict.height;
                }
            });
            const rows = [];
            for (let r = 0; r <= maxRow; r++) {
                const hdrH = rowHeightsMap[String(r + 1)] || headerHeightLookup[r];
                const ftrH = footerHeightLookup[r];
                const h = hdrH || ftrH;
                rows.push(h ? Math.max(Math.round(h * 1.333), 14) + 'px' : '20px');
            }
            base.gridTemplateRows = rows.join(' ');

            return base;
        });

        // --- Cell override editor triggers ---

        const openCellEditor = (cell) => {
            if (cell.isDummy) return;
            editingCell.value = cell;
            editStandardValue.value = (cell.currentOverrides && cell.currentOverrides.standard) || "";
            editDafValue.value = (cell.currentOverrides && cell.currentOverrides.daf) || "";
            editorMessage.value = "";
        };

        const startCellSelection = (target) => {
            selectionTarget.value = target;
            editingCellClosedTemporarily.value = editingCell.value;
            editingCell.value = null; // Close standard modal temporarily
        };

        const cancelCellSelection = () => {
            editingCell.value = editingCellClosedTemporarily.value;
            editingCellClosedTemporarily.value = null;
            selectionTarget.value = null;
        };

        const insertRef = (target, value) => {
            if (target === 'standard') {
                editStandardValue.value = value;
            } else if (target === 'daf') {
                editDafValue.value = value;
            }
        };

        const handleCellClick = (cell) => {
            if (cell.isDummy) return;
            if (selectionTarget.value) {
                // If there's a selection target, insert cell reference format
                // e.g. =SheetName!B3
                const refStr = `=${store.currentSheetName}!${cell.address}`;
                if (selectionTarget.value === 'standard') {
                    editStandardValue.value = refStr;
                } else if (selectionTarget.value === 'daf') {
                    editDafValue.value = refStr;
                }
                // Reopen the cell editor
                editingCell.value = editingCellClosedTemporarily.value;
                editingCellClosedTemporarily.value = null;
                selectionTarget.value = null;
            } else {
                openCellEditor(cell);
            }
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
            showDummyData,
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
            saveCellOverrides,
            selectionTarget,
            startCellSelection,
            cancelCellSelection,
            insertRef,
            handleCellClick
        };
    }
};
