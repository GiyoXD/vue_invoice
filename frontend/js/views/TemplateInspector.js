
import { ref, computed, onMounted, watch } from 'vue';

export default {
    template: `
        <div class="inspector-view fade-in">
            <h1>Template Inspector</h1>
            
            <div class="inspector-layout">
                <!-- Sidebar: Template List -->
                <div class="history-sidebar card">
                    <h3>Available Templates</h3>
                    <div v-if="templates.length === 0" style="color: #94a3b8; font-size: 0.875rem;">No templates found.</div>
                    <div class="history-list">
                        <div v-for="t in templates" :key="t.name" 
                             class="history-item" :class="{ active: selectedTemplateName === t.name }"
                             @click="loadTemplate(t)">
                            <div class="h-date">
                                {{ t.name }}
                                <span v-if="t.bundle_name && t.name !== t.bundle_name" style="font-size: 0.8em; color: #64748b; font-weight: normal;">({{ t.bundle_name }})</span>
                            </div>
                            <div class="h-file" style="font-size: 0.75rem; color: #64748b;">Source: {{ t.source_file }}</div>
                            <div class="h-stats">Updated: {{ formatTime(t.modified) }}</div>
                        </div>
                    </div>
                     <button class="btn-small" @click="fetchTemplates" style="width: 100%; margin-top: 1rem;">Refresh List</button>
                </div>

                <!-- Main: Details -->
                <div class="inspector-main card">
                    <div v-if="!currentTemplate" style="text-align: center; padding: 2rem; color: #64748b;">
                        <p>Select a template from the list to inspect.</p>
                    </div>

                    <div v-if="currentTemplate" class="template-viewer">
                        <div class="status-box info" style="margin-bottom: 1rem; display: flex; justify-content: space-between; align-items: flex-start;">
                            <div>
                                <strong>Viewing:</strong> {{ currentTemplateName }} <br>
                                <span style="font-size: 0.85em; opacity: 0.8;">Source: {{ currentTemplateFingerprint?.source_file }}</span>
                            </div>
                            <button class="btn-danger" @click="deleteTemplate" title="Delete Template" style="padding: 0.25rem 0.5rem; font-size: 0.875rem;">
                                Delete Template
                            </button>
                        </div>

                        <!-- Sheet Selector -->
                        <div class="sheet-tabs" style="margin-bottom: 1rem; display: flex; gap: 0.5rem;">
                            <button v-for="(sheetData, sheetName) in templateLayout" :key="sheetName"
                                    class="btn-small" 
                                    :class="{ 'btn-primary': currentSheetName === sheetName }"
                                    @click="currentSheetName = sheetName">
                                {{ sheetName }}
                            </button>
                        </div>
                        
                        <!-- Zoom & View Controls -->
                        <div class="zoom-controls" style="margin-bottom: 0.5rem; display: flex; gap: 0.5rem; align-items: center;">
                            <button class="btn-small" @click="zoomOut" title="Zoom Out">-</button>
                            <span style="font-size: 0.875rem; min-width: 3rem; text-align: center;">{{ zoomPercentage }}%</span>
                            <button class="btn-small" @click="zoomIn" title="Zoom In">+</button>
                            <button class="btn-small" @click="resetZoom" title="Reset Zoom">Reset</button>

                            <div style="width: 1px; height: 1.5rem; background: #cbd5e1; margin: 0 0.5rem;"></div>

                            <label style="display: flex; align-items: center; gap: 0.25rem; font-size: 0.875rem; cursor: pointer; user-select: none;">
                                <input type="checkbox" v-model="showFullText"> Wrap Text
                            </label>
                        </div>
                        
                        <!-- Excel Grid -->
                        <div class="excel-grid-container" style="overflow: auto; border: 1px solid #e2e8f0; max-height: 600px;">
                            <div class="excel-grid" :style="gridStyle">
                                <!-- Render Cells -->
                                <div v-for="cell in gridCells" :key="cell.id"
                                     class="excel-cell"
                                     :style="cell.style"
                                     :title="'[' + cell.address + '] ' + cell.content">
                                     <span v-if="cell.isFormula" style="color: blue; font-style: italic;">{{ cell.content }}</span>
                                     <span v-else>{{ cell.content }}</span>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    `,
    setup() {
        const templates = ref([]);
        const selectedTemplateName = ref(null);
        const currentTemplate = ref(null);
        const currentSheetName = ref(null);
        const zoomLevel = ref(1.0);
        const showFullText = ref(false);

        const zoomIn = () => {
            zoomLevel.value = Math.min(zoomLevel.value + 0.1, 3.0);
        };
        const zoomOut = () => {
            zoomLevel.value = Math.max(zoomLevel.value - 0.1, 0.2);
        };
        const resetZoom = () => {
            zoomLevel.value = 1.0;
        };

        // Fetch list
        const fetchTemplates = async () => {
            try {
                const res = await fetch('/api/templates');
                if (res.ok) {
                    templates.value = await res.json();
                }
            } catch (e) {
                console.error("Failed to fetch templates", e);
            }
        };

        const loadTemplate = async (t) => {
            selectedTemplateName.value = t.name;
            currentTemplate.value = null; // Clear immediately to prevent showing old data while loading
            try {
                const url = t.bundle_name 
                    ? `/api/template/view?name=${encodeURIComponent(t.name)}&bundle=${encodeURIComponent(t.bundle_name)}&_t=${Date.now()}`
                    : `/api/template/view?name=${encodeURIComponent(t.name)}&_t=${Date.now()}`;
                    
                const res = await fetch(url);
                if (res.ok) {
                    currentTemplate.value = await res.json();
                    // Default to first sheet
                    const sheets = Object.keys(currentTemplate.value?.template_layout || {});
                    if (sheets.length > 0) currentSheetName.value = sheets[0];
                }
            } catch (e) {
                console.error("Failed to load template", e);
            }
        };

        const deleteTemplate = async () => {
            if (!selectedTemplateName.value) return;
            
            // Find the full template object to get the bundle_name
            const t = templates.value.find(tmpl => tmpl.name === selectedTemplateName.value);
            const bundleName = t?.bundle_name || selectedTemplateName.value;
            
            if (!confirm(`WARNING: Are you sure you want to permanently delete the ENTIRE template bundle for '${bundleName}'?\n\nThis will delete all variants (Base, KH, VN, etc) and configuration files within the bundle folder.`)) {
                return;
            }
            try {
                const url = t?.bundle_name
                    ? `/api/template/${encodeURIComponent(selectedTemplateName.value)}?bundle=${encodeURIComponent(t.bundle_name)}`
                    : `/api/template/${encodeURIComponent(selectedTemplateName.value)}`;
                    
                const res = await fetch(url, {
                    method: 'DELETE'
                });
                if (res.ok) {
                    currentTemplate.value = null;
                    selectedTemplateName.value = null;
                    currentSheetName.value = null;
                    await fetchTemplates();
                    alert(`Template bundle deleted successfully.`);
                } else {
                    const data = await res.json();
                    alert(`Failed to delete template: ${data.error || res.statusText}`);
                }
            } catch (e) {
                console.error("Error deleting template", e);
                alert('An error occurred while deleting the template.');
            }
        };

        const currentTemplateName = computed(() => selectedTemplateName.value);
        const currentTemplateFingerprint = computed(() => currentTemplate.value?.fingerprint);
        const templateLayout = computed(() => currentTemplate.value?.template_layout || {});

        const zoomPercentage = computed(() => Math.round(zoomLevel.value * 100));

        const currentSheetData = computed(() => {
            if (!currentSheetName.value || !templateLayout.value) return null;
            return templateLayout.value[currentSheetName.value];
        });

        // --- Grid Generation Logic ---

        // Helper to convert A1 to {row, col} (0-indexed)
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

        /**
         * Converts a 0-indexed column number to a column letter (e.g. 0 -> A, 25 -> Z, 26 -> AA).
         */
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

        /**
         * Flattens the grouped style map into a per-cell lookup.
         * Supports both new grouped format {hashId: [coords]} and legacy per-cell format {coord: styleDict}.
         */
        const flattenStyles = (stylesRaw, stylePalette) => {
            const result = {};
            for (const [key, value] of Object.entries(stylesRaw)) {
                if (Array.isArray(value)) {
                    // New grouped format: key = style_id, value = ["A1", "B2", ...]
                    const resolved = stylePalette[key] || {};
                    for (const coord of value) {
                        result[coord] = resolved;
                    }
                } else if (typeof value === 'object' && value !== null) {
                    // Legacy per-cell format: key = coord, value = style dict
                    result[key] = value;
                } else if (typeof value === 'string') {
                    // Legacy per-cell with palette ref: key = coord, value = style_id string
                    result[key] = stylePalette[value] || {};
                }
            }
            return result;
        };

        /**
         * Builds a CSS style object for a grid cell given its position, style dict, and merge info.
         */
        const buildCellCss = (r, c, cellStyle, mergeInfo) => {
            return {
                gridColumnStart: c + 1,
                gridColumnEnd: mergeInfo ? c + 1 + mergeInfo.colspan : c + 2,
                gridRowStart: r + 1,
                gridRowEnd: mergeInfo ? r + 1 + mergeInfo.rowspan : r + 2,
                border: '1px solid #cbd5e1',
                padding: '4px',
                fontSize: (cellStyle.font?.size || 11) + 'pt',
                fontWeight: cellStyle.font?.bold ? 'bold' : 'normal',
                fontFamily: cellStyle.font?.name || 'Arial',
                textAlign: cellStyle.alignment?.horizontal || 'left',
                verticalAlign: cellStyle.alignment?.vertical || 'bottom',
                backgroundColor: '#fff',
                whiteSpace: (showFullText.value || cellStyle.alignment?.wrap_text) ? 'normal' : 'nowrap',
                overflow: showFullText.value ? 'visible' : 'hidden',
                wordBreak: showFullText.value ? 'break-word' : 'normal',
                color: '#000'
            };
        };

        const gridCells = computed(() => {
            if (!currentSheetData.value) return [];

            const sheet = currentSheetData.value;
            const content = sheet.header_content || {};
            const stylePalette = sheet.style_palette || {};
            const styles = flattenStyles(sheet.header_styles || {}, stylePalette);

            // Normalize merges: support both dict {"A1:B2": "val"} and array ["A1:B2"]
            const mergesRaw = sheet.header_merges || {};
            const merges = Array.isArray(mergesRaw) ? mergesRaw : Object.keys(mergesRaw);

            // --- Collect footer content into the same coordinate maps ---
            const footerContent = {};
            const footerStyles = {};
            const footerMergeRanges = []; // strings like "A10:C10"
            const footerRows = sheet.footer_rows || [];

            // We place footer rows right after the last header row
            // First determine header extent to set footer offset
            let headerMaxRow = 0;
            Object.keys(content).forEach(addr => {
                const { row } = parseAddress(addr);
                if (row > headerMaxRow) headerMaxRow = row;
            });
            Object.keys(styles).forEach(addr => {
                const { row } = parseAddress(addr);
                if (row > headerMaxRow) headerMaxRow = row;
            });
            merges.forEach(range => {
                const parts = range.split(":");
                if (parts.length === 2) {
                    const e = parseAddress(parts[1]);
                    if (e.row > headerMaxRow) headerMaxRow = e.row;
                }
            });

            const footerBaseRow = headerMaxRow + 1; // 0-indexed row for first footer row

            footerRows.forEach(rowDict => {
                const relIdx = rowDict.relative_index ?? 0;
                const absRow = footerBaseRow + relIdx; // 0-indexed

                // Process cells
                for (const cellDict of (rowDict.cells || [])) {
                    const colIdx = cellDict.col_index; // 1-based
                    const addr = `${colToLetter(colIdx - 1)}${absRow + 1}`;

                    if (cellDict.value != null) {
                        footerContent[addr] = String(cellDict.value);
                    }
                    if (cellDict.style_id) {
                        footerStyles[addr] = stylePalette[cellDict.style_id] || {};
                    }
                }

                // Process merges
                for (const mDict of (rowDict.merges || [])) {
                    const minCol = mDict.min_col; // 1-based
                    const maxCol = mDict.max_col;
                    const rowSpan = mDict.row_span || 1;
                    const startAddr = `${colToLetter(minCol - 1)}${absRow + 1}`;
                    const endAddr = `${colToLetter(maxCol - 1)}${absRow + rowSpan}`;
                    footerMergeRanges.push(`${startAddr}:${endAddr}`);
                }
            });

            // Merge header + footer into unified maps
            const allContent = { ...content, ...footerContent };
            const allStyles = { ...styles, ...footerStyles };
            const allMerges = [...merges, ...footerMergeRanges];

            // Determine grid bounds
            let maxRow = 0;
            let maxCol = 0;

            Object.keys(allContent).forEach(addr => {
                const { row, col } = parseAddress(addr);
                if (row > maxRow) maxRow = row;
                if (col > maxCol) maxCol = col;
            });

            Object.keys(allStyles).forEach(addr => {
                const { row, col } = parseAddress(addr);
                if (row > maxRow) maxRow = row;
                if (col > maxCol) maxCol = col;
            });

            allMerges.forEach(range => {
                const parts = range.split(":");
                if (parts.length === 2) {
                    const s = parseAddress(parts[0]);
                    const e = parseAddress(parts[1]);
                    if (e.row > maxRow) maxRow = e.row;
                    if (e.col > maxCol) maxCol = e.col;
                }
            });

            maxRow += 2;
            maxCol += 2;

            const cells = [];
            const occupied = new Set();

            // Process Merges first to mark occupied
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

            // Iterate grid
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
                        content: cellContent,
                        style: buildCellCss(r, c, cellStyle, mergeInfo),
                        isFormula: typeof cellContent === 'string' && cellContent.startsWith('=')
                    });
                }
            }
            return cells;
        });

        const gridStyle = computed(() => {
            // We could define dynamic row heights/col widths here if we parse them
            // For now, let's just use auto.
            return {
                display: 'grid',
                // Use repeat(auto-fill, minmax(...)) or just a large grid
                // Better to set specific row heights if available
                gap: '0',
                backgroundColor: '#f1f5f9',
                transform: `scale(${zoomLevel.value})`,
                transformOrigin: 'top left',
                width: 'fit-content' // Ensure grid doesn't stretch weirdly when zoomed out
            };
        });

        // Helper Time
        const formatTime = (ts) => {
            if (!ts) return '';
            const d = new Date(ts);
            const pad = (n) => String(n).padStart(2, '0');
            return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
        };

        onMounted(() => {
            fetchTemplates();
        });

        return {
            templates,
            selectedTemplateName,
            currentTemplate,
            currentTemplateName,
            currentTemplateFingerprint,
            templateLayout,
            currentSheetName,
            zoomLevel,
            zoomPercentage,
            showFullText,
            zoomIn,
            zoomOut,
            resetZoom,
            gridCells,
            gridStyle,
            fetchTemplates,
            loadTemplate,
            deleteTemplate,
            formatTime
        };
    }
};
