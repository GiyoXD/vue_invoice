
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
                            <div class="h-date">{{ t.name }}</div>
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
                        <div class="status-box info" style="margin-bottom: 1rem;">
                            <strong>Viewing:</strong> {{ currentTemplateName }} <br>
                            <span style="font-size: 0.85em; opacity: 0.8;">Source: {{ currentTemplateFingerprint?.source_file }}</span>
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
                        
                        <!-- Excel Grid -->
                        <div class="excel-grid-container" style="overflow: auto; border: 1px solid #e2e8f0; max-height: 600px;">
                            <div class="excel-grid" :style="gridStyle">
                                <!-- Render Cells -->
                                <div v-for="cell in gridCells" :key="cell.id"
                                     class="excel-cell"
                                     :style="cell.style"
                                     :title="cell.address">
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
            try {
                const res = await fetch(`/api/template/view?name=${encodeURIComponent(t.name)}`);
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

        const currentTemplateName = computed(() => selectedTemplateName.value);
        const currentTemplateFingerprint = computed(() => currentTemplate.value?.fingerprint);
        const templateLayout = computed(() => currentTemplate.value?.template_layout || {});

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

        const gridCells = computed(() => {
            if (!currentSheetData.value) return [];

            const sheet = currentSheetData.value;
            const content = sheet.header_content || {};
            const styles = sheet.header_styles || {};
            const merges = sheet.header_merges || [];

            // Determine grid bounds
            let maxRow = 0;
            let maxCol = 0;

            // Check content bounds
            Object.keys(content).forEach(addr => {
                const { row, col } = parseAddress(addr);
                if (row > maxRow) maxRow = row;
                if (col > maxCol) maxCol = col;
            });

            // Check styles bounds (styles might exist for empty cells)
            Object.keys(styles).forEach(addr => {
                const { row, col } = parseAddress(addr);
                if (row > maxRow) maxRow = row;
                if (col > maxCol) maxCol = col;
            });

            // Also check merges bounds
            merges.forEach(startEnd => {
                const [start, end] = startEnd.split(":");
                const s = parseAddress(start);
                const e = parseAddress(end);
                if (e.row > maxRow) maxRow = e.row;
                if (e.col > maxCol) maxCol = e.col;
            });

            // Add some padding
            maxRow += 2;
            maxCol += 2;

            const cells = [];
            const occupied = new Set(); // track merged cells

            // Process Merges first to mark occupied
            const mergedRanges = {};
            merges.forEach(range => {
                const [start, end] = range.split(":");
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

            // Iterate 1..maxRow, 1..maxCol
            for (let r = 0; r <= maxRow; r++) {
                for (let c = 0; c <= maxCol; c++) {
                    if (occupied.has(`${r},${c}`)) continue;

                    // Reconstruct Address (simple implementation for A..Z, AA..AZ not handled perfectly but okay for now)
                    // Actually let's use a standard col converter
                    let colLetter = "";
                    let tempCol = c + 1;
                    while (tempCol > 0) {
                        let rem = (tempCol - 1) % 26;
                        colLetter = String.fromCharCode(65 + rem) + colLetter;
                        tempCol = Math.floor((tempCol - 1) / 26);
                    }
                    const address = `${colLetter}${r + 1}`;

                    const cellContent = content[address] || "";
                    const cellStyle = styles[address] || {};
                    const mergeInfo = mergedRanges[address];

                    // Build Style Object
                    const cssStyle = {
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
                        whiteSpace: cellStyle.alignment?.wrap_text ? 'normal' : 'nowrap',
                        overflow: 'hidden',
                        color: '#000'
                    };

                    cells.push({
                        id: address,
                        address: address,
                        content: cellContent,
                        style: cssStyle,
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
                backgroundColor: '#f1f5f9'
            };
        });

        // Helper Time
        const formatTime = (ts) => {
            if (!ts) return '';
            return new Date(ts).toLocaleString();
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
            gridCells,
            gridStyle,
            fetchTemplates,
            loadTemplate,
            formatTime
        };
    }
};
