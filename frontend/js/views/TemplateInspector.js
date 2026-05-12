
import { ref, computed, onMounted, watch } from 'vue';

export default {
    template: `
        <div class="max-w-[1600px] mx-auto py-8 fade-in h-screen flex flex-col">
            <h1 class="text-4xl font-extrabold tracking-tight text-transparent bg-clip-text bg-gradient-to-r from-blue-400 to-emerald-400 drop-shadow-md flex-shrink-0">Template Inspector</h1>
            
            <div class="flex gap-6 mt-8 flex-grow min-h-0">
                <!-- Sidebar: Template List -->
                <div class="w-[300px] bg-slate-800/80 backdrop-blur-md border border-slate-700/50 shadow-2xl rounded-2xl p-6 flex flex-col flex-shrink-0">
                    <h3 class="text-xl font-bold text-slate-100 mb-4 mt-0">Available Templates</h3>
                    
                    <div class="mb-4">
                        <input type="text" v-model="searchQuery" placeholder="Search templates..." class="w-full bg-slate-900 border border-slate-700 rounded-lg px-4 py-2 text-slate-100 focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 transition-all" />
                        <button class="w-full px-4 py-2 bg-slate-700 hover:bg-slate-600 text-slate-200 rounded-lg shadow-sm transition-colors mt-4 flex-shrink-0" @click="fetchTemplates">Refresh List</button>
                    </div>

                    <div v-if="filteredTemplates.length === 0" class="text-slate-400 text-base">No templates found.</div>
                    <div class="flex-1 overflow-y-auto flex flex-col gap-2 pr-2">
                        <div v-for="t in filteredTemplates" :key="t.name + (t.bundle_name || '') + (t.source_file || '')" 
                             class="p-3 bg-slate-900/50 border border-slate-700 rounded-lg cursor-pointer transition-all hover:border-blue-500 hover:bg-slate-800" :class="{ 'border-blue-500 bg-slate-800': selectedTemplateName === t.name }"
                             @click="loadTemplate(t)">
                            <div class="text-emerald-400 font-bold text-sm mb-1 truncate drop-shadow-sm">
                                {{ t.name }}
                                <span v-if="t.bundle_name && t.name !== t.bundle_name" class="text-lg text-red-500 font-bold ml-1">({{ t.bundle_name }})</span>
                            </div>
                            <div class="text-slate-400 text-xs mb-2 truncate">Source: {{ t.source_file }}</div>
                            <div class="text-slate-500 text-xs">Updated: {{ formatTime(t.modified) }}</div>
                        </div>
                    </div>
                </div>

                <!-- Main: Details -->
                <div class="flex-1 bg-slate-800/80 backdrop-blur-md border border-slate-700/50 shadow-2xl rounded-2xl p-6 flex flex-col min-w-0">
                    <div v-if="!currentTemplate" class="text-center p-8 text-muted">
                        <p>Select a template from the list to inspect.</p>
                    </div>

                    <div v-if="currentTemplate" class="template-viewer">
                        <div class="flex items-center justify-between p-4 bg-blue-500/10 border border-blue-500/20 rounded-xl mb-4 flex-shrink-0">
                            <div>
                                <strong class="text-blue-400">Viewing:</strong> <span class="text-slate-200">{{ currentTemplateName }}</span> <br>
                                <span class="text-sm text-slate-400 opacity-80">Source: {{ currentTemplateFingerprint?.source_file }}</span>
                            </div>
                            <button class="px-4 py-2 bg-red-500 hover:bg-red-400 text-white font-medium rounded-lg shadow-sm transition-colors text-sm" @click="deleteTemplate" title="Delete Template">
                                Delete Template
                            </button>
                        </div>

                        <!-- Client Notes Section -->
                        <div class="mb-4 bg-slate-900/50 border border-slate-700 rounded-xl p-4">
                            <div class="flex justify-between items-center mb-2">
                                <h4 class="m-0 text-sm text-slate-400 flex items-center gap-2">
                                    <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" class="lucide lucide-notebook-pen"><path d="M11 2H9a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2v-4"/><path d="m16 2 4 4-8 8H8v-4l8-8Z"/><path d="M15 5 19 9"/></svg>
                                    Client Notes / Remarks
                                </h4>
                                <button v-if="!isEditingNotes" class="px-3 py-1 bg-slate-700 hover:bg-slate-600 text-slate-200 rounded text-xs transition-colors" @click="isEditingNotes = true">Edit</button>
                                <div v-else class="flex gap-1">
                                    <button class="px-3 py-1 bg-slate-600 hover:bg-slate-500 text-white rounded text-xs transition-colors" @click="cancelEditNotes">Cancel</button>
                                    <button class="px-3 py-1 bg-blue-600 hover:bg-blue-500 text-white rounded text-xs transition-colors" @click="saveNotes" :disabled="isSavingNotes">
                                        {{ isSavingNotes ? 'Saving...' : 'Save' }}
                                    </button>
                                </div>
                            </div>
                            <div v-if="!isEditingNotes">
                                <div v-if="templateNotes" class="text-sm text-primary whitespace-pre-wrap leading-relaxed">{{ templateNotes }}</div>
                                <div v-else class="text-sm text-muted italic">No notes for this client yet. Click Edit to add.</div>
                            </div>
                            <div v-else>
                                <textarea v-model="editingNotesText" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-4 py-2 text-slate-100 focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 transition-all text-sm min-h-[100px]" placeholder="Enter things to remember for this client..."></textarea>
                            </div>
                        </div>

                        <!-- Client Profile Section -->
                        <div v-if="clientProfile" class="mb-4 bg-slate-900/50 border border-slate-700 rounded-xl p-4">
                            <h4 class="mb-3 text-sm text-slate-400 flex items-center gap-2">
                                <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/></svg>
                                Client Profile
                            </h4>
                            <div class="text-sm grid grid-cols-120-1fr gap-x-4 gap-y-2">
                                <div class="text-muted font-medium">Company:</div>
                                <div class="text-primary">{{ clientProfile.fullname || '—' }}</div>

                                <div class="text-muted font-medium">Address:</div>
                                <div class="text-primary whitespace-pre-line">{{ clientProfile.address || '—' }}</div>

                                <div class="text-muted font-medium">Contact:</div>
                                <div class="text-primary whitespace-pre-line">{{ clientProfile.contact || '—' }}</div>

                                <div class="text-muted font-medium">Shipping:</div>
                                <div class="text-primary">{{ clientProfile.shipping || '—' }}</div>
                            </div>
                        </div>

                        <!-- Table Information Section -->
                        <div v-if="currentTemplate && currentTemplate.table_info" class="mb-4 bg-slate-900/50 border border-slate-700 rounded-xl p-4">
                            <h4 class="mb-2 text-sm text-slate-400 flex items-center gap-2">
                                <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" class="lucide lucide-table-properties"><path d="M15 2H9a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V4a2 2 0 0 0-2-2Z"/><path d="M9 10h12"/><path d="M9 14h12"/><path d="M9 18h12"/><path d="M9 6h12"/><path d="M11 2v20"/></svg>
                                Table Information
                            </h4>
                            <div class="text-sm grid grid-cols-auto-1fr gap-x-4 gap-y-2">
                                <div class="text-muted font-medium">Fallback Desc (Standard):</div>
                                <div class="text-primary">{{ currentTemplate.table_info.fallback_description?.standard || 'None' }}</div>

                                <div class="text-muted font-medium">Fallback Desc (DAF):</div>
                                <div class="text-primary">{{ currentTemplate.table_info.fallback_description?.daf || 'None' }}</div>

                                <div class="text-muted font-medium">HS Code:</div>
                                <div class="text-primary">{{ currentTemplate.table_info.hs_code || 'None' }}</div>
                            </div>
                        </div>

                        <!-- Sheet Selector -->
                        <div class="sheet-tabs mb-4 flex gap-2">
                            <button v-for="(sheetData, sheetName) in templateLayout" :key="sheetName"
                                    class="px-4 py-2 rounded-lg text-sm font-medium transition-colors border border-transparent" 
                                    :class="currentSheetName === sheetName ? 'bg-blue-600 text-white' : 'bg-slate-700 text-slate-300 hover:bg-slate-600'"
                                    @click="currentSheetName = sheetName">
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

                            <label class="flex items-center gap-1 text-sm select-none cursor-pointer">
                                <input type="checkbox" v-model="showFullText"> Wrap Text
                            </label>
                        </div>
                        
                        <!-- Excel Grid -->
                        <div class="excel-grid-container overflow-auto max-h-75vh relative">
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
                                <div class="bg-slate-900 border border-slate-700 rounded-2xl p-6 shadow-2xl min-w-[400px]">
                                    <h3 class="m-0 mb-3 text-base text-primary">Cell {{ editingCell.address }}</h3>

                                    <div class="mb-4 p-4 bg-slate-800 rounded-lg text-sm border border-slate-700/50">
                                        <span class="text-secondary">Current (default):</span>
                                        <span class="text-primary ml-2">
                                            {{ (typeof editingCell.rawContent === 'object' && editingCell.rawContent !== null) ? (editingCell.rawContent.default ?? "") : (editingCell.rawContent || '(empty)') }}
                                        </span>
                                    </div>

                                    <div class="mb-3">
                                        <label class="block text-secondary text-sm mb-1">Base Value <span class="text-blue-400 text-xs">(applies to ALL modes)</span></label>
                                        <input type="text" v-model="editStandardValue" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-4 py-2 text-slate-100 focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 transition-all" placeholder="Enter value for standard, custom, DAF..." @keyup.enter="saveCellOverrides" />
                                        <p class="text-muted text-xs mt-1 mb-0">This value will be used in Standard, Custom, DAF, and any other mode.</p>
                                    </div>

                                    <div class="mb-3">
                                        <label class="block text-secondary text-sm mb-1">DAF Override <span class="text-amber-400 text-xs">(takes priority in DAF mode)</span></label>
                                        <input type="text" v-model="editDafValue" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-4 py-2 text-slate-100 focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 transition-all" placeholder="Leave empty to use base value" @keyup.enter="saveCellOverrides" />
                                        <p class="text-muted text-xs mt-1 mb-0">Only used when generating in DAF mode. If empty, the base value is used.</p>
                                    </div>

                                    <div v-if="editingCell.currentOverrides" class="mb-3 px-3 py-2 bg-blue-100 border border-blue-200 rounded text-sm">
                                        <div class="text-blue-400 mb-1">Existing overrides:</div>
                                        <div v-for="(v, k) in editingCell.currentOverrides" :key="k" class="text-blue-300">
                                            <strong>{{ k }}:</strong> {{ v }}
                                        </div>
                                    </div>

                                    <div class="flex gap-3 justify-end mt-6">
                                        <button class="px-4 py-2 bg-slate-600 hover:bg-slate-500 text-white rounded-lg transition-colors" @click="closeEditor">Cancel</button>
                                        <button class="px-4 py-2 bg-blue-600 hover:bg-blue-500 text-white rounded-lg transition-colors shadow-lg shadow-blue-500/20" @click="saveCellOverrides" :disabled="isSavingCell">
                                            {{ isSavingCell ? 'Saving...' : 'Save Overrides' }}
                                        </button>
                                    </div>
                                    <div v-if="editorMessage" class="mt-2 text-sm" :class="editorMessageType === 'error' ? 'text-red-500' : 'text-emerald-500'">
                                        {{ editorMessage }}
                                    </div>
                                    <div class="mt-4 p-3 bg-amber-500/10 border border-amber-500/30 rounded-lg text-xs text-amber-400 leading-snug">
                                        ⚠ If footer overrides appear shifted after re-generating, the Excel template structure likely changed (rows added/removed). Re-apply overrides after verifying cell positions or delete the template to create a new one.
                                    </div>
                                </div>
                            </div>
                        </teleport>
                    </div>
                </div>
            </div>
        </div>
    `,
    setup() {
        const templates = ref([]);
        const searchQuery = ref("");
        const selectedTemplateName = ref(null);
        const currentTemplate = ref(null);
        const currentSheetName = ref(null);
        const zoomLevel = ref(0.6);
        const showFullText = ref(false);

        // Client Notes state
        const isEditingNotes = ref(false);
        const isSavingNotes = ref(false);
        const editingNotesText = ref("");

        const templateNotes = computed(() => currentTemplate.value?.notes || "");
        const clientProfile = computed(() => currentTemplate.value?.client_profile || null);

        const saveNotes = async () => {
            if (!selectedTemplateName.value) return;
            isSavingNotes.value = true;

            const t = templates.value.find(tmpl => tmpl.name === selectedTemplateName.value);
            try {
                const res = await fetch('/api/template/notes', {
                    method: 'PATCH',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        template_name: selectedTemplateName.value,
                        bundle_name: t?.bundle_name || "",
                        notes: editingNotesText.value
                    })
                });

                if (res.ok) {
                    if (currentTemplate.value) {
                        currentTemplate.value.notes = editingNotesText.value;
                    }
                    isEditingNotes.value = false;
                } else {
                    const data = await res.json();
                    alert(`Failed to save notes: ${data.error || 'Unknown error'}`);
                }
            } catch (e) {
                console.error("Error saving notes", e);
                alert("Failed to save notes. See console for details.");
            } finally {
                isSavingNotes.value = false;
            }
        };

        const cancelEditNotes = () => {
            isEditingNotes.value = false;
            editingNotesText.value = templateNotes.value;
        };

        watch(currentTemplate, (newVal) => {
            if (newVal) {
                editingNotesText.value = newVal.notes || "";
                isEditingNotes.value = false;
            }
        });

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

        const loadTemplate = async (t, preserveState = false) => {
            selectedTemplateName.value = t.name;
            if (!preserveState) {
                currentTemplate.value = null; // Clear immediately to prevent showing old data while loading
            }
            try {
                const url = t.bundle_name
                    ? `/api/template/view?name=${encodeURIComponent(t.name)}&bundle=${encodeURIComponent(t.bundle_name)}&_t=${Date.now()}`
                    : `/api/template/view?name=${encodeURIComponent(t.name)}&_t=${Date.now()}`;

                const res = await fetch(url);
                if (res.ok) {
                    currentTemplate.value = await res.json();
                    const sheets = Object.keys(currentTemplate.value?.template_layout || {});

                    // Default to first sheet, unless preserving state and current sheet is still valid
                    if (!preserveState || !currentSheetName.value || !sheets.includes(currentSheetName.value)) {
                        if (sheets.length > 0) currentSheetName.value = sheets[0];
                        else currentSheetName.value = null;
                    }
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

        const filteredTemplates = computed(() => {
            const q = (searchQuery.value || "").trim().toLowerCase();
            if (!q) return templates.value;

            console.log(`Filtering for: "${q}"`);
            const filtered = templates.value.filter(t => {
                const nameMatch = (t.name || "").toLowerCase().includes(q);
                const bundleMatch = (t.bundle_name || "").toLowerCase().includes(q);
                if (nameMatch || bundleMatch) {
                    console.log(`Match found: ${t.name} (Bundle: ${t.bundle_name})`);
                }
                return nameMatch || bundleMatch;
            });
            return filtered;
        });

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
         * Convert Excel ARGB color string ("FF4472C4" or "4472C4") to CSS hex.
         */
        const excelColorToCss = (colorStr) => {
            if (!colorStr || typeof colorStr !== 'string') return null;
            let hex = colorStr.replace(/^#/, '');
            if (hex.length === 8) hex = hex.substring(2); // Strip alpha
            if (hex.length === 6 && /^[0-9A-Fa-f]{6}$/.test(hex)) return `#${hex}`;
            return null;
        };

        /**
         * Map Excel border style name to CSS border string.
         */
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
         * Now renders borders, fill colors, and font colors from the style palette.
         */
        const buildCellCss = (r, c, cellStyle, mergeInfo, isEmpty) => {
            // --- Borders ---
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

            // --- Fill / Background Color ---
            // Empty cells without explicit fill get transparent bg so overflow text shows through
            let bgColor = (isEmpty && !cellStyle.fill?.color) ? 'transparent' : '#fff';
            if (cellStyle.fill?.color) {
                const parsed = excelColorToCss(cellStyle.fill.color);
                if (parsed) bgColor = parsed;
            }

            // --- Font Color ---
            let fontColor = '#000';
            if (cellStyle.font?.color) {
                const parsed = excelColorToCss(cellStyle.font.color);
                if (parsed) fontColor = parsed;
            }

            // --- Alignment (flex-based since .excel-cell is display:flex) ---
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

        /**
         * Shared computed: scans sheet data once to determine grid dimensions.
         * Both gridCells and gridStyle reference this to avoid duplicated bounds scanning.
         */
        const sheetBounds = computed(() => {
            const sheet = currentSheetData.value;
            if (!sheet) return { maxRow: 0, maxCol: 0, headerMaxRow: 0, footerBaseRow: 0 };

            const content = sheet.template_header_content || sheet.header_content || {};
            const stylePalette = sheet.style_palette || {};
            const styles = flattenStyles(sheet.template_header_styles || sheet.header_styles || {}, stylePalette);
            const mergesRaw = sheet.template_header_merges || sheet.header_merges || {};
            const merges = Array.isArray(mergesRaw) ? mergesRaw : Object.keys(mergesRaw);
            const footerRows = sheet.template_footer_rows || sheet.footer_rows || [];

            // Header extent
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

            // Footer extent
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

            // Normalize merges: support both dict {"A1:B2": "val"} and array ["A1:B2"]
            const mergesRaw = sheet.template_header_merges || sheet.header_merges || {};
            const merges = Array.isArray(mergesRaw) ? mergesRaw : Object.keys(mergesRaw);

            // --- Collect footer content into the same coordinate maps ---
            const footerContent = {};
            const footerStyles = {};
            const footerMergeRanges = []; // strings like "A10:C10"
            const footerRows = sheet.template_footer_rows || sheet.footer_rows || [];

            const { maxRow, maxCol, footerBaseRow } = sheetBounds.value;

            footerRows.forEach(rowDict => {
                const relIdx = rowDict.relative_index ?? 0;
                const absRow = footerBaseRow + relIdx; // 0-indexed

                // Process cells
                for (const cellDict of (rowDict.cells || [])) {
                    const colIdx = cellDict.col_index; // 1-based
                    const addr = `${colToLetter(colIdx - 1)}${absRow + 1}`;

                    if (cellDict.value !== undefined && cellDict.value !== null) {
                        footerContent[addr] = cellDict.value;
                    }
                    if (cellDict.style_id) {
                        footerStyles[addr] = stylePalette[cellDict.style_id] || {};
                    }
                }

                // Process merges
                for (const mDict of (rowDict.merges || [])) {
                    const minCol = mDict.min_col; // 1-based
                    const maxColM = mDict.max_col;
                    const rowSpan = mDict.row_span || 1;
                    const startAddr = `${colToLetter(minCol - 1)}${absRow + 1}`;
                    const endAddr = `${colToLetter(maxColM - 1)}${absRow + rowSpan}`;
                    footerMergeRanges.push(`${startAddr}:${endAddr}`);
                }
            });

            // Merge header + footer into unified maps
            const allContent = { ...content, ...footerContent };
            const allStyles = { ...styles, ...footerStyles };
            const allMerges = [...merges, ...footerMergeRanges];

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

            // --- Column widths (Excel char units → px: width * 7.5) ---
            const cols = [];
            for (let c = 0; c <= maxCol; c++) {
                const letter = colToLetter(c);
                const w = colWidthsMap[letter];
                cols.push(w ? Math.max(Math.round(w * 7.5), 20) + 'px' : '64px');
            }
            base.gridTemplateColumns = cols.join(' ');

            // --- Row heights (Excel points → px: pt * 1.333) ---
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

        // Helper Time
        const formatTime = (ts) => {
            if (!ts) return '';
            const d = new Date(ts);
            const pad = (n) => String(n).padStart(2, '0');
            return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
        };

        /**
         * Opens the cell override editor popup for the clicked cell.
         */
        const openCellEditor = (cell) => {
            editingCell.value = cell;
            // Pre-fill with existing overrides if present
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

        /**
         * Saves mode-specific overrides for the currently editing cell via PATCH /api/template/cell.
         */
        const saveCellOverrides = async () => {
            if (!editingCell.value || !currentSheetName.value) return;
            isSavingCell.value = true;
            editorMessage.value = "";

            const t = templates.value.find(tmpl => tmpl.name === selectedTemplateName.value);
            try {
                const res = await fetch('/api/template/cell', {
                    method: 'PATCH',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        template_name: selectedTemplateName.value,
                        bundle_name: t?.bundle_name || "",
                        sheet_name: currentSheetName.value,
                        cell_address: editingCell.value.address,
                        overrides: {
                            standard: editStandardValue.value,
                            daf: editDafValue.value
                        }
                    })
                });
                const data = await res.json();
                if (res.ok) {
                    editorMessage.value = "Saved!";
                    editorMessageType.value = "success";
                    // Reload the template to reflect changes while preserving state
                    setTimeout(async () => {
                        closeEditor();
                        if (t) await loadTemplate(t, true);
                    }, 500);
                } else {
                    editorMessage.value = data.error || "Save failed";
                    editorMessageType.value = "error";
                }
            } catch (e) {
                editorMessage.value = e.message;
                editorMessageType.value = "error";
            } finally {
                isSavingCell.value = false;
            }
        };

        onMounted(() => {
            fetchTemplates();
        });

        return {
            templates,
            searchQuery,
            filteredTemplates,
            selectedTemplateName,
            currentTemplate,
            currentTemplateName,
            currentTemplateFingerprint,
            templateLayout,
            currentSheetName,
            templateNotes,
            clientProfile,
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
            formatTime,
            editingCell,
            editStandardValue,
            editDafValue,
            isSavingCell,
            editorMessage,
            editorMessageType,
            openCellEditor,
            closeEditor,
            saveCellOverrides,
            // Notes
            isEditingNotes,
            isSavingNotes,
            editingNotesText,
            templateNotes,
            saveNotes,
            cancelEditNotes
        };
    }
};
