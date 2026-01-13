import { ref, computed, onMounted } from 'vue';

export default {
    template: `
        <div class="inspector-view fade-in">
            <h1>Data Inspector & Registry</h1>
            
            <div class="inspector-layout">
                <!-- Sidebar: History -->
                <div class="history-sidebar card">
                    <h3>Recent Runs</h3>
                     <div v-if="historyList.length === 0" style="color: #94a3b8; font-size: 0.875rem;">No history found.</div>
                    <div class="history-list">
                        <div v-for="run in historyList" :key="run.filename" 
                             class="history-item" @click="loadHistoryItem(run)">
                            <div class="h-date">{{ formatTime(run.timestamp) }}</div>
                            <div class="h-file">{{ run.output_file }}</div>
                            <div class="h-stats">{{ run.item_count }} items • {{ run.status }}</div>
                        </div>
                    </div>
                    <button class="btn-small" @click="fetchHistory" style="width: 100%; margin-top: 1rem;">Refresh List</button>
                </div>

                <!-- Main: Details -->
                <div class="inspector-main card">
                     <div class="flex-row" style="display: flex; gap: 1rem; align-items: flex-end; margin-bottom: 1rem;">
                        <div style="flex-grow: 1;">
                            <label style="display: block; margin-bottom: 0.5rem; color: #94a3b8;">Load Metadata File (Manual)</label>
                            <input type="file" @change="loadMetadataFile" accept=".json" />
                        </div>
                         <button class="nav-btn" @click="clearInspector" v-if="inspectorData">Clear</button>
                     </div>

                    <div v-if="!inspectorData" style="text-align: center; padding: 2rem; color: #64748b;">
                        <p>Select a run from the left 👈 or upload a file.</p>
                    </div>

                    <div v-if="inspectorData">
                         <div class="status-box info" style="margin-bottom: 1rem; display: flex; justify-content: space-between; align-items: center;">
                            <div>
                                <strong>Viewing:</strong> {{ inspectorData.output_file || 'Uploaded File' }}
                                <span style="opacity: 0.7; margin-left: 1rem;">{{ inspectorData.timestamp }}</span>
                            </div>
                            <button v-if="inspectorData.output_path_absolute" class="btn-small" 
                                    style="background: #2563eb; color: white; border: none;"
                                    @click="downloadExcel(inspectorData.output_path_absolute)">
                                Download .xlsx 📥
                            </button>
                         </div>
                    
                        <div class="table-container">
                            <table class="data-table">
                                <thead>
                                    <tr>
                                        <th>#</th>
                                        <th>PO</th>
                                        <th>Item Code</th>
                                        <th>Description</th>
                                        <th>PCS</th>
                                        <th>SQFT</th>
                                        <th>Pallets</th>
                                        <th>Net/Gross/CBM</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <tr v-for="(row, index) in inspectorItems" :key="index">
                                        <td>{{ index + 1 }}</td>
                                        <td contenteditable="true" spellcheck="false">{{ row.po }}</td>
                                        <td contenteditable="true" spellcheck="false">{{ row.item }}</td>
                                        <td contenteditable="true" spellcheck="false">{{ row.description }}</td>
                                        <td contenteditable="true" spellcheck="false">{{ row.pcs }}</td>
                                        <td contenteditable="true" spellcheck="false">{{ row.sqft }}</td>
                                        <td contenteditable="true" spellcheck="false">{{ row.pallet_count }}</td>
                                        <td contenteditable="true" spellcheck="false">{{ row.net }} / {{ row.gross }} / {{ row.cbm }}</td>
                                    </tr>
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    `,
    setup() {
        const uploadedMetadata = ref(null);
        const historyList = ref([]);

        // Computed
        const inspectorData = computed(() => {
            return uploadedMetadata.value;
        });

        const inspectorItems = computed(() => {
            return inspectorData.value?.database_export?.packing_list_items || [];
        });

        // Methods
        const fetchHistory = async () => {
            try {
                const res = await fetch('/api/history');
                if (res.ok) {
                    historyList.value = await res.json();
                }
            } catch (e) {
                console.error("Failed to fetch history", e);
            }
        };

        const loadHistoryItem = async (run) => {
            try {
                const res = await fetch(`/api/history/view?filename=${encodeURIComponent(run.filename)}`);
                if (res.ok) {
                    const data = await res.json();
                    uploadedMetadata.value = data;
                } else {
                    alert("Failed to load history item.");
                }
            } catch (e) {
                console.error("Error loading history item", e);
            }
        };

        const loadMetadataFile = (event) => {
            const file = event.target.files[0];
            if (!file) return;
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    uploadedMetadata.value = JSON.parse(e.target.result);
                } catch (err) { alert("Invalid JSON file"); }
            };
            reader.readAsText(file);
        };

        const clearInspector = () => {
            uploadedMetadata.value = null;
        };

        const downloadExcel = (path) => {
            if (!path) return;
            window.location.href = `/api/download?path=${encodeURIComponent(path)}`;
        };

        const formatTime = (ts) => {
            if (!ts) return '';
            return new Date(ts).toLocaleString();
        };

        // Initialize
        onMounted(() => {
            fetchHistory();
        });

        // API to expose to parent if needed, or just keep internal
        // For simple tab switching, internal state is fine, but if we want to "Search in inspector" from Generator,
        // we might need a shared store. For now, let's keep it self-contained.

        return {
            historyList,
            inspectorData,
            inspectorItems,
            fetchHistory,
            loadHistoryItem,
            loadMetadataFile,
            clearInspector,
            downloadExcel,
            formatTime
        };
    }
};
