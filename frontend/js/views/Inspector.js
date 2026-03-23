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
                             class="history-item" :class="run.type" @click="loadHistoryItem(run)">
                            <div class="h-date">{{ formatTime(run.timestamp) }}</div>
                            <div class="h-file">{{ run.output_file }}</div>
                            <div class="h-stats">
                                <span class="badge">{{ run.type }}</span>
                                {{ run.item_count }} items • {{ run.status }}
                            </div>
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
                         <div class="status-box info inspector-status-bar" style="margin-bottom: 1rem;">
                            <div class="inspector-status-info">
                                <strong>Viewing:</strong> {{ inspectorData.output_file || currentRun?.output_file || 'Uploaded File' }}
                                <span style="opacity: 0.7; margin-left: 1rem;">{{ inspectorData.timestamp || currentRun?.timestamp }}</span>
                                <span v-if="currentRun?.type === 'accepted'" class="badge accepted" style="margin-left:1rem;">ACCEPTED</span>
                            </div>
                            <div class="inspector-actions">
                                <button v-if="currentRun?.type === 'processed'" class="btn-small btn-accept" 
                                        @click="acceptCurrentRun">
                                    Accept & Save Check ✅
                                </button>
                                <button v-if="currentRun?.type === 'processed'" class="btn-small btn-reject" 
                                        @click="rejectCurrentRun">
                                    Reject & Delete ❌
                                </button>
                                <button v-if="inspectorData.output_path_absolute" class="btn-small" 
                                        style="background: #2563eb; color: white; border: none;"
                                        @click="downloadExcel(inspectorData.output_path_absolute)">
                                    Download .xlsx 📥
                                </button>
                            </div>
                         </div>

                         <!-- WARNING: Already in DB -->
                         <div v-if="existingInDb && currentRun?.type === 'processed'" class="status-box" style="margin-bottom: 1rem; background-color: #fef2f2; border: 1px solid #f87171; color: #991b1b; padding: 1rem; border-radius: 8px;">
                             <strong style="font-size: 1.1em;">⚠️ WARNING: Database Collision</strong>
                             <p style="margin-top: 0.5rem; margin-bottom: 0;">
                                 Invoice <strong>{{ currentRun?.filename || inspectorData.output_file }}</strong> is ALREADY in the database.
                                 Accepting it again will <strong>REPLACE</strong> all existing records for this invoice.
                             </p>
                         </div>
                    
                        <div class="table-container">
                            <table class="data-table">
                                <thead>
                                    <tr>
                                        <th>#</th>
                                        <th>DC</th>
                                        <th>PO</th>
                                        <th>Prod Order</th>
                                        <th>Prod Date</th>
                                        <th>Line No</th>
                                        <th>Direction</th>
                                        <th>Item Code</th>
                                        <th>Ref Code</th>
                                        <th>Description</th>
                                        <th>Level</th>
                                        <th>PCS</th>
                                        <th>SQFT</th>
                                        <th>Pallets</th>
                                        <th>Raw Pallets</th>
                                        <th>Net</th>
                                        <th>Gross</th>
                                        <th>CBM</th>
                                        <th>Unit Price</th>
                                        <th>Amount</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <tr v-for="(row, index) in inspectorItems" :key="index" :style="row.is_adjustment ? 'background: #f0fdf4; font-weight: 500;' : ''">
                                        <td>{{ index + 1 }}</td>
                                        <td contenteditable="true" spellcheck="false">{{ row.col_dc || '' }}</td>
                                        <td contenteditable="true" spellcheck="false">{{ row.col_po || row.po }}</td>
                                        <td contenteditable="true" spellcheck="false">{{ row.col_production_order_no || row.production_order_no || '' }}</td>
                                        <td contenteditable="true" spellcheck="false">{{ row.col_production_date || '' }}</td>
                                        <td contenteditable="true" spellcheck="false">{{ row.col_line_no || '' }}</td>
                                        <td contenteditable="true" spellcheck="false">{{ row.col_direction || '' }}</td>
                                        <td contenteditable="true" spellcheck="false">{{ row.col_item || row.item }}</td>
                                        <td contenteditable="true" spellcheck="false">{{ row.col_reference_code || '' }}</td>
                                        <td contenteditable="true" spellcheck="false">{{ row.col_desc || row.description }}</td>
                                        <td contenteditable="true" spellcheck="false">{{ row.col_level || '' }}</td>
                                        <td contenteditable="true" spellcheck="false">{{ row.col_qty_pcs || row.pcs }}</td>
                                        <td contenteditable="true" spellcheck="false">{{ row.col_qty_sf || row.sqft }}</td>
                                        <td contenteditable="true" spellcheck="false">{{ row.col_pallet_count || row.pallet_count }}</td>
                                        <td contenteditable="true" spellcheck="false">{{ row.col_pallet_count_raw !== undefined ? row.col_pallet_count_raw : '' }}</td>
                                        <td contenteditable="true" spellcheck="false">
                                            <span v-if="!row.is_adjustment">{{ row.col_net || row.net }}</span>
                                        </td>
                                        <td contenteditable="true" spellcheck="false">
                                            <span v-if="!row.is_adjustment">{{ row.col_gross || row.gross }}</span>
                                        </td>
                                        <td contenteditable="true" spellcheck="false">
                                            <span v-if="!row.is_adjustment">{{ row.col_cbm_raw || row.col_cbm || row.cbm }}</span>
                                        </td>
                                        <td contenteditable="true" spellcheck="false">{{ row.col_unit_price || '' }}</td>
                                        <td contenteditable="true" spellcheck="false">{{ row.col_amount || row.amount }}</td>
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
        const currentRun = ref(null);
        const existingInDb = ref(false);

        // Computed
        const inspectorData = computed(() => {
            return uploadedMetadata.value;
        });

        const inspectorItems = computed(() => {
            const data = inspectorData.value;
            if (!data) return [];
            
            let items = [];
            
            // 1. Add Price Adjustments as top rows (highest rows)
            if (data.price_adjustment && Array.isArray(data.price_adjustment)) {
                data.price_adjustment.forEach(adj => {
                    items.push({
                        col_desc: adj[0],
                        col_amount: adj[1],
                        is_adjustment: true
                    });
                });
            }
            
            // 2. Add Main Items — raw_data only (unprocessed, never distributed).
            // Flatten all tables into one list.
            const mainItems = (data.raw_data || []).flat();
            items = items.concat(mainItems);
            
            return items;
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
                const res = await fetch(`/api/history/view?filename=${encodeURIComponent(run.filename)}&source=${run.type || 'run_log'}`);
                if (res.ok) {
                    const data = await res.json();
                    uploadedMetadata.value = data;
                    currentRun.value = run;
                    
                    // Immediately check if this file already exists in the DB
                    if (run.type === 'processed') {
                        try {
                            const checkRes = await fetch('/api/registry/check', {
                                method: 'POST',
                                headers: { 'Content-Type': 'application/json' },
                                body: JSON.stringify({ filename: run.filename })
                            });
                            if (checkRes.ok) {
                                const checkData = await checkRes.json();
                                existingInDb.value = checkData.exists;
                            } else {
                                existingInDb.value = false;
                            }
                        } catch (e) {
                            existingInDb.value = false;
                        }
                    } else {
                        // If it's already accepted, it exists, but we don't need the red warning box for standard viewing
                        existingInDb.value = false;
                    }
                } else {
                    alert("Failed to load history item.");
                }
            } catch (e) {
                console.error("Error loading history item", e);
            }
        };

        const acceptCurrentRun = async () => {
            if (!currentRun.value || !currentRun.value.filename) return;
            
            try {
                // Check if already in DB
                const checkRes = await fetch('/api/registry/check', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ filename: currentRun.value.filename })
                });
                
                if (checkRes.ok) {
                    const checkData = await checkRes.json();
                    if (checkData.exists) {
                        if (!confirm(`⚠️ WARNING: The invoice "${currentRun.value.filename}" is ALREADY in the database.\n\nDo you want to REPLACE the existing data?`)) {
                            return;
                        }
                    } else {
                        if (!confirm(`Accept and save "${currentRun.value.filename}" to database?`)) return;
                    }
                } else {
                    if (!confirm(`Accept and save "${currentRun.value.filename}" to database?`)) return;
                }
            } catch (e) {
                console.error("Failed to check registry", e);
                if (!confirm(`Accept and save "${currentRun.value.filename}" to database?`)) return;
            }
            
            try {
                const res = await fetch('/api/registry/accept', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ filename: currentRun.value.filename })
                });
                const result = await res.json();
                if (res.ok) {
                    alert(result.message || "Saved successfully!");
                    // Refresh history and clear view or update currentRun
                    await fetchHistory();
                    // Find the newly accepted item in history and load it
                    const acceptedItem = historyList.value.find(h => h.filename === currentRun.value.filename && h.type === 'accepted');
                    if (acceptedItem) {
                        loadHistoryItem(acceptedItem);
                    } else {
                        uploadedMetadata.value = null;
                        currentRun.value = null;
                    }
                } else {
                    alert("Error: " + result.error);
                }
            } catch (e) {
                console.error("Failed to accept run", e);
                alert("Failed to accept run");
            }
        };

        const rejectCurrentRun = async () => {
            if (!currentRun.value || !currentRun.value.filename) return;
            if (!confirm(`Reject and delete "${currentRun.value.filename}"? This cannot be undone.`)) return;
            
            try {
                const res = await fetch('/api/registry/reject', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ filename: currentRun.value.filename })
                });
                const result = await res.json();
                if (res.ok) {
                    alert(result.message || "Deleted successfully!");
                    uploadedMetadata.value = null;
                    currentRun.value = null;
                    await fetchHistory();
                } else {
                    alert("Error: " + result.error);
                }
            } catch (e) {
                console.error("Failed to reject run", e);
                alert("Failed to reject run");
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
            acceptCurrentRun,
            rejectCurrentRun,
            loadMetadataFile,
            clearInspector,
            downloadExcel,
            formatTime,
            currentRun,
            existingInDb
        };
    }
};
