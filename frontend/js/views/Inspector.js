import { ref, computed, onMounted, watch } from 'vue';
import { recommendTruck } from '../utils/truck.js';

export default {
    template: `
        <div class="max-w-[1600px] mx-auto py-8 fade-in h-screen flex flex-col">
            <h1 class="text-4xl font-extrabold tracking-tight text-transparent bg-clip-text bg-gradient-to-r from-blue-400 to-emerald-400 drop-shadow-md flex-shrink-0">Data Inspector & Registry</h1>
            
            <div class="flex gap-6 mt-8 flex-grow min-h-0">
                <!-- Sidebar: History -->
                <div class="w-[300px] bg-slate-800/80 backdrop-blur-md border border-slate-700/50 shadow-2xl rounded-2xl p-6 flex flex-col flex-shrink-0">
                    <h3 class="text-xl font-bold text-slate-100 mb-4 mt-0">Recent Runs</h3>
                     <div v-if="historyList.length === 0" class="text-secondary text-sm">No history found.</div>
                    <div class="flex-1 overflow-y-auto flex flex-col gap-2 pr-2">
                        <div v-for="run in historyList" :key="run.filename" 
                             class="p-3 bg-slate-900/50 border border-slate-700 rounded-lg cursor-pointer transition-all hover:border-blue-500 hover:bg-slate-800" :class="run.type" @click="loadHistoryItem(run)">
                            <div class="text-emerald-400 font-bold text-sm mb-1 truncate drop-shadow-sm" :title="run.output_file">{{ run.output_file }}</div>
                            <div class="text-slate-400 text-xs mb-2">{{ formatTime(run.timestamp) }}</div>
                            <div class="text-slate-500 text-xs flex items-center">
                                <span class="text-emerald-500 font-bold mr-1">{{ run.item_count }}</span> items • {{ run.status }}
                            </div>
                        </div>
                    </div>
                    <button class="w-full px-4 py-2 bg-slate-700 hover:bg-slate-600 text-slate-200 rounded-lg shadow-sm transition-colors mt-4 flex-shrink-0" @click="fetchHistory">Refresh List</button>
                </div>

                <!-- Main: Details -->
                <div class="flex-1 bg-slate-800/80 backdrop-blur-md border border-slate-700/50 shadow-2xl rounded-2xl p-6 flex flex-col min-w-0">                     <!-- Quick Check Totals Summary Panel -->
                     <div v-if="inspectorData" class="flex flex-wrap gap-4 items-center justify-between p-4 bg-slate-900/60 border border-slate-700 rounded-xl mb-4 flex-shrink-0">
                         <div class="flex flex-wrap gap-6 items-center">
                             <div class="flex flex-col">
                                 <span class="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Total PCS</span>
                                 <span class="text-lg font-bold text-blue-400 mt-0.5">{{ formatNumber(inspectorTotals.pcs) }}</span>
                             </div>
                             <div class="h-8 w-px bg-slate-700/50"></div>
                             <div class="flex flex-col">
                                 <span class="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Total SQFT</span>
                                 <span class="text-lg font-bold text-emerald-400 mt-0.5">{{ formatNumber(inspectorTotals.sqft) }}</span>
                             </div>
                             <div class="h-8 w-px bg-slate-700/50"></div>
                             <div class="flex flex-col">
                                 <span class="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Total Pallets</span>
                                 <span class="text-lg font-bold text-yellow-400 mt-0.5">{{ formatNumber(inspectorTotals.pallets) }}</span>
                             </div>
                             <div class="h-8 w-px bg-slate-700/50"></div>
                             <div class="flex flex-col">
                                 <span class="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Total Net (KGS)</span>
                                 <span class="text-lg font-bold text-cyan-400 mt-0.5">{{ formatNumber(inspectorTotals.net) }}</span>
                             </div>
                             <div class="h-8 w-px bg-slate-700/50"></div>
                              <div class="flex flex-col">
                                  <span class="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Total Gross (KGS)</span>
                                  <span class="text-lg font-bold text-orange-400 mt-0.5">{{ formatNumber(inspectorTotals.gross) }}</span>
                              </div>
                              <div class="h-8 w-px bg-slate-700/50"></div>
                              <div class="flex flex-col">
                                  <span class="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Total CBM</span>
                                  <span class="text-lg font-bold text-teal-400 mt-0.5">{{ formatNumber(inspectorTotals.cbm) }} m³</span>
                              </div>
                              <div class="h-8 w-px bg-slate-700/50"></div>
                              <div class="flex flex-col">
                                  <span class="text-[10px] font-bold text-slate-400 uppercase tracking-wider font-semibold">Recommended Truck</span>
                                  <div class="flex items-center gap-2 mt-0.5">
                                      <span v-if="recommendedTruckInfo" class="text-xs font-bold px-2.5 py-1 rounded-full border transition-all" :class="recommendedTruckInfo.color" :title="recommendedTruckInfo.description">
                                          {{ recommendedTruckInfo.displayName }}
                                      </span>
                                      <span v-else class="text-sm font-bold text-slate-500">—</span>
                                      <label class="flex items-center gap-1 cursor-pointer text-[10px] font-bold text-slate-400 uppercase tracking-wider select-none bg-slate-900 border border-slate-700/50 rounded-lg px-2 py-1 hover:border-slate-500 transition-colors">
                                          <input type="checkbox" v-model="isWideCargo" accent-color="#10b981" class="rounded bg-slate-950 border-slate-800 w-3 h-3 cursor-pointer" />
                                          <span>Wide Pallets</span>
                                      </label>
                                  </div>
                              </div>
                              <div class="h-8 w-px bg-slate-700/50"></div>
                              <div class="flex flex-col">
                                  <span class="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Total Amount</span>
                                  <span class="text-lg font-bold text-purple-400 mt-0.5">$ {{ formatNumber(inspectorTotals.amount) }}</span>
                              </div>
                          </div>
                          <button class="px-5 py-2 bg-slate-700 hover:bg-slate-600 text-slate-200 rounded-lg transition-colors font-medium text-sm flex-shrink-0" @click="clearInspector">Clear View</button>
                      </div>

                     <div v-if="!inspectorData" class="text-center p-8 text-slate-500">
                         <p>Select a run from the left 👈 to inspect invoice data.</p>
                     </div>     

                    <div v-if="inspectorData" class="flex flex-col min-h-0 flex-1">
                         <div class="flex items-center justify-between p-4 bg-blue-500/10 border border-blue-500/20 rounded-xl mb-4 flex-shrink-0">
                            <div class="text-sm">
                                <strong class="text-blue-400">Viewing:</strong> <span class="text-slate-200">{{ inspectorData.output_file || currentRun?.output_file || 'Uploaded File' }}</span>
                                <span class="opacity-50 ml-4 text-slate-400">{{ inspectorData.timestamp || currentRun?.timestamp }}</span>
                                <span v-if="currentRun?.type === 'accepted'" class="ml-4 bg-emerald-500/20 text-emerald-400 px-2 py-1 rounded text-xs font-bold border border-emerald-500/30">ACCEPTED</span>
                            </div>
                            <div class="flex gap-2">
                                <button v-if="currentRun?.type === 'processed'" class="px-4 py-2 bg-emerald-500 hover:bg-emerald-400 text-white font-medium rounded-lg shadow-sm transition-colors text-sm" 
                                        @click="acceptCurrentRun">
                                    Accept & Save Check ✅
                                </button>
                                <button v-if="currentRun?.type === 'processed'" class="px-4 py-2 bg-red-500 hover:bg-red-400 text-white font-medium rounded-lg shadow-sm transition-colors text-sm" 
                                        @click="rejectCurrentRun">
                                    Reject & Delete ❌
                                </button>
                                <button v-if="inspectorData.output_path_absolute" class="px-4 py-2 bg-blue-600 hover:bg-blue-500 text-white font-medium rounded-lg shadow-sm transition-colors text-sm" 
                                        @click="downloadExcel(inspectorData.output_path_absolute)">
                                    Download .xlsx 📥
                                </button>
                            </div>
                         </div>

                         <!-- WARNING: Already in DB -->
                         <div v-if="existingInDb && currentRun?.type === 'processed'" class="mb-4 bg-red-500/10 border border-red-500/30 text-red-400 p-4 rounded-xl flex-shrink-0">
                             <strong class="text-lg">⚠️ WARNING: Database Collision</strong>
                             <p class="mt-2 m-0 text-sm">
                                 Invoice <strong class="text-white">{{ currentRun?.filename || inspectorData.output_file }}</strong> is ALREADY in the database.
                                 Accepting it again will <strong class="text-white">REPLACE</strong> all existing records for this invoice.
                             </p>
                         </div>
                    
                        <div class="flex-1 overflow-auto border border-slate-700 rounded-lg custom-scrollbar">
                            <table class="w-full border-collapse text-sm text-slate-300 min-w-max">
                                <thead class="bg-slate-900/80 sticky top-0 z-10 border-b border-slate-700">
                                    <tr>
                                        <th class="p-2 text-left font-medium border-r border-slate-700/50">#</th>
                                        <th class="p-2 text-left font-medium border-r border-slate-700/50">DC</th>
                                        <th class="p-2 text-left font-medium border-r border-slate-700/50">PO</th>
                                        <th class="p-2 text-left font-medium border-r border-slate-700/50">Prod Order</th>
                                        <th class="p-2 text-left font-medium border-r border-slate-700/50">Prod Date</th>
                                        <th class="p-2 text-left font-medium border-r border-slate-700/50">Line No</th>
                                        <th class="p-2 text-left font-medium border-r border-slate-700/50">Direction</th>
                                        <th class="p-2 text-left font-medium border-r border-slate-700/50">Item Code</th>
                                        <th class="p-2 text-left font-medium border-r border-slate-700/50">Ref Code</th>
                                        <th class="p-2 text-left font-medium border-r border-slate-700/50 min-w-[200px]">Description</th>
                                        <th class="p-2 text-left font-medium border-r border-slate-700/50">Level</th>
                                        <th class="p-2 text-right font-medium border-r border-slate-700/50">PCS</th>
                                        <th class="p-2 text-right font-medium border-r border-slate-700/50">SQFT</th>
                                        <th class="p-2 text-right font-medium border-r border-slate-700/50">Pallets</th>

                                        <th class="p-2 text-right font-medium border-r border-slate-700/50">Net</th>
                                        <th class="p-2 text-right font-medium border-r border-slate-700/50">Gross</th>
                                        <th class="p-2 text-right font-medium border-r border-slate-700/50">CBM</th>
                                        <th class="p-2 text-right font-medium border-r border-slate-700/50">Unit Price</th>
                                        <th class="p-2 text-right font-medium">Amount</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <tr v-for="(row, index) in inspectorItems" :key="index" :class="row.is_adjustment ? 'bg-emerald-900/30 font-medium' : 'hover:bg-slate-700/30 border-b border-slate-700/50'">
                                        <td>{{ index + 1 }}</td>
                                        <td>{{ row.col_dc || '' }}</td>
                                        <td>{{ row.col_po || row.po }}</td>
                                        <td>{{ row.col_production_order_no || row.production_order_no || '' }}</td>
                                        <td>{{ row.col_production_date || '' }}</td>
                                        <td>{{ row.col_line_no || '' }}</td>
                                        <td>{{ row.col_direction || '' }}</td>
                                        <td>{{ row.col_item || row.item }}</td>
                                        <td>{{ row.col_reference_code || '' }}</td>
                                        <td>{{ row.col_desc || row.description }}</td>
                                        <td>{{ row.col_level || '' }}</td>
                                        <td>{{ formatNumber(row.col_qty_pcs || row.pcs) }}</td>
                                        <td>{{ formatNumber(row.col_qty_sf || row.sqft) }}</td>
                                        <td>{{ formatNumber(row.col_pallet_count || row.pallet_count) }}</td>

                                        <td>
                                            <span v-if="!row.is_adjustment">{{ formatNumber(row.col_net || row.net) }}</span>
                                        </td>
                                        <td>
                                            <span v-if="!row.is_adjustment">{{ formatNumber(row.col_gross || row.gross) }}</span>
                                        </td>
                                        <td>
                                            <span v-if="!row.is_adjustment">{{ formatNumber(row.col_cbm_raw || row.col_cbm || row.cbm) }}</span>
                                        </td>
                                        <td>{{ formatNumber(row.col_unit_price || '') }}</td>
                                        <td>{{ formatNumber(row.col_amount || row.amount) }}</td>
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

        const inspectorTotals = computed(() => {
            const data = inspectorData.value;
            const gt = data?.footer_data?.grand_total || {};
            
            // Add any extra price adjustments to the backend amount total
            let adjustmentSum = 0;
            if (data?.price_adjustment && Array.isArray(data.price_adjustment)) {
                data.price_adjustment.forEach(adj => {
                    const val = Number(adj[1]);
                    if (isNaN(val)) {
                        throw new Error(`Invalid price_adjustment value: "${adj[1]}" is not a number`);
                    }
                    adjustmentSum += val;
                });
            }

            return {
                pcs: gt.col_qty_pcs || 0,
                sqft: gt.col_qty_sf || 0,
                pallets: gt.col_pallet_count || 0,
                net: gt.col_net || 0,
                gross: gt.col_gross || 0,
                cbm: gt.col_cbm || 0,
                amount: Number(gt.col_amount || 0) + adjustmentSum
            };
        });

        const isWideCargo = ref(false);

        const recommendedTruckInfo = computed(() => {
            const gross = inspectorTotals.value?.gross || 0;
            const cbm = inspectorTotals.value?.cbm || 0;
            const pallets = inspectorTotals.value?.pallets || 0;
            return recommendTruck(gross, cbm, pallets, isWideCargo.value);
        });

        watch(uploadedMetadata, (newData) => {
            if (newData) {
                const desc = String(newData.footer_data?.grand_total?.col_desc || '').toUpperCase();
                const file = String(newData.output_file || '').toUpperCase();
                if (desc.includes('LEATHER') || file.includes('JF') || file.includes('JLFTLT')) {
                    isWideCargo.value = true;
                } else {
                    isWideCargo.value = false;
                }
            }
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
        const formatNumber = (val) => {
            if (val === null || val === undefined || val === '') return '';
            const num = Number(val);
            if (isNaN(num)) return val;
            if (Number.isInteger(num)) return num.toString();
            return parseFloat(num.toFixed(4)).toString();
        };

        const fetchHistory = async (autoLoadLatest = false) => {
            try {
                const res = await fetch('/api/history');
                if (res.ok) {
                    historyList.value = await res.json();
                    if (autoLoadLatest && historyList.value.length > 0) {
                        loadHistoryItem(historyList.value[0]);
                    }
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
            fetchHistory(true);
        });

        // API to expose to parent if needed, or just keep internal
        // For simple tab switching, internal state is fine, but if we want to "Search in inspector" from Generator,
        // we might need a shared store. For now, let's keep it self-contained.

        return {
            historyList,
            inspectorData,
            inspectorItems,
            inspectorTotals,
            fetchHistory,
            loadHistoryItem,
            acceptCurrentRun,
            rejectCurrentRun,
            clearInspector,
            downloadExcel,
            formatTime,
            formatNumber,
            currentRun,
            existingInDb,
            recommendedTruckInfo,
            isWideCargo
        };
    }
};
