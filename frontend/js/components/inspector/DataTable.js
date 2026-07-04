import { useInspectorStore } from '../../stores/inspectorStore.js';

export default {
    name: 'DataTable',
    template: `
        <div v-if="store.inspectorData" class="flex flex-col min-h-0 flex-1">
            <div class="flex items-center justify-between p-4 bg-blue-500/10 border border-blue-500/20 rounded-xl mb-4 flex-shrink-0">
                <div class="text-sm">
                    <strong class="text-blue-400">Viewing:</strong> <span class="text-slate-200">{{ store.inspectorData.output_file || store.currentRun?.output_file || 'Uploaded File' }}</span>
                    <span class="opacity-50 ml-4 text-slate-400">{{ store.inspectorData.timestamp || store.currentRun?.timestamp }}</span>
                    <span v-if="store.currentRun?.type === 'accepted'" class="ml-4 bg-emerald-500/20 text-emerald-400 px-2 py-1 rounded text-xs font-bold border border-emerald-500/30">ACCEPTED</span>
                </div>
                <div class="flex gap-2">
                    <button v-if="store.currentRun?.type === 'processed'" class="px-4 py-2 bg-emerald-500 hover:bg-emerald-400 text-white font-medium rounded-lg shadow-sm transition-colors text-sm" 
                            @click="store.acceptCurrentRun">
                        Accept & Save Check ✅
                    </button>
                    <button v-if="store.currentRun?.type === 'processed'" class="px-4 py-2 bg-red-500 hover:bg-red-400 text-white font-medium rounded-lg shadow-sm transition-colors text-sm" 
                            @click="store.rejectCurrentRun">
                        Reject & Delete ❌
                    </button>
                    <button v-if="store.inspectorData.output_path_absolute" class="px-4 py-2 bg-blue-600 hover:bg-blue-500 text-white font-medium rounded-lg shadow-sm transition-colors text-sm" 
                            @click="store.downloadExcel(store.inspectorData.output_path_absolute)">
                        Download .xlsx 📥
                    </button>
                </div>
            </div>

            <!-- WARNING: Already in DB -->
            <div v-if="store.existingInDb && store.currentRun?.type === 'processed'" class="mb-4 bg-red-500/10 border border-red-500/30 text-red-400 p-4 rounded-xl flex-shrink-0">
                <strong class="text-lg">⚠️ WARNING: Database Collision</strong>
                <p class="mt-2 m-0 text-sm">
                    Invoice <strong class="text-white">{{ store.currentRun?.filename || store.inspectorData.output_file }}</strong> is ALREADY in the database.
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
                        <tr v-for="(row, index) in store.inspectorItems" :key="index" :class="row.is_adjustment ? 'bg-emerald-900/30 font-medium' : 'hover:bg-slate-700/30 border-b border-slate-700/50'">
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
                            <td>{{ row.col_pallet_no || '' }}</td>

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
    `,
    setup() {
        const store = useInspectorStore();

        const formatNumber = (val) => {
            if (val === null || val === undefined || val === '') return '';
            const num = Number(val);
            if (isNaN(num)) return val;
            if (Number.isInteger(num)) return num.toString();
            return parseFloat(num.toFixed(4)).toString();
        };

        return { store, formatNumber };
    }
};
