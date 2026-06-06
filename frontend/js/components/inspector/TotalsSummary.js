import { useInspectorStore } from '../../stores/inspectorStore.js';

export default {
    name: 'TotalsSummary',
    template: `
        <div v-if="store.inspectorData" class="flex flex-wrap gap-4 items-center justify-between p-4 bg-slate-900/60 border border-slate-700 rounded-xl mb-4 flex-shrink-0">
            <div class="flex flex-wrap gap-6 items-center">
                <div class="flex flex-col">
                    <span class="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Total PCS</span>
                    <span class="text-lg font-bold text-blue-400 mt-0.5">{{ formatNumber(store.inspectorTotals.pcs) }}</span>
                </div>
                <div class="h-8 w-px bg-slate-700/50"></div>
                <div class="flex flex-col">
                    <span class="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Total SQFT</span>
                    <span class="text-lg font-bold text-emerald-400 mt-0.5">{{ formatNumber(store.inspectorTotals.sqft) }}</span>
                </div>
                <div class="h-8 w-px bg-slate-700/50"></div>
                <div class="flex flex-col">
                    <span class="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Total Pallets</span>
                    <span class="text-lg font-bold text-yellow-400 mt-0.5">{{ formatNumber(store.inspectorTotals.pallets) }}</span>
                </div>
                <div class="h-8 w-px bg-slate-700/50"></div>
                <div class="flex flex-col">
                    <span class="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Total Net (KGS)</span>
                    <span class="text-lg font-bold text-cyan-400 mt-0.5">{{ formatNumber(store.inspectorTotals.net) }}</span>
                </div>
                <div class="h-8 w-px bg-slate-700/50"></div>
                <div class="flex flex-col">
                    <span class="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Total Gross (KGS)</span>
                    <span class="text-lg font-bold text-orange-400 mt-0.5">{{ formatNumber(store.inspectorTotals.gross) }}</span>
                </div>
                <div class="h-8 w-px bg-slate-700/50"></div>
                <div class="flex flex-col">
                    <span class="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Total CBM</span>
                    <span class="text-lg font-bold text-teal-400 mt-0.5">{{ formatNumber(store.inspectorTotals.cbm) }} m³</span>
                </div>
                <div class="h-8 w-px bg-slate-700/50"></div>
                <div class="flex flex-col">
                    <span class="text-[10px] font-bold text-slate-400 uppercase tracking-wider font-semibold">Recommended Truck</span>
                    <div class="flex items-center gap-2 mt-0.5">
                        <span v-if="store.recommendedTruckInfo" class="text-xs font-bold px-2.5 py-1 rounded-full border transition-all" :class="store.recommendedTruckInfo.color" :title="store.recommendedTruckInfo.description">
                            {{ store.recommendedTruckInfo.displayName }}
                        </span>
                        <span v-else class="text-sm font-bold text-slate-500">—</span>
                        <label class="flex items-center gap-1 cursor-pointer text-[10px] font-bold text-slate-400 uppercase tracking-wider select-none bg-slate-900 border border-slate-700/50 rounded-lg px-2 py-1 hover:border-slate-500 transition-colors">
                            <input type="checkbox" v-model="store.isWideCargo" accent-color="#10b981" class="rounded bg-slate-950 border-slate-800 w-3 h-3 cursor-pointer" />
                            <span>Wide Pallets</span>
                        </label>
                    </div>
                </div>
                <div class="h-8 w-px bg-slate-700/50"></div>
                <div class="flex flex-col">
                    <span class="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Total Amount</span>
                    <span class="text-lg font-bold text-purple-400 mt-0.5">$ {{ formatNumber(store.inspectorTotals.amount) }}</span>
                </div>
            </div>
            <button class="px-5 py-2 bg-slate-700 hover:bg-slate-600 text-slate-200 rounded-lg transition-colors font-medium text-sm flex-shrink-0" @click="store.clearInspector">Clear View</button>
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
