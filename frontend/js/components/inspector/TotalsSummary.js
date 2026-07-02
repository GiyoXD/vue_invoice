import { useInspectorStore } from '../../stores/inspectorStore.js';

export default {
    name: 'TotalsSummary',
    template: `
        <div v-if="store.inspectorData" class="flex flex-wrap gap-4 items-center justify-between py-2.5 px-4 bg-slate-900/60 border border-slate-700 rounded-xl mb-4 flex-shrink-0">
            <div class="flex flex-wrap items-center gap-y-2">
                <!-- Amount -->
                <div class="flex flex-col pr-6 border-r border-slate-700/50 min-w-[95px]">
                    <span class="text-xs font-medium text-slate-400 uppercase tracking-wider">Amount</span>
                    <span class="text-sm font-bold text-purple-400 mt-0.5">$ {{ formatNumber(store.inspectorTotals.amount) }}</span>
                </div>

                <!-- SQFT (PCS) -->
                <div class="flex flex-col px-6 border-r border-slate-700/50 min-w-[135px]">
                    <span class="text-xs font-medium text-slate-400 uppercase tracking-wider">SQFT (PCS)</span>
                    <div class="mt-0.5 font-bold text-sm">
                        <span class="text-emerald-400">{{ formatNumber(store.inspectorTotals.sqft) }}</span>
                        <span class="text-slate-400 font-medium ml-1.5 text-xs">({{ formatNumber(store.inspectorTotals.pcs) }} pcs)</span>
                    </div>
                </div>

                <!-- Pallets -->
                <div class="flex flex-col px-6 border-r border-slate-700/50 min-w-[75px]">
                    <span class="text-xs font-medium text-slate-400 uppercase tracking-wider">Pallets</span>
                    <span class="text-sm font-bold text-yellow-400 mt-0.5">{{ formatNumber(store.inspectorTotals.pallets) }}</span>
                </div>

                <!-- Weight Net/Gross -->
                <div class="flex flex-col px-6 border-r border-slate-700/50 min-w-[160px]">
                    <span class="text-xs font-medium text-slate-400 uppercase tracking-wider">Weight (Net/Gross)</span>
                    <div class="mt-0.5 font-bold text-sm">
                        <span class="text-cyan-400">{{ formatNumber(store.inspectorTotals.net) }}</span>
                        <span class="text-slate-500 mx-0.5">/</span>
                        <span class="text-orange-400">{{ formatNumber(store.inspectorTotals.gross) }}</span>
                        <span class="text-slate-400 font-medium ml-0.5 text-xs">kg</span>
                    </div>
                </div>

                <!-- Volume -->
                <div class="flex flex-col px-6 border-r border-slate-700/50 min-w-[85px]">
                    <span class="text-xs font-medium text-slate-400 uppercase tracking-wider">CBM</span>
                    <span class="text-sm font-bold text-teal-400 mt-0.5">{{ formatNumber(store.inspectorTotals.cbm) }} <span class="text-slate-400 font-medium text-xs">m³</span></span>
                </div>

                <!-- Recommended Truck -->
                <div class="flex flex-col pl-6 min-w-[220px]">
                    <span class="text-xs font-medium text-slate-400 uppercase tracking-wider font-semibold">Recommended Truck</span>
                    <div class="flex items-center gap-2 mt-0.5">
                        <span v-if="store.recommendedTruckInfo" class="text-xs font-bold px-2.5 py-0.5 rounded-full border transition-all" :class="store.recommendedTruckInfo.color" :title="store.recommendedTruckInfo.description">
                            {{ store.recommendedTruckInfo.displayName }}
                        </span>
                        <span v-else class="text-xs font-bold text-slate-500">—</span>
                        <label class="flex items-center gap-1 cursor-pointer text-[10px] font-bold text-slate-400 uppercase tracking-wider select-none bg-slate-900 border border-slate-700/50 rounded px-1.5 py-0.5 hover:border-slate-500 transition-colors">
                            <input type="checkbox" v-model="store.isWideCargo" accent-color="#10b981" class="rounded bg-slate-950 border-slate-800 w-3 h-3 cursor-pointer" />
                            <span class="px-1">Wide</span>
                        </label>
                    </div>
                </div>
            </div>
            <button class="px-4 py-1.5 bg-slate-700 hover:bg-slate-600 text-slate-200 rounded-lg transition-colors font-medium text-xs flex-shrink-0" @click="store.clearInspector">Clear View</button>
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
