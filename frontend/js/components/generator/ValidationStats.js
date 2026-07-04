import { useGeneratorStore } from '../../stores/generatorStore.js';
import { storeToRefs } from 'pinia';

export default {
    name: 'ValidationStats',
    template: `
        <!-- VALIDATION CARD -->
        <div class="bg-emerald-500/10 border border-emerald-500/30 shadow-2xl rounded-2xl p-6 mb-8 delay-100 fade-in" v-if="validationData && !isGenerating && !generationError">
            <div class="flex justify-between items-end mb-4 border-b border-emerald-500/20 pb-3">
                <h3 class="text-emerald-400 m-0 text-lg font-bold">✅ Invoice Generated Successfully</h3>
                <span class="text-xs text-emerald-400/70">{{ validationData.timestamp }}</span>
            </div>

            <!-- Redesigned Summary Panel (Concise and Spacious) -->
            <div class="flex flex-wrap gap-4 items-center justify-between py-2.5 px-4 bg-slate-900/60 border border-slate-700/50 rounded-xl">
                <div class="flex flex-wrap items-center gap-y-2">
                    <!-- Amount -->
                    <div class="flex flex-col pr-6 border-r border-slate-700/50 min-w-[95px]">
                        <span class="text-xs font-medium text-slate-400 uppercase tracking-wider">Amount</span>
                        <span class="text-sm font-bold text-purple-400 mt-0.5">$ {{ formatNumber(totalAmount) }}</span>
                    </div>

                    <!-- SQFT (PCS) -->
                    <div class="flex flex-col px-6 border-r border-slate-700/50 min-w-[135px]">
                        <span class="text-xs font-medium text-slate-400 uppercase tracking-wider">SQFT (PCS)</span>
                        <div class="mt-0.5 font-bold text-sm">
                            <span class="text-emerald-400">{{ formatNumber(summaryStats?.total_sqft) }}</span>
                            <span class="text-slate-400 font-medium ml-1.5 text-xs">({{ formatNumber(summaryStats?.total_pcs) }} pcs)</span>
                        </div>
                    </div>

                    <!-- Pallets -->
                    <div class="flex flex-col px-6 border-r border-slate-700/50 min-w-[75px]">
                        <span class="text-xs font-medium text-slate-400 uppercase tracking-wider">Pallets</span>
                        <span class="text-sm font-bold text-yellow-400 mt-0.5">{{ formatNumber(summaryStats?.total_pallets) }}</span>
                    </div>

                    <!-- Weight Net/Gross -->
                    <div class="flex flex-col px-6 border-r border-slate-700/50 min-w-[160px]">
                        <span class="text-xs font-medium text-slate-400 uppercase tracking-wider">Weight (Net/Gross)</span>
                        <div class="mt-0.5 font-bold text-sm">
                            <span class="text-cyan-400">{{ formatNumber(weightStats?.net) }}</span>
                            <span class="text-slate-500 mx-0.5">/</span>
                            <span class="text-orange-400">{{ formatNumber(weightStats?.gross) }}</span>
                            <span class="text-slate-400 font-medium ml-0.5 text-xs">kg</span>
                        </div>
                    </div>

                    <!-- Volume -->
                    <div class="flex flex-col px-6 border-r border-slate-700/50 min-w-[85px]">
                        <span class="text-xs font-medium text-slate-400 uppercase tracking-wider">CBM</span>
                        <span class="text-sm font-bold text-teal-400 mt-0.5">{{ formatNumber(weightStats?.cbm) }} <span class="text-slate-400 font-medium text-xs">m³</span></span>
                    </div>

                    <!-- Recommended Truck -->
                    <div class="flex flex-col pl-6 min-w-[220px]">
                        <span class="text-xs font-medium text-slate-400 uppercase tracking-wider font-semibold">Recommended Truck</span>
                        <div class="flex items-center gap-2 mt-0.5">
                            <span v-if="recommendedTruckInfo" class="text-xs font-bold px-2.5 py-0.5 rounded-full border transition-all" :class="recommendedTruckInfo.color" :title="recommendedTruckInfo.description">
                                {{ recommendedTruckInfo.displayName }}
                            </span>
                            <span v-else class="text-xs font-bold text-slate-500">—</span>
                            <label class="flex items-center gap-1 cursor-pointer text-[10px] font-bold text-slate-400 uppercase tracking-wider select-none bg-slate-900 border border-slate-700/50 rounded px-1.5 py-0.5 hover:border-slate-500 transition-colors">
                                <input type="checkbox" v-model="isWideCargo" accent-color="#10b981" class="rounded bg-slate-950 border-slate-800 w-3 h-3 cursor-pointer" />
                                <span class="px-1">Wide</span>
                            </label>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    `,
    setup() {
        const store = useGeneratorStore();
        const {
            validationData,
            isGenerating,
            generationError,
            summaryStats,
            weightStats,
            recommendedTruckInfo,
            isWideCargo,
            totalAmount
        } = storeToRefs(store);

        const formatNumber = (val) => {
            if (val === null || val === undefined || val === '') return '';
            const num = Number(val);
            if (isNaN(num)) return val;
            if (Number.isInteger(num)) return num.toString();
            return parseFloat(num.toFixed(4)).toString();
        };

        return {
            validationData,
            isGenerating,
            generationError,
            summaryStats,
            weightStats,
            recommendedTruckInfo,
            isWideCargo,
            totalAmount,
            formatNumber
        };
    }
};
