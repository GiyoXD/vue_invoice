import { useGeneratorStore } from '../stores/generatorStore.js';
import { storeToRefs } from 'pinia';

export default {
    name: 'ValidationStats',
    template: `
        <!-- VALIDATION CARD -->
        <div class="bg-emerald-500/10 border border-emerald-500/30 shadow-2xl rounded-2xl p-8 mb-8 delay-100 fade-in" v-if="validationData && !isGenerating && !generationError">
            <div class="flex justify-between items-end mb-6 border-b border-emerald-500/20 pb-4">
                <h3 class="text-emerald-400 m-0 text-xl font-bold">✅ Invoice Generated Successfully</h3>
                <span class="text-sm text-emerald-400/70">{{ validationData.timestamp }}</span>
            </div>

            <div v-if="summaryStats" class="grid grid-cols-1 md:grid-cols-3 gap-6 mb-6">
                <div class="flex flex-col gap-1 p-4 bg-emerald-500/5 rounded-xl border border-emerald-500/10">
                    <span class="text-emerald-400/80 text-sm font-bold uppercase tracking-wider">Total Items</span>
                    <span class="text-emerald-300 text-2xl font-mono">{{ summaryStats.total_pcs?.toLocaleString() || 0 }}</span>
                </div>
                <div class="flex flex-col gap-1 p-4 bg-emerald-500/5 rounded-xl border border-emerald-500/10">
                    <span class="text-emerald-400/80 text-sm font-bold uppercase tracking-wider">Total SQFT</span>
                    <span class="text-emerald-300 text-2xl font-mono">{{ summaryStats.total_sqft?.toLocaleString(undefined, {maximumFractionDigits: 2}) || 0 }}</span>
                </div>
                <div class="flex flex-col gap-1 p-4 bg-emerald-500/5 rounded-xl border border-emerald-500/10">
                    <span class="text-emerald-400/80 text-sm font-bold uppercase tracking-wider">Total Pallets</span>
                    <span class="text-emerald-300 text-2xl font-mono">{{ summaryStats.total_pallets || 0 }}</span>
                </div>
            </div>

            <div v-if="weightStats" class="grid grid-cols-1 md:grid-cols-3 gap-6 mb-6">
                <div class="flex flex-col gap-1 p-4 bg-emerald-500/5 rounded-xl border border-emerald-500/10">
                    <span class="text-emerald-400/80 text-sm font-bold uppercase tracking-wider">Net Weight</span>
                    <span class="text-emerald-300 text-2xl font-mono">{{ weightStats.net?.toLocaleString() }} kg</span>
                </div>
                <div class="flex flex-col gap-1 p-4 bg-emerald-500/5 rounded-xl border border-emerald-500/10">
                    <span class="text-emerald-400/80 text-sm font-bold uppercase tracking-wider">Gross Weight</span>
                    <span class="text-emerald-300 text-2xl font-mono">{{ weightStats.gross?.toLocaleString() }} kg</span>
                </div>
                <div class="flex flex-col gap-1 p-4 bg-emerald-500/5 rounded-xl border border-emerald-500/10">
                    <span class="text-emerald-400/80 text-sm font-bold uppercase tracking-wider">Total CBM</span>
                    <span class="text-emerald-300 text-2xl font-mono">{{ weightStats.cbm?.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 3}) }} m³</span>
                </div>
            </div>

            <!-- Recommended Shipping Vehicle Card -->
            <div v-if="recommendedTruckInfo" class="mt-6 p-4 rounded-xl border flex items-center justify-between transition-all" :class="recommendedTruckInfo.color">
                <div class="flex items-center gap-3">
                    <span class="text-2xl">🚛</span>
                    <div class="text-left flex flex-col gap-1">
                        <div class="flex items-center gap-2">
                            <h4 class="m-0 text-sm font-bold uppercase tracking-wider text-slate-400">Recommended Shipping Vehicle</h4>
                            <label class="flex items-center gap-1 cursor-pointer text-[10px] font-bold text-slate-400 uppercase tracking-wider select-none bg-slate-900 border border-slate-700/50 rounded-lg px-2 py-0.5 hover:border-slate-500 transition-colors">
                                <input type="checkbox" v-model="isWideCargo" accent-color="#10b981" class="rounded bg-slate-950 border-slate-800 w-3 h-3 cursor-pointer" />
                                <span>Wide Pallets</span>
                            </label>
                        </div>
                        <p class="m-0 text-lg font-extrabold text-slate-100 mt-0.5">{{ recommendedTruckInfo.displayName }}</p>
                    </div>
                </div>
                <div class="text-right">
                    <span class="text-[10px] font-bold text-slate-400 block uppercase tracking-wider">Tonnage Limit</span>
                    <span class="text-sm font-semibold text-slate-300 mt-0.5">
                        Max {{ recommendedTruckInfo.maxWeight === Infinity ? 'Limit Exceeded' : recommendedTruckInfo.maxWeight.toLocaleString() + ' kg' }}
                    </span>
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
            isWideCargo
        } = storeToRefs(store);

        return {
            validationData,
            isGenerating,
            generationError,
            summaryStats,
            weightStats,
            recommendedTruckInfo,
            isWideCargo
        };
    }
};
