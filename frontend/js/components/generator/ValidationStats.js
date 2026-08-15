import { useGeneratorStore } from '../../stores/generatorStore.js';
import { storeToRefs } from 'pinia';
import { computed } from 'vue';

export default {
    name: 'ValidationStats',
    template: `
        <!-- VALIDATION CARD -->
        <div class="bg-emerald-500/10 border border-emerald-500/30 shadow-2xl rounded-2xl p-6 mb-8 delay-100 fade-in" v-if="validationData && !isGenerating && !generationError">
            <div class="flex justify-between items-end mb-4 border-b border-emerald-500/20 pb-3">
                <h3 v-if="isGenerated" class="text-emerald-400 m-0 text-lg font-bold">✅ Invoice Generated Successfully</h3>
                <h3 v-else class="text-sky-400 m-0 text-lg font-bold">📊 Processed Summary (Preview)</h3>
                <span class="text-xs text-emerald-400/70">{{ formatTimestamp(validationData.metadata?.timestamp || validationData.timestamp) }}</span>
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
                    <div class="flex flex-col pl-6 min-w-[280px]">
                        <div class="flex items-center justify-between gap-4">
                            <span class="text-xs font-medium text-slate-400 uppercase tracking-wider font-semibold">Recommended Truck</span>
                            <span class="text-[10px] font-bold text-slate-400 select-none bg-slate-900 border border-slate-700/50 rounded px-1 py-0.5" v-if="detectedPalletDims && (detectedPalletDims.length !== 1.2 || detectedPalletDims.width !== 1.0)">
                                {{ detectedPalletDims.length }}m × {{ detectedPalletDims.width }}m
                            </span>
                        </div>
                        <div class="flex items-center gap-2 mt-0.5">
                            <span v-if="recommendedTruckInfo" class="text-xs font-bold px-2.5 py-0.5 rounded-full border transition-all" :class="recommendedTruckInfo.color" :title="recommendedTruckInfo.description">
                                {{ recommendedTruckInfo.displayName }}
                            </span>
                            <span v-else class="text-xs font-bold text-slate-500">—</span>
                            
                            <div class="flex items-center gap-1.5 bg-slate-950 border border-slate-700/50 rounded px-1.5 py-0.5 select-none">
                                <label class="flex items-center gap-1 cursor-pointer text-[10px] font-bold text-slate-400 uppercase tracking-wider hover:text-slate-200 transition-colors">
                                    <input type="checkbox" v-model="isWideCargo" accent-color="#10b981" class="rounded bg-slate-950 border-slate-850 w-3 h-3 cursor-pointer" />
                                    <span>Wide</span>
                                </label>
                                
                                <span class="text-slate-700 font-bold text-[10px] select-none">|</span>
                                
                                <label class="text-[10px] font-bold text-slate-400 uppercase tracking-wider flex items-center gap-1">
                                    <span>Stack:</span>
                                    <select v-model="maxStackingLayers" class="bg-slate-950 text-slate-300 border-0 rounded px-0.5 py-0 text-[10px] font-bold cursor-pointer focus:ring-0 focus:outline-none">
                                        <option :value="3">3x</option>
                                        <option :value="2">2x</option>
                                        <option :value="1">1x</option>
                                    </select>
                                </label>
                            </div>
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
            generationStatus,
            summaryStats,
            weightStats,
            recommendedTruckInfo,
            isWideCargo,
            maxStackingLayers,
            detectedPalletDims,
            totalAmount
        } = storeToRefs(store);

        const isGenerated = computed(() => generationStatus.value?.type === 'success');

        const formatNumber = (val) => {
            if (val === null || val === undefined || val === '') return '';
            const num = Number(val);
            if (isNaN(num)) return val;
            return num.toLocaleString('en-US', {
                maximumFractionDigits: 4
            });
        };

        const formatTimestamp = (ts) => {
            if (!ts) return '';
            try {
                const d = new Date(ts);
                if (isNaN(d.getTime())) return ts;
                return d.toLocaleString('en-US', {
                    month: 'short',
                    day: 'numeric',
                    year: 'numeric',
                    hour: '2-digit',
                    minute: '2-digit',
                    second: '2-digit'
                });
            } catch {
                return ts;
            }
        };

        return {
            validationData,
            isGenerating,
            generationError,
            isGenerated,
            summaryStats,
            weightStats,
            recommendedTruckInfo,
            isWideCargo,
            maxStackingLayers,
            detectedPalletDims,
            totalAmount,
            formatNumber,
            formatTimestamp
        };
    }
};
