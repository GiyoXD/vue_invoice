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
                <div class="flex flex-col pl-6 min-w-[280px]">
                    <div class="flex items-center gap-2">
                        <span class="text-xs font-medium text-slate-400 uppercase tracking-wider font-semibold">Recommended Truck</span>
                        <span class="text-[10px] font-bold text-slate-300 select-none bg-slate-900 border border-slate-700/60 rounded px-1.5 py-0.5" v-if="store.detectedPalletDims">
                            {{ store.detectedPalletDims.length }}m × {{ store.detectedPalletDims.width }}m<span v-if="store.detectedPalletDims.height"> × {{ store.detectedPalletDims.height }}m</span>
                        </span>
                    </div>
                    <div class="flex items-center gap-2.5 mt-1 flex-wrap">
                        <span v-if="store.recommendedTruckInfo" class="text-xs font-bold px-2.5 py-0.5 rounded-full border transition-all" :class="store.recommendedTruckInfo.color" :title="store.recommendedTruckInfo.description">
                            {{ store.recommendedTruckInfo.tier === 'optimal' ? '⭐ ' : '' }}{{ store.recommendedTruckInfo.displayName }}<span v-if="store.recommendedTruckInfo.utilization" class="text-[9px] opacity-80 ml-1 font-normal">({{ store.recommendedTruckInfo.utilization.lengthPct }}% len<span v-if="store.recommendedTruckInfo.utilization.heightPct">, {{ store.recommendedTruckInfo.utilization.heightPct }}% hgt</span>)</span>
                        </span>
                        <span v-else class="text-xs font-bold text-slate-500">—</span>

                        <div v-if="store.recommendedTruckInfo?.alternatives" class="flex items-center gap-1.5">
                            <span v-if="store.recommendedTruckInfo.alternatives.buffer" class="text-[10px] font-semibold px-1.5 py-0.5 rounded bg-slate-900 border border-slate-700 text-slate-300" :title="'Safe Buffer (+1 Size): ' + store.recommendedTruckInfo.alternatives.buffer.displayName">
                                🛡️ Buffer: {{ store.recommendedTruckInfo.alternatives.buffer.name }}
                            </span>
                            <span v-if="store.recommendedTruckInfo.alternatives.tight" class="text-[10px] font-semibold px-1.5 py-0.5 rounded bg-slate-900 border border-slate-800 text-slate-500" :title="'Tight (-1 Size): ' + store.recommendedTruckInfo.alternatives.tight.displayName + (store.recommendedTruckInfo.alternatives.tight.reason ? ' (' + store.recommendedTruckInfo.alternatives.tight.reason + ')' : '')">
                                ⚡ Tight: {{ store.recommendedTruckInfo.alternatives.tight.name }}
                            </span>
                        </div>

                        <span class="text-slate-700 font-bold text-xs select-none">|</span>
                        
                        <div class="flex items-center gap-1.5 bg-slate-950 border border-slate-700/50 rounded px-1.5 py-0.5 select-none">
                            <label 
                                class="text-[10px] font-bold text-slate-400 uppercase tracking-wider flex items-center gap-1"
                                title="Max Stacking Layers: Number of pallet tiers stacked vertically"
                            >
                                <span>Stack:</span>
                                <select v-model="store.maxStackingLayers" class="bg-slate-950 text-slate-300 border-0 rounded px-0.5 py-0 text-[10px] font-bold cursor-pointer focus:ring-0 focus:outline-none">
                                    <option :value="3" title="3 tiers high">3 Tiers (3x)</option>
                                    <option :value="2" title="2 tiers high">2 Tiers (2x)</option>
                                    <option :value="1" title="1 tier only (No stacking)">1 Tier (1x)</option>
                                </select>
                            </label>
                            <span v-if="store.recommendedTruckInfo?.safetyAlert" 
                                  class="text-[10px] font-bold text-amber-300 bg-amber-950/80 border border-amber-600/60 rounded px-1.5 py-0.5 animate-pulse cursor-help select-none"
                                  :title="store.recommendedTruckInfo.safetyAlert.message">
                                {{ store.recommendedTruckInfo.safetyAlert.label }}
                            </span>
                            
                            <span class="text-slate-700 font-bold text-[10px] select-none">|</span>
                            
                            <label 
                                class="text-[10px] font-bold text-slate-400 uppercase tracking-wider flex items-center gap-1"
                                title="Pallet Loading Orientation: Width-wise (long side spans width) vs Length-wise (long side runs along length)"
                            >
                                <span>Load:</span>
                                <select v-model="store.loadingOrientation" class="bg-slate-950 text-slate-300 border-0 rounded px-0.5 py-0 text-[10px] font-bold cursor-pointer focus:ring-0 focus:outline-none">
                                    <option value="width" title="Width-wise (long side spans truck width)">Width ↔</option>
                                    <option value="length" title="Length-wise (long side runs along truck length)">Length ↕</option>
                                </select>
                            </label>
                        </div>
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
