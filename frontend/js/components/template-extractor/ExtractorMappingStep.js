import { useTemplateExtractorStore } from '../../stores/templateExtractorStore.js';

export default {
    name: 'ExtractorMappingStep',
    template: `
        <div class="bg-slate-800/80 backdrop-blur-md border border-slate-700/50 shadow-2xl rounded-2xl p-8 mb-8 delay-100 animate-in fade-in" v-if="store.currentStep === 2">
            <h2>2. Map Unrecognized Headers</h2>
            <p class="text-secondary mb-6">
                We found some headers we don't recognize. Please map them to system fields.
            </p>

            <div class="form-group">
                <label>Template Prefix (Unique ID)</label>
                <input type="text" v-model="store.filePrefix" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-4 py-2 text-slate-100 focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 transition-all" placeholder="e.g. MOTO, JLFHM" />
                
                <!-- Show preview of what will be created -->
                <div v-if="store.filePrefix && store.isDualMode" class="mt-2 py-2 px-3 bg-blue-100 border border-blue-200 rounded-md text-sm text-blue-300">
                    📁 Will create: <strong>{{ store.filePrefix }}</strong> (KH + VN variants) in database
                </div>
                <div v-else-if="store.filePrefix && store.singleFileSuffix" class="mt-2 py-2 px-3 bg-blue-100 border border-blue-200 rounded-md text-sm text-blue-300">
                    📁 Will create: <strong>{{ store.filePrefix }}</strong> ({{ store.singleFileSuffix }} variant) in database
                </div>
            </div>

            <!-- PRICING MODE SELECTOR -->
            <div class="form-group mt-4">
                <label>Pricing Mode</label>
                <select v-model="store.pricingMode" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-4 py-2 text-slate-100 focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 transition-all w-72">
                    <option value="standard">Standard (SQFT × Unit Price)</option>
                    <option value="net">Net Weight (Global Price per kg)</option>
                </select>
                <p v-if="store.pricingMode === 'net'" class="text-emerald-500 text-sm mt-1">
                    ⚖️ Generator will ask for a global unit price at invoice time.
                </p>
            </div>

            <div v-if="store.allMissingHeaders.length === 0" class="status-box success">
                ✅ All headers recognized automatically!
            </div>

            <div v-else class="flex flex-col gap-2 mt-4 max-h-[400px] overflow-y-auto pr-2 custom-scrollbar">
                <div v-for="(headerText, index) in store.allMissingHeaders" :key="index" style="display: grid; grid-template-columns: 1fr 2fr auto; gap: 1rem; align-items: center;" class="bg-white/5 p-3 rounded-md">
                    <div class="font-bold text-yellow-400 truncate" :title="headerText">"{{ headerText }}"</div>
                    <select v-model="store.userMappings[headerText]" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-4 py-2 text-slate-100 focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 transition-all text-sm" :disabled="store.confirmedHeaders.includes(headerText)">
                        <option value="" disabled selected>Select a field...</option>
                        <option v-for="opt in store.systemOptions" :value="opt.id" :key="opt.id">
                            {{ opt.label }} ({{ opt.id }})
                        </option>
                    </select>
                    <button 
                        @click="store.toggleMapping(headerText)"
                        :class="[store.confirmedHeaders.includes(headerText) ? 'btn-danger' : 'btn-success', 'btn-sm text-sm py-2 px-4 min-w-[90px] h-full whitespace-nowrap']">
                        {{ store.confirmedHeaders.includes(headerText) ? 'Remove' : 'Add' }}
                    </button>
                </div>
            </div>
            
            <!-- FOOTER MAPPINGS -->
            <div v-if="store.allMissingFooters.length > 0" class="mt-8">
                <div class="flex items-center gap-2 mb-2">
                    <span class="text-xl">🔍</span>
                    <h3 class="m-0 text-emerald-400">Unconfirmed Footer Label</h3>
                </div>
                <p class="text-secondary mb-4 text-sm">
                    We detected a potential footer label via partial match. If you confirm it, it will be mapped permanently so future templates are scanned exactly.
                </p>
                <div v-for="(footerText, idx) in store.allMissingFooters" :key="'f'+idx" class="bg-emerald-900/30 border border-emerald-500/20 p-4 rounded-md mb-2 flex justify-between items-center">
                    <div class="font-bold text-emerald-400">"{{ footerText }}"</div>
                    <button class="btn-sm min-w-24" :class="store.confirmedFooters.includes(footerText) ? 'btn-secondary' : 'btn-success'" @click="store.toggleFooter(footerText)">
                        {{ store.confirmedFooters.includes(footerText) ? 'Confirmed ✓' : 'Confirm It' }}
                    </button>
                </div>
            </div>

            <div class="flex-row flex gap-4 mt-8">
                <button class="px-6 py-3 bg-slate-700 hover:bg-slate-600 text-slate-200 rounded-full transition-colors" @click="store.currentStep = 1">Back</button>
                <button class="w-full px-6 py-3 bg-gradient-to-r from-emerald-500 to-teal-500 hover:from-emerald-400 hover:to-teal-400 text-white font-bold rounded-full shadow-lg shadow-emerald-500/20 transition-all transform hover:-translate-y-0.5 disabled:opacity-50 disabled:cursor-not-allowed disabled:transform-none" @click="store.generateTemplate" :disabled="store.isProcessing || !store.filePrefix">
                    {{ store.isProcessing ? 'Generating...' : 'Create Template' }}
                </button>
            </div>
             
            <!-- PROACTIVE WARNINGS PANEL -->
            <div v-if="store.proactiveWarnings && store.proactiveWarnings.length > 0" class="warning-panel mb-6">
                <div class="warning-header flex items-center gap-2 mb-3">
                    <span class="warning-icon text-xl">⚠️</span>
                    <h3 class="m-0 text-amber-700 text-base">Template Structural Warnings</h3>
                </div>
                <ul class="m-0 pl-6 text-amber-800 text-sm">
                    <li v-for="(msg, idx) in store.proactiveWarnings" :key="idx" class="mb-2 leading-snug">
                        {{ msg }}
                    </li>
                </ul>
            </div>

            <div v-if="store.statusMessage" :class="['status-box', store.statusType]" class="mt-4">
                {{ store.statusMessage }}
            </div>
        </div>
    `,
    setup() {
        const store = useTemplateExtractorStore();
        return { store };
    }
};
