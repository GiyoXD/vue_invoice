import { ref, reactive, computed } from 'vue';

export default {
    template: `
        <div class="template-extractor-view fade-in">
            <h1>New Template Extractor</h1>
            
            <!-- STEP 1: UPLOAD -->
            <div class="bg-slate-800/80 backdrop-blur-md border border-slate-700/50 shadow-2xl rounded-2xl p-8 mb-8" v-if="currentStep === 1">
                <h2>1. Analyze Invoice Source</h2>
                <p class="text-secondary mb-4">
                    Upload a sample invoice file. Upload <strong>2 files</strong> to auto-create KH + VN versions.
                </p>
                
                <input type="file" @change="handleFileUpload" accept=".xlsx, .xls" multiple class="block w-full text-sm text-slate-400 file:mr-4 file:py-2.5 file:px-4 file:rounded-full file:border-0 file:text-sm file:font-semibold file:bg-blue-500/10 file:text-blue-400 hover:file:bg-blue-500/20 transition-all cursor-pointer" />
                
                <!-- Show selected files with KH/VN labels -->
                <div v-if="selectedFiles.length > 0" class="mt-4">
                    <div v-for="(file, idx) in selectedFiles" :key="idx" 
                         class="flex items-center gap-3 py-2 px-3 mb-2 bg-white/5 rounded-md border border-white/10">
                        <span v-if="selectedFiles.length === 2" 
                              class="px-2 py-1 rounded text-xs font-bold border"
                              :class="idx === 0 ? 'bg-blue-400/20 text-blue-400 border-blue-400/30' : 'bg-yellow-400/20 text-yellow-400 border-yellow-400/30'">
                            {{ idx === 0 ? 'KH' : 'VN' }}
                        </span>
                        <span class="text-primary">📄 {{ file.name }}</span>
                    </div>
                    
                    <!-- Single-file suffix selector -->
                    <div v-if="selectedFiles.length === 1" class="mt-3 flex items-center gap-3">
                        <label class="text-secondary text-sm">Version suffix:</label>
                        <select v-model="singleFileSuffix" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-4 py-2 text-slate-100 focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 transition-all w-40">
                            <option value="KH">KH version</option>
                            <option value="VN">VN version</option>
                        </select>
                    </div>

                    <!-- Ignore missing description check -->
                    <div class="mt-4 flex items-center gap-2">
                        <input type="checkbox" id="ignore-missing-desc" v-model="ignoreMissingDescription" class="rounded bg-slate-950 border-slate-700 text-blue-500 focus:ring-blue-500 focus:ring-offset-slate-900 w-4 h-4 cursor-pointer" />
                        <label for="ignore-missing-desc" class="text-sm text-slate-300 select-none cursor-pointer flex items-center gap-1.5">
                            ⚠️ Ignore missing description error (Bypass DES column checks)
                        </label>
                    </div>

                    <div v-if="selectedFiles.length > 2" class="status-box error mt-2">
                        ⚠️ Maximum 2 files allowed. Only the first 2 will be used.
                    </div>
                </div>
                
                <button class="w-full px-6 py-3 mt-4 bg-gradient-to-r from-blue-500 to-blue-600 hover:from-blue-400 hover:to-blue-500 text-white font-medium rounded-full shadow-lg shadow-blue-500/30 transition-all transform hover:-translate-y-0.5 disabled:opacity-50 disabled:cursor-not-allowed disabled:transform-none" @click="analyzeFiles" :disabled="selectedFiles.length === 0 || isProcessing">
                    {{ isProcessing ? 'Analyzing...' : 'Analyze & Extract' }}
                </button>
                
                 <div v-if="statusMessage" :class="['status-box', statusType]">
                    <div class="flex flex-col gap-2">
                        <span class="break-words text-xs whitespace-pre-wrap leading-relaxed">{{ statusMessage }}</span>
                        <button v-if="statusType === 'error' && (statusMessage.includes('Missing Description') || statusMessage.includes('description') || statusMessage.includes('ValueError'))" 
                                @click="forceAnalyze" 
                                class="mt-2 px-4 py-2.5 bg-amber-500 hover:bg-amber-400 text-slate-950 text-xs font-bold rounded-lg transition-all self-start flex items-center gap-1.5 shadow-md transform hover:-translate-y-0.5 cursor-pointer">
                            ⚠️ Force Analyze & Bypass This Error
                        </button>
                    </div>
                </div>
            </div>

            <!-- STEP 2: MAP HEADERS -->
            <div class="bg-slate-800/80 backdrop-blur-md border border-slate-700/50 shadow-2xl rounded-2xl p-8 mb-8 delay-100 animate-in fade-in" v-if="currentStep === 2">
                <h2>2. Map Unrecognized Headers</h2>
                <p class="text-secondary mb-6">
                    We found some headers we don't recognize. Please map them to system fields.
                </p>

                <div class="form-group">
                    <label>Template Prefix (Unique ID)</label>
                    <input type="text" v-model="filePrefix" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-4 py-2 text-slate-100 focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 transition-all" placeholder="e.g. MOTO, JLFHM" />
                    
                    <!-- Show preview of what will be created -->
                    <div v-if="filePrefix && isDualMode" class="mt-2 py-2 px-3 bg-blue-100 border border-blue-200 rounded-md text-sm text-blue-300">
                        📁 Will create: <strong>{{ filePrefix }}</strong> (KH + VN variants) in database
                    </div>
                    <div v-else-if="filePrefix && singleFileSuffix" class="mt-2 py-2 px-3 bg-blue-100 border border-blue-200 rounded-md text-sm text-blue-300">
                        📁 Will create: <strong>{{ filePrefix }}</strong> ({{ singleFileSuffix }} variant) in database
                    </div>
                </div>

                <!-- PRICING MODE SELECTOR -->
                <div class="form-group mt-4">
                    <label>Pricing Mode</label>
                    <select v-model="pricingMode" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-4 py-2 text-slate-100 focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 transition-all w-72">
                        <option value="standard">Standard (SQFT × Unit Price)</option>
                        <option value="net">Net Weight (Global Price per kg)</option>
                    </select>
                    <p v-if="pricingMode === 'net'" class="text-emerald-500 text-sm mt-1">
                        ⚖️ Generator will ask for a global unit price at invoice time.
                    </p>
                </div>

                <div v-if="allMissingHeaders.length === 0" class="status-box success">
                    ✅ All headers recognized automatically!
                </div>

                <div v-else class="flex flex-col gap-2 mt-4 max-h-[400px] overflow-y-auto pr-2 custom-scrollbar">
                    <div v-for="(headerText, index) in allMissingHeaders" :key="index" style="display: grid; grid-template-columns: 1fr 2fr auto; gap: 1rem; align-items: center;" class="bg-white/5 p-3 rounded-md">
                        <div class="font-bold text-yellow-400 truncate" :title="headerText">"{{ headerText }}"</div>
                        <select v-model="userMappings[headerText]" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-4 py-2 text-slate-100 focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 transition-all text-sm" :disabled="confirmedHeaders.includes(headerText)">
                            <option value="" disabled selected>Select a field...</option>
                            <option v-for="opt in systemOptions" :value="opt.id">
                                {{ opt.label }} ({{ opt.id }})
                            </option>
                        </select>
                        <button 
                            @click="toggleMapping(headerText)"
                            :class="[confirmedHeaders.includes(headerText) ? 'btn-danger' : 'btn-success', 'btn-sm text-sm py-2 px-4 min-w-[90px] h-full whitespace-nowrap']">
                            {{ confirmedHeaders.includes(headerText) ? 'Remove' : 'Add' }}
                        </button>
                    </div>
                </div>
                
                <!-- FOOTER MAPPINGS -->
                <div v-if="allMissingFooters.length > 0" class="mt-8">
                    <div class="flex items-center gap-2 mb-2">
                        <span class="text-xl">🔍</span>
                        <h3 class="m-0 text-emerald-400">Unconfirmed Footer Label</h3>
                    </div>
                    <p class="text-secondary mb-4 text-sm">
                        We detected a potential footer label via partial match. If you confirm it, it will be mapped permanently so future templates are scanned exactly.
                    </p>
                    <div v-for="(footerText, idx) in allMissingFooters" :key="'f'+idx" class="bg-emerald-900/30 border border-emerald-500/20 p-4 rounded-md mb-2 flex justify-between items-center">
                        <div class="font-bold text-emerald-400">"{{ footerText }}"</div>
                        <button class="btn-sm min-w-24" :class="confirmedFooters.includes(footerText) ? 'btn-secondary' : 'btn-success'" @click="toggleFooter(footerText)">
                            {{ confirmedFooters.includes(footerText) ? 'Confirmed ✓' : 'Confirm It' }}
                        </button>
                    </div>
                </div>

                <div class="flex-row flex gap-4 mt-8">
                    <button class="px-6 py-3 bg-slate-700 hover:bg-slate-600 text-slate-200 rounded-full transition-colors" @click="currentStep = 1">Back</button>
                    <button class="w-full px-6 py-3 bg-gradient-to-r from-emerald-500 to-teal-500 hover:from-emerald-400 hover:to-teal-400 text-white font-bold rounded-full shadow-lg shadow-emerald-500/20 transition-all transform hover:-translate-y-0.5 disabled:opacity-50 disabled:cursor-not-allowed disabled:transform-none" @click="generateTemplate" :disabled="isProcessing || !filePrefix">
                        {{ isProcessing ? 'Generating...' : 'Create Template' }}
                    </button>
                </div>
                 
                <!-- PROACTIVE WARNINGS PANEL -->
                <div v-if="proactiveWarnings && proactiveWarnings.length > 0" class="warning-panel mb-6">
                    <div class="warning-header flex items-center gap-2 mb-3">
                        <span class="warning-icon text-xl">⚠️</span>
                        <h3 class="m-0 text-amber-700 text-base">Template Structural Warnings</h3>
                    </div>
                    <ul class="m-0 pl-6 text-amber-800 text-sm">
                        <li v-for="(msg, idx) in proactiveWarnings" :key="idx" class="mb-2 leading-snug">
                            {{ msg }}
                        </li>
                    </ul>
                </div>

                <div v-if="statusMessage" :class="['status-box', statusType]" class="mt-4">
                    {{ statusMessage }}
                </div>
            </div>

            <!-- STEP 3: SUCCESS -->
            <div class="bg-slate-800/80 backdrop-blur-md border border-slate-700/50 shadow-2xl rounded-2xl p-8 mb-8 text-center delay-100 animate-in zoom-in" v-if="currentStep === 3">
                <div class="text-6xl mb-4">🎉</div>
                <h2>Template Created!</h2>
                <p class="text-secondary mb-4">
                    The template <strong>{{ filePrefix }}</strong> has been configured successfully.
                </p>
                <div v-if="bundlePath" class="bg-emerald-500/10 p-4 rounded-xl mb-6 text-left">
                    <p class="text-emerald-300 m-0 mb-2 text-sm">📁 Bundle created at:</p>
                    <code class="text-emerald-500 text-xs break-all">{{ bundlePath }}</code>
                    <div v-if="generatedPrefixes.length > 1" class="mt-2 pt-2 border-t border-emerald-500/20">
                        <p class="text-emerald-300 m-0 mb-1 text-xs">Contains:</p>
                        <div v-for="p in generatedPrefixes" :key="p" class="text-emerald-400 text-xs">
                            ✅ {{ p }}
                        </div>
                    </div>
                </div>
                <p class="text-secondary mb-8">
                    You can now go to the Generator and process invoices for this company.
                </p>
                <button class="w-full px-6 py-3 mt-4 bg-gradient-to-r from-blue-500 to-blue-600 hover:from-blue-400 hover:to-blue-500 text-white font-medium rounded-full shadow-lg shadow-blue-500/30 transition-all transform hover:-translate-y-0.5" @click="resetFlow">Process Another</button>
            </div>

            <!-- GLOBAL MAPPINGS -->
            <div class="bg-slate-800/80 backdrop-blur-md border border-slate-700/50 shadow-2xl rounded-2xl p-8 mb-8 mt-8">
                <div class="flex justify-between items-center cursor-pointer" @click="showMappings = !showMappings">
                    <h2>Manage Global Mappings</h2>
                    <span>{{ showMappings ? '▲ Collapse' : '▼ Expand' }}</span>
                </div>
                
                <div v-if="showMappings" class="mt-4">
                    <p class="text-secondary mb-4">
                        View and edit the globally recognized mappings. These are used to automatically match headers and sheets in templates.
                    </p>
                    
                    <div class="flex gap-4 mb-4">
                        <select v-model="activeMappingType" @change="switchMappingType($event.target.value)" class="flex-none bg-slate-900 border border-slate-700 rounded-lg px-4 py-2 text-slate-100 focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 transition-all w-72 font-bold">
                            <option value="header_text_mappings">Header Mappings</option>
                            <option value="sheet_mappings">Sheet Mappings</option>
                            <option value="shipping_header_map">Shipping Header Map</option>
                            <option value="footer_label_mappings">Footer Labels (Total)</option>
                        </select>
                        <div class="flex-1 relative">
                            <input type="text" v-model="mappingSearch" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-4 py-2 text-slate-100 focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 transition-all" placeholder="Search..." />
                        </div>
                    </div>

                    <!-- Add New Mapping Row -->
                    <div v-if="activeMappingType === 'sheet_mappings'" style="display: grid; grid-template-columns: 1fr 1fr auto; gap: 0.5rem; align-items: center;" class="mb-4 p-2 bg-emerald-500/5 border border-emerald-500/30 border-dashed rounded-md">
                        <input type="text" v-model="newMappingKey" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-4 py-2 text-slate-100 focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 transition-all" placeholder="Sheet Name (e.g. INV)" />
                        <select v-model="newMappingVal" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-4 py-2 text-slate-100 focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 transition-all">
                            <option value="" disabled selected>Select classification...</option>
                            <option value="aggregation">Aggregation (Invoice / Contract)</option>
                            <option value="processed_tables">Processed Tables (Packing List)</option>
                        </select>
                        <button class="px-6 py-2 bg-blue-500 hover:bg-blue-600 text-white font-medium rounded-lg shadow-sm transition-colors disabled:opacity-50 disabled:cursor-not-allowed min-w-[100px]" @click.prevent="addNewMapping" :disabled="!newMappingKey || !newMappingVal">Add</button>
                    </div>
                    <div v-else style="display: grid; grid-template-columns: 1fr 1fr auto; gap: 0.5rem; align-items: center;" class="mb-4 p-2 bg-emerald-500/5 border border-emerald-500/30 border-dashed rounded-md">
                        <input type="text" v-model="newMappingKey" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-4 py-2 text-slate-100 focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 transition-all" :placeholder="activeMappingType === 'shipping_header_map' ? 'New Keyword (e.g. PO)' : (activeMappingType === 'footer_label_mappings' ? 'New Footer Label (e.g. GRAND TOTAL)' : 'New Input Text (e.g. Qty(SF))')" />
                        
                        <input v-if="activeMappingType === 'footer_label_mappings'" type="text" v-model="newMappingVal" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-4 py-2 text-slate-100 focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 transition-all" placeholder="Auto-filled" :disabled="true" />
                        <select v-else v-model="newMappingVal" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-4 py-2 text-slate-100 focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 transition-all">
                            <option value="" disabled selected>Select system field...</option>
                            <option v-for="opt in systemOptions" :value="opt.id">{{ opt.label }} ({{ opt.id }})</option>
                        </select>
                        
                        <button class="px-6 py-2 bg-blue-500 hover:bg-blue-600 text-white font-medium rounded-lg shadow-sm transition-colors disabled:opacity-50 disabled:cursor-not-allowed min-w-[100px]" @click.prevent="addNewMapping" :disabled="!newMappingKey || !newMappingVal">Add</button>
                    </div>

                    <div class="max-h-[400px] overflow-y-auto border border-white/10 rounded-md p-2">
                        <div class="mapping-grid grid gap-2">
                            <!-- Header Row -->
                            <div v-if="activeMappingType === 'sheet_mappings'" style="display: grid; grid-template-columns: 1fr 1fr auto; gap: 0.5rem; align-items: center;" class="font-bold p-2 border-b border-white/10">
                                <div>Sheet Name</div>
                                <div>Classification</div>
                                <div class="w-20 text-center">Action</div>
                            </div>
                            <div v-else style="display: grid; grid-template-columns: 1fr 1fr auto; gap: 0.5rem; align-items: center;" class="font-bold p-2 border-b border-white/10">
                                <div>{{ activeMappingType === 'shipping_header_map' ? 'Keyword' : (activeMappingType === 'footer_label_mappings' ? 'Footer Target Text' : 'Original Text (Excel)') }}</div>
                                <div>{{ activeMappingType === 'shipping_header_map' ? 'Mapped Target (System)' : (activeMappingType === 'footer_label_mappings' ? 'Type' : 'Mapped Target (System)') }}</div>
                                <div class="w-20 text-center">Action</div>
                            </div>
                            
                            <template v-if="activeMappingType === 'sheet_mappings'">
                                <div v-for="(val, sheetName) in filteredMappings" :key="sheetName" style="display: grid; grid-template-columns: 1fr 1fr auto; gap: 0.5rem; align-items: center;" class="bg-white/5 p-2 rounded">
                                    <input type="text" :value="sheetName" @change="updateMappingHeader(sheetName, $event.target.value)" class="w-full h-9 bg-slate-900 border border-slate-700 rounded px-3 text-slate-100 focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 transition-all text-sm" />
                                    
                                    <select :value="val" @change="updateMappingColId(sheetName, $event.target.value)" class="w-full h-9 bg-slate-900 border border-slate-700 rounded px-3 text-slate-100 focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 transition-all text-sm">
                                        <option value="aggregation">Aggregation (Invoice / Contract)</option>
                                        <option value="processed_tables">Processed Tables (Packing List)</option>
                                    </select>
                                    
                                    <button class="h-9 px-4 bg-red-500/80 hover:bg-red-500 text-white rounded cursor-pointer transition-colors w-full text-sm font-medium" @click="deleteMapping(sheetName)">Delete</button>
                                </div>
                            </template>
                            <template v-else>
                                <div v-for="(colId, headerText) in filteredMappings" :key="headerText" style="display: grid; grid-template-columns: 1fr 1fr auto; gap: 0.5rem; align-items: center;" class="bg-white/5 p-2 rounded">
                                    <input type="text" :value="headerText" @change="updateMappingHeader(headerText, $event.target.value)" class="w-full h-9 bg-slate-900 border border-slate-700 rounded px-3 text-slate-100 focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 transition-all text-sm" />
                                    
                                    <input v-if="activeMappingType === 'footer_label_mappings'" type="text" :value="colId" @change="updateMappingColId(headerText, $event.target.value)" class="w-full h-9 bg-slate-900 border border-slate-700 rounded px-3 text-slate-100 focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 transition-all text-sm" :disabled="true" />
                                    
                                    <select v-else :value="colId" @change="updateMappingColId(headerText, $event.target.value)" class="w-full h-9 bg-slate-900 border border-slate-700 rounded px-3 text-slate-100 focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 transition-all text-sm">
                                        <option v-for="opt in systemOptions" :value="opt.id">
                                            {{ opt.label }} ({{ opt.id }})
                                        </option>
                                        <option v-if="!systemOptions.find(o => o.id === colId)" :value="colId">{{ colId }} (Unknown)</option>
                                    </select>
                                    
                                    <button class="h-9 px-4 bg-red-500/80 hover:bg-red-500 text-white rounded cursor-pointer transition-colors w-full text-sm font-medium" @click="deleteMapping(headerText)">Delete</button>
                                </div>
                            </template>
                            <div v-if="Object.keys(filteredMappings).length === 0" class="p-4 text-center text-secondary">
                                No mappings found matching your search.
                            </div>
                        </div>
                    </div>
                    
                    <div class="mt-4 flex justify-end">
                        <button class="px-6 py-3 bg-emerald-500 hover:bg-emerald-400 text-white font-bold rounded-full shadow-lg shadow-emerald-500/20 transition-all transform hover:-translate-y-0.5 disabled:opacity-50 disabled:cursor-not-allowed" @click="saveMappings" :disabled="isSavingMappings">
                            {{ isSavingMappings ? 'Saving...' : 'Save Mappings' }}
                        </button>
                    </div>
                    
                    <div v-if="mappingStatusMessage" :class="['status-box', mappingStatusType]" class="mt-4">
                        {{ mappingStatusMessage }}
                    </div>
                </div>
            </div>
        </div>
    `,
    setup() {
        const currentStep = ref(1);
        const selectedFiles = ref([]);
        const singleFileSuffix = ref("KH");
        const isProcessing = ref(false);
        const statusMessage = ref("");
        const statusType = ref("info");
        const ignoreMissingDescription = ref(false);

        const showMappings = ref(false);
        const globalMappings = ref({});
        const mappingSearch = ref("");
        const isSavingMappings = ref(false);
        const mappingStatusMessage = ref("");
        const mappingStatusType = ref("info");
        const activeMappingType = ref("header_text_mappings");
        const newMappingKey = ref("");
        const newMappingVal = ref("");
        const newMappingType = ref("aggregation");
        const footerKeywords = ref([]);

        // Data
        const fileTokens = ref([]); // Array of { filename, missingHeaders }
        const allMissingHeaders = ref([]); // Deduplicated list across all files
        const allMissingFooters = ref([]); // Deduplicated footers
        const filePrefix = ref("");
        const userMappings = reactive({});
        const confirmedHeaders = ref([]);
        const confirmedFooters = ref([]);
        const proactiveWarnings = ref([]); // warnings from analysis
        const bundlePath = ref("");
        const generatedPrefixes = ref([]);
        const pricingMode = ref('standard'); // 'standard' or 'net'

        const systemOptions = ref([]);

        /**
         * Returns true when user uploaded 2 files (KH + VN mode).
         */
        const isDualMode = computed(() => selectedFiles.value.length >= 2);

        // Load options on mount
        const fetchOptions = async () => {
            try {
                const res = await fetch('/api/blueprint/options');
                if (res.ok) {
                    systemOptions.value = await res.json();
                }
            } catch (e) {
                console.error("Failed to fetch options", e);
            }
        };
        fetchOptions();

        const fetchMappings = async () => {
            try {
                const res = await fetch(`/api/blueprint/mappings?mapping_type=${activeMappingType.value}`);
                if (res.ok) {
                    globalMappings.value = await res.json();
                }
            } catch (e) {
                console.error("Failed to fetch mappings", e);
            }
        };
        fetchMappings();

        const fetchFooterMappings = async () => {
            try {
                const res = await fetch(`/api/blueprint/mappings?mapping_type=footer_label_mappings`);
                if (res.ok) {
                    const data = await res.json();
                    footerKeywords.value = Object.keys(data).map(k => k.toUpperCase());
                }
            } catch (e) {
                console.error("Failed to fetch footer mappings", e);
            }
        };
        fetchFooterMappings();

        const filteredMappings = computed(() => {
            if (!mappingSearch.value) return globalMappings.value;
            const term = mappingSearch.value.toLowerCase();
            const result = {};
            for (const [key, val] of Object.entries(globalMappings.value)) {
                if (key.toLowerCase().includes(term) || (typeof val === 'string' && val.toLowerCase().includes(term))) {
                    result[key] = val;
                }
            }
            return result;
        });

        const updateMappingHeader = (oldKey, newKey) => {
            const trimmed = newKey.trim();
            if (oldKey === trimmed || !trimmed) return;
            if (globalMappings.value[trimmed]) {
                alert("Header mapping already exists.");
                return;
            }

            const newMappings = { ...globalMappings.value };
            newMappings[trimmed] = newMappings[oldKey];
            delete newMappings[oldKey];
            globalMappings.value = newMappings;
        };

        const updateMappingColId = (key, newColId) => {
            globalMappings.value = { ...globalMappings.value, [key]: newColId };
        };

        const deleteMapping = (key) => {
            if (confirm(`Are you sure you want to delete the mapping for "${key}"?`)) {
                const newMappings = { ...globalMappings.value };
                delete newMappings[key];
                globalMappings.value = newMappings;
            }
        };

        const addNewMapping = () => {
            if (newMappingKey.value && newMappingVal.value) {
                globalMappings.value = {
                    ...globalMappings.value,
                    [newMappingKey.value]: newMappingVal.value
                };
                newMappingKey.value = "";
                if (activeMappingType.value === 'sheet_mappings') {
                    newMappingVal.value = "";
                } else {
                    newMappingVal.value = activeMappingType.value === 'footer_label_mappings' ? 'Footer Keyword' : '';
                }
            }
        };

        const switchMappingType = async (type) => {
            activeMappingType.value = type;
            await fetchMappings();
            mappingStatusMessage.value = "";
            newMappingKey.value = "";
            
            if (type === 'sheet_mappings') {
                newMappingVal.value = "";
            } else {
                newMappingVal.value = type === 'footer_label_mappings' ? 'Footer Keyword' : '';
            }
            
            if (type === 'footer_label_mappings') {
                footerKeywords.value = Object.keys(globalMappings.value).map(k => k.toUpperCase());
            }
        };

        const saveMappings = async () => {
            isSavingMappings.value = true;
            mappingStatusMessage.value = "Saving mappings...";
            mappingStatusType.value = "info";
            try {
                const res = await fetch('/api/blueprint/mappings', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        mapping_type: activeMappingType.value,
                        mappings: globalMappings.value
                    })
                });
                if (res.ok) {
                    mappingStatusMessage.value = "Mappings saved successfully!";
                    mappingStatusType.value = "success";
                    
                    if (activeMappingType.value === 'footer_label_mappings') {
                        footerKeywords.value = Object.keys(globalMappings.value).map(k => k.toUpperCase());
                    }
                    
                    setTimeout(() => { mappingStatusMessage.value = ""; }, 3000);
                } else {
                    const data = await res.json();
                    throw new Error(data.error || "Save failed");
                }
            } catch (e) {
                mappingStatusType.value = "error";
                mappingStatusMessage.value = e.message;
            } finally {
                isSavingMappings.value = false;
            }
        };

        /**
         * Handles file input change. Accepts up to 2 files.
         */
        const handleFileUpload = (e) => {
            const files = Array.from(e.target.files).slice(0, 2);
            selectedFiles.value = files;
            singleFileSuffix.value = "KH";
            statusMessage.value = "";
        };

        const toggleMapping = (headerText) => {
            if (confirmedHeaders.value.includes(headerText)) {
                confirmedHeaders.value = confirmedHeaders.value.filter(h => h !== headerText);
            } else {
                if (!userMappings[headerText]) {
                    alert("Please select a field first.");
                    return;
                }
                confirmedHeaders.value.push(headerText);
            }
        };

        const toggleFooter = (footerText) => {
            if (confirmedFooters.value.includes(footerText)) {
                confirmedFooters.value = confirmedFooters.value.filter(f => f !== footerText);
            } else {
                confirmedFooters.value.push(footerText);
            }
        };

        const analyzeFiles = async () => {
            if (selectedFiles.value.length === 0) return;
            isProcessing.value = true;
            statusMessage.value = "Scanning template structure...";
            allMissingHeaders.value = [];
            allMissingFooters.value = [];
            fileTokens.value = [];

            // Refresh footer mappings before analysis to prevent stale keywords
            await fetchFooterMappings();

            try {
                const headerSet = new Set();
                const footerSet = new Set();
                const warningSet = new Set();

                for (const file of selectedFiles.value) {
                    const formData = new FormData();
                    formData.append('file', file);

                    const res = await fetch(`/api/template/analyze?ignore_missing_description=${ignoreMissingDescription.value}`, { method: 'POST', body: formData });
                    const data = await res.json();

                    if (!res.ok) {
                        throw new Error(data.error || `Analysis failed for ${file.name}`);
                    }

                    // Collect file token info
                    fileTokens.value.push({
                        filename: data.temp_filename,
                        missingHeaders: (data.missing_headers || []).map(h => h.text)
                    });

                    // Collect unique missing headers across all files
                    for (const h of (data.missing_headers || [])) {
                        headerSet.add(h.text);
                    }
                    
                    // Collect missing footers with case-insensitive deduplication
                    for (const f of (data.missing_footers || [])) {
                        const fUpper = f.toUpperCase().trim();
                        // Only add if not already in global mappings
                        if (!footerKeywords.value.includes(fUpper)) {
                            // Check if already in our current set (case-insensitive)
                            const exists = Array.from(footerSet).some(existing => existing.toUpperCase().trim() === fUpper);
                            if (!exists) {
                                footerSet.add(f);
                            }
                        }
                    }

                    // Collect proactive warnings
                    if (data.warnings && data.warnings.length > 0) {
                        data.warnings.forEach(w => warningSet.add(w));
                    }
                }

                allMissingHeaders.value = Array.from(headerSet);
                allMissingFooters.value = Array.from(footerSet);
                proactiveWarnings.value = Array.from(warningSet);

                if (allMissingHeaders.value.length > 0 || allMissingFooters.value.length > 0) {
                    statusMessage.value = "Unmapped fields found. Please review.";
                } else if (proactiveWarnings.value.length > 0) {
                    statusMessage.value = "Template analyzed with warnings.";
                } else {
                    statusMessage.value = "Structure looks clean!";
                }

                // Suggest prefix from first filename
                filePrefix.value = selectedFiles.value[0].name.split('.')[0];
                currentStep.value = 2;

            } catch (e) {
                statusType.value = "error";
                statusMessage.value = e.message;
            } finally {
                isProcessing.value = false;
            }
        };

        /**
         * Generate template(s).
         * Single file: 1 API call. Dual files: 2 API calls into same bundle_dir_name.
         */
        const generateTemplate = async () => {
            if (!filePrefix.value) {
                alert("Please enter a prefix");
                return;
            }
            isProcessing.value = true;
            statusMessage.value = "Generating bundle configuration...";
            statusType.value = "info";
            generatedPrefixes.value = [];

            try {
                // Collect confirmed mappings
                const finalMappings = {};
                for (const [key, value] of Object.entries(userMappings)) {
                    if (confirmedHeaders.value.includes(key)) {
                        finalMappings[key] = value;
                    }
                }

                if (isDualMode.value) {
                    // --- DUAL MODE: 2 files → KH + VN ---
                    const suffixes = ['KH', 'VN'];
                    const baseName = filePrefix.value;

                    for (let i = 0; i < Math.min(fileTokens.value.length, 2); i++) {
                        const localeVal = suffixes[i];
                        const displayPrefix = `${baseName}_${localeVal}`;
                        statusMessage.value = `Generating ${displayPrefix}...`;

                        const res = await fetch('/api/template/generate', {
                            method: 'POST',
                            headers: { 'Content-Type': 'application/json' },
                            body: JSON.stringify({
                                customer_code: baseName,
                                locale: localeVal,
                                user_mappings: finalMappings,
                                temp_filename: fileTokens.value[i].filename,
                                bundle_dir_name: baseName,
                                confirmed_footers: confirmedFooters.value,
                                pricing_mode: pricingMode.value,
                                ignore_missing_description: ignoreMissingDescription.value
                            })
                        });
                        const data = await res.json();

                        if (!res.ok) {
                            throw new Error(data.error || `Generation failed for ${displayPrefix}`);
                        }

                        generatedPrefixes.value.push(displayPrefix);
                        bundlePath.value = data.bundle_path || '';
                    }

                    currentStep.value = 3;

                } else {
                    // --- SINGLE MODE: 1 file ---
                    const baseName = filePrefix.value;
                    const localeVal = singleFileSuffix.value;
                    const displayPrefix = `${baseName}_${localeVal}`;

                    const res = await fetch('/api/template/generate', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({
                            customer_code: baseName,
                            locale: localeVal,
                            user_mappings: finalMappings,
                            temp_filename: fileTokens.value[0].filename,
                            bundle_dir_name: baseName,
                            confirmed_footers: confirmedFooters.value,
                            pricing_mode: pricingMode.value,
                            ignore_missing_description: ignoreMissingDescription.value
                        })
                    });
                    const data = await res.json();

                    if (!res.ok) {
                        throw new Error(data.error || "Generation failed");
                    }

                    generatedPrefixes.value.push(displayPrefix);
                    bundlePath.value = data.bundle_path || '';
                    currentStep.value = 3;
                }

            } catch (e) {
                statusType.value = "error";
                statusMessage.value = e.message;
            } finally {
                isProcessing.value = false;
            }
        };

        const resetFlow = () => {
            currentStep.value = 1;
            selectedFiles.value = [];
            singleFileSuffix.value = "KH";
            filePrefix.value = "";
            allMissingHeaders.value = [];
            statusMessage.value = "";
            bundlePath.value = "";
            generatedPrefixes.value = [];
            fileTokens.value = [];
            for (const prop of Object.getOwnPropertyNames(userMappings)) {
                delete userMappings[prop];
            }
            confirmedHeaders.value = [];
            confirmedFooters.value = [];
            allMissingFooters.value = [];
            proactiveWarnings.value = [];
            ignoreMissingDescription.value = false;
        };

        const forceAnalyze = async () => {
            ignoreMissingDescription.value = true;
            statusMessage.value = "";
            statusType.value = "info";
            await analyzeFiles();
        };

        return {
            currentStep,
            selectedFiles,
            singleFileSuffix,
            isDualMode,
            isProcessing,
            statusMessage,
            statusType,
            ignoreMissingDescription,
            forceAnalyze,
            handleFileUpload,
            analyzeFiles,
            generateTemplate,
            resetFlow,
            filePrefix,
            allMissingHeaders,
            allMissingFooters,
            userMappings,
            confirmedHeaders,
            confirmedFooters,
            toggleMapping,
            toggleFooter,
            systemOptions,
            bundlePath,
            generatedPrefixes,
            showMappings,
            globalMappings,
            mappingSearch,
            isSavingMappings,
            mappingStatusMessage,
            mappingStatusType,
            filteredMappings,
            updateMappingHeader,
            updateMappingColId,
            deleteMapping,
            saveMappings,
            activeMappingType,
            switchMappingType,
            newMappingKey,
            newMappingVal,
            newMappingType,
            addNewMapping,
            proactiveWarnings,
            pricingMode
        };
    }
};
