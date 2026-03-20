import { ref, reactive, computed } from 'vue';

export default {
    template: `
        <div class="template-extractor-view fade-in">
            <h1>New Template Extractor</h1>
            
            <!-- STEP 1: UPLOAD -->
            <div class="card" v-if="currentStep === 1">
                <h2>1. Analyze Invoice Source</h2>
                <p style="color: #94a3b8; margin-bottom: 1rem;">
                    Upload a sample invoice file. Upload <strong>2 files</strong> to auto-create KH + VN versions.
                </p>
                
                <input type="file" @change="handleFileUpload" accept=".xlsx, .xls" multiple />
                
                <!-- Show selected files with KH/VN labels -->
                <div v-if="selectedFiles.length > 0" style="margin-top: 1rem;">
                    <div v-for="(file, idx) in selectedFiles" :key="idx" 
                         style="display: flex; align-items: center; gap: 0.75rem; padding: 0.5rem 0.75rem; margin-bottom: 0.5rem; background: rgba(255,255,255,0.03); border-radius: 6px; border: 1px solid rgba(255,255,255,0.08);">
                        <span v-if="selectedFiles.length === 2" 
                              :style="{
                                  padding: '0.2rem 0.6rem',
                                  borderRadius: '4px',
                                  fontSize: '0.75rem',
                                  fontWeight: 'bold',
                                  background: idx === 0 ? 'rgba(59, 130, 246, 0.2)' : 'rgba(234, 179, 8, 0.2)',
                                  color: idx === 0 ? '#60a5fa' : '#facc15',
                                  border: idx === 0 ? '1px solid rgba(59, 130, 246, 0.3)' : '1px solid rgba(234, 179, 8, 0.3)'
                              }">
                            {{ idx === 0 ? 'KH' : 'VN' }}
                        </span>
                        <span style="color: #e2e8f0;">📄 {{ file.name }}</span>
                    </div>
                    
                    <!-- Single-file suffix selector -->
                    <div v-if="selectedFiles.length === 1" style="margin-top: 0.75rem; display: flex; align-items: center; gap: 0.75rem;">
                        <label style="color: #94a3b8; font-size: 0.875rem;">Version suffix:</label>
                        <select v-model="singleFileSuffix" class="input-field" style="width: 160px;">
                            <option value="_KH">KH version</option>
                            <option value="_VN">VN version</option>
                        </select>
                    </div>

                    <div v-if="selectedFiles.length > 2" class="status-box error" style="margin-top: 0.5rem;">
                        ⚠️ Maximum 2 files allowed. Only the first 2 will be used.
                    </div>
                </div>
                
                <button class="btn" @click="analyzeFiles" :disabled="selectedFiles.length === 0 || isProcessing" style="margin-top: 1rem;">
                    {{ isProcessing ? 'Analyzing...' : 'Analyze & Extract' }}
                </button>
                
                 <div v-if="statusMessage" :class="['status-box', statusType]">
                    {{ statusMessage }}
                </div>
            </div>

            <!-- STEP 2: MAP HEADERS -->
            <div class="card" v-if="currentStep === 2" style="animation-delay: 0.1s">
                <h2>2. Map Unrecognized Headers</h2>
                <p style="color: #94a3b8; margin-bottom: 1.5rem;">
                    We found some headers we don't recognize. Please map them to system fields.
                </p>

                <div class="form-group">
                    <label>Template Prefix (Unique ID)</label>
                    <input type="text" v-model="filePrefix" class="input-field" placeholder="e.g. MOTO, JLFHM" />
                    
                    <!-- Show preview of what will be created -->
                    <div v-if="filePrefix && isDualMode" style="margin-top: 0.5rem; padding: 0.5rem 0.75rem; background: rgba(59, 130, 246, 0.08); border: 1px solid rgba(59, 130, 246, 0.2); border-radius: 6px; font-size: 0.85rem; color: #93c5fd;">
                        📁 Will create: <strong>{{ filePrefix }}_KH</strong> + <strong>{{ filePrefix }}_VN</strong> in <code>bundled/{{ filePrefix }}/</code>
                    </div>
                    <div v-else-if="filePrefix && singleFileSuffix" style="margin-top: 0.5rem; padding: 0.5rem 0.75rem; background: rgba(59, 130, 246, 0.08); border: 1px solid rgba(59, 130, 246, 0.2); border-radius: 6px; font-size: 0.85rem; color: #93c5fd;">
                        📁 Will create: <strong>{{ filePrefix }}{{ singleFileSuffix }}</strong> in <code>bundled/{{ filePrefix }}/</code>
                    </div>
                </div>

                <div v-if="allMissingHeaders.length === 0" class="status-box success">
                    ✅ All headers recognized automatically!
                </div>

                <div v-else class="mapping-grid" style="display: grid; gap: 1rem; margin-top: 1rem;">
                    <div v-for="(headerText, index) in allMissingHeaders" :key="index" style="background: rgba(255,255,255,0.03); padding: 1rem; border-radius: 6px;">
                        <div style="font-weight: bold; margin-bottom: 0.5rem; color: #fbbf24;">"{{ headerText }}"</div>
                        <div style="display: flex; gap: 0.5rem; align-items: center;">
                            <select v-model="userMappings[headerText]" class="input-field" style="width: 100%;" :disabled="confirmedHeaders.includes(headerText)">
                                <option value="" disabled selected>Select a field...</option>
                                <option v-for="opt in systemOptions" :value="opt.id">
                                    {{ opt.label }} ({{ opt.id }})
                                </option>
                            </select>
                            <button 
                                class="btn-sm" 
                                :class="confirmedHeaders.includes(headerText) ? 'btn-danger' : 'btn-success'"
                                @click="toggleMapping(headerText)"
                                style="font-size: 0.8rem; padding: 0.3rem 0.6rem; min-width: 60px;">
                                {{ confirmedHeaders.includes(headerText) ? 'Remove' : 'Add' }}
                            </button>
                        </div>
                    </div>
                </div>

                <div class="flex-row" style="margin-top: 2rem; gap: 1rem; display: flex;">
                    <button class="nav-btn" @click="currentStep = 1">Back</button>
                    <button class="btn" @click="generateTemplate" :disabled="isProcessing || !filePrefix">
                        {{ isProcessing ? 'Generating...' : 'Create Template' }}
                    </button>
                </div>
                 
                <!-- PROACTIVE WARNINGS PANEL -->
                <div v-if="proactiveWarnings && proactiveWarnings.length > 0" class="warning-panel" style="margin-bottom: 1.5rem;">
                    <div class="warning-header" style="display: flex; align-items: center; gap: 0.5rem; margin-bottom: 0.75rem;">
                        <span class="warning-icon" style="font-size: 1.25rem;">⚠️</span>
                        <h3 style="margin: 0; color: #b45309; font-size: 1rem;">Template Structural Warnings</h3>
                    </div>
                    <ul style="margin: 0; padding-left: 1.5rem; color: #92400e; font-size: 0.9rem;">
                        <li v-for="(msg, idx) in proactiveWarnings" :key="idx" style="margin-bottom: 0.5rem; line-height: 1.4;">
                            {{ msg }}
                        </li>
                    </ul>
                </div>

                <div v-if="statusMessage" :class="['status-box', statusType]" style="margin-top: 1rem;">
                    {{ statusMessage }}
                </div>
            </div>

            <!-- STEP 3: SUCCESS -->
            <div class="card" v-if="currentStep === 3" style="text-align: center; animation-delay: 0.1s">
                <div style="font-size: 4rem; margin-bottom: 1rem;">🎉</div>
                <h2>Template Created!</h2>
                <p style="color: #94a3b8; margin-bottom: 1rem;">
                    The template <strong>{{ filePrefix }}</strong> has been configured successfully.
                </p>
                <div v-if="bundlePath" style="background: rgba(34, 197, 94, 0.1); padding: 1rem; border-radius: 0.5rem; margin-bottom: 1.5rem; text-align: left;">
                    <p style="color: #86efac; margin: 0 0 0.5rem 0; font-size: 0.875rem;">📁 Bundle created at:</p>
                    <code style="color: #22c55e; font-size: 0.8rem; word-break: break-all;">{{ bundlePath }}</code>
                    <div v-if="generatedPrefixes.length > 1" style="margin-top: 0.5rem; padding-top: 0.5rem; border-top: 1px solid rgba(34, 197, 94, 0.2);">
                        <p style="color: #86efac; margin: 0 0 0.25rem 0; font-size: 0.8rem;">Contains:</p>
                        <div v-for="p in generatedPrefixes" :key="p" style="color: #4ade80; font-size: 0.8rem;">
                            ✅ {{ p }}
                        </div>
                    </div>
                </div>
                <p style="color: #94a3b8; margin-bottom: 2rem;">
                    You can now go to the Generator and process invoices for this company.
                </p>
                <button class="btn" @click="resetFlow">Process Another</button>
            </div>

            <!-- GLOBAL MAPPINGS -->
            <div class="card" style="margin-top: 2rem;">
                <div style="display: flex; justify-content: space-between; align-items: center; cursor: pointer;" @click="showMappings = !showMappings">
                    <h2>Manage Global Mappings</h2>
                    <span>{{ showMappings ? '▲ Collapse' : '▼ Expand' }}</span>
                </div>
                
                <div v-if="showMappings" style="margin-top: 1rem;">
                    <p style="color: #94a3b8; margin-bottom: 1rem;">
                        View and edit the globally recognized mappings. These are used to automatically match headers and sheets in templates.
                    </p>
                    
                    <div style="display: flex; gap: 0.5rem; margin-bottom: 1rem;">
                        <select v-model="activeMappingType" @change="switchMappingType($event.target.value)" class="input-field" style="width: 280px; font-weight: bold;">
                            <option value="header_text_mappings">Header Mappings</option>
                            <option value="sheet_name_mappings">Sheet Name Mappings</option>
                            <option value="shipping_header_map">Shipping Header Map</option>
                        </select>
                        <input type="text" v-model="mappingSearch" class="input-field" placeholder="Search..." style="flex: 1;" />
                    </div>

                    <!-- Add New Mapping Row -->
                    <div style="display: grid; grid-template-columns: 1fr 1fr auto; gap: 0.5rem; margin-bottom: 1rem; padding: 0.5rem; background: rgba(34, 197, 94, 0.05); border: 1px dashed #22c55e; border-radius: 6px; align-items: center;">
                        <input type="text" v-model="newMappingKey" class="input-field" :placeholder="activeMappingType === 'shipping_header_map' ? 'Col ID (e.g. col_grade)' : 'New Input Text (e.g. Qty(SF))'" style="padding: 0.5rem;" />
                        
                        <input v-if="activeMappingType === 'sheet_name_mappings' || activeMappingType === 'shipping_header_map'" type="text" v-model="newMappingVal" class="input-field" :placeholder="activeMappingType === 'shipping_header_map' ? 'Keywords (comma-separated)' : 'Target Name (e.g. Packing list)'" style="padding: 0.5rem;" />
                        <select v-else v-model="newMappingVal" class="input-field" style="padding: 0.5rem;">
                            <option value="" disabled selected>Select system field...</option>
                            <option v-for="opt in systemOptions" :value="opt.id">{{ opt.label }} ({{ opt.id }})</option>
                        </select>
                        
                        <button class="btn" @click.prevent="addNewMapping" style="margin: 0; min-width: 80px;" :disabled="!newMappingKey || !newMappingVal">Add</button>
                    </div>

                    <div style="max-height: 400px; overflow-y: auto; border: 1px solid rgba(255,255,255,0.1); border-radius: 6px; padding: 0.5rem;">
                        <div class="mapping-grid" style="display: grid; gap: 0.5rem;">
                            <!-- Header Row -->
                            <div style="display: grid; grid-template-columns: 1fr 1fr auto; gap: 0.5rem; font-weight: bold; padding: 0.5rem; border-bottom: 1px solid rgba(255,255,255,0.1);">
                                <div>{{ activeMappingType === 'shipping_header_map' ? 'Column ID' : 'Original Text (Excel)' }}</div>
                                <div>{{ activeMappingType === 'shipping_header_map' ? 'Keywords (comma-separated)' : 'Mapped Target (System)' }}</div>
                                <div style="width: 70px; text-align: center;">Action</div>
                            </div>
                            
                            <div v-for="(colId, headerText) in filteredMappings" :key="headerText" style="display: grid; grid-template-columns: 1fr 1fr auto; gap: 0.5rem; align-items: center; background: rgba(255,255,255,0.03); padding: 0.5rem; border-radius: 4px;">
                                <input type="text" :value="headerText" @change="updateMappingHeader(headerText, $event.target.value)" class="input-field" style="padding: 0.3rem;" />
                                
                                <input v-if="activeMappingType === 'sheet_name_mappings' || activeMappingType === 'shipping_header_map'" type="text" :value="colId" @change="updateMappingColId(headerText, $event.target.value)" class="input-field" style="padding: 0.3rem;" />
                                
                                <select v-else :value="colId" @change="updateMappingColId(headerText, $event.target.value)" class="input-field" style="padding: 0.3rem;">
                                    <option v-for="opt in systemOptions" :value="opt.id">
                                        {{ opt.label }} ({{ opt.id }})
                                    </option>
                                    <option v-if="!systemOptions.find(o => o.id === colId)" :value="colId">{{ colId }} (Unknown)</option>
                                </select>
                                
                                <button class="btn-sm" @click="deleteMapping(headerText)" style="padding: 0.3rem; min-width: 70px; background-color: #ef4444; color: white; border: none; border-radius: 4px; cursor: pointer;">Delete</button>
                            </div>
                            <div v-if="Object.keys(filteredMappings).length === 0" style="padding: 1rem; text-align: center; color: #94a3b8;">
                                No mappings found matching your search.
                            </div>
                        </div>
                    </div>
                    
                    <div style="margin-top: 1rem; display: flex; justify-content: flex-end;">
                        <button class="btn" @click="saveMappings" :disabled="isSavingMappings" style="background-color: #22c55e;">
                            {{ isSavingMappings ? 'Saving...' : 'Save Mappings' }}
                        </button>
                    </div>
                    
                    <div v-if="mappingStatusMessage" :class="['status-box', mappingStatusType]" style="margin-top: 1rem;">
                        {{ mappingStatusMessage }}
                    </div>
                </div>
            </div>
        </div>
    `,
    setup() {
        const currentStep = ref(1);
        const selectedFiles = ref([]);
        const singleFileSuffix = ref("_KH");
        const isProcessing = ref(false);
        const statusMessage = ref("");
        const statusType = ref("info");

        const showMappings = ref(false);
        const globalMappings = ref({});
        const mappingSearch = ref("");
        const isSavingMappings = ref(false);
        const mappingStatusMessage = ref("");
        const mappingStatusType = ref("info");
        const activeMappingType = ref("header_text_mappings");
        const newMappingKey = ref("");
        const newMappingVal = ref("");

        // Data
        const fileTokens = ref([]); // Array of { filename, missingHeaders }
        const allMissingHeaders = ref([]); // Deduplicated list across all files
        const filePrefix = ref("");
        const userMappings = reactive({});
        const confirmedHeaders = ref([]);
        const proactiveWarnings = ref([]); // warnings from analysis
        const bundlePath = ref("");
        const generatedPrefixes = ref([]);

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

        const filteredMappings = computed(() => {
            if (!mappingSearch.value) return globalMappings.value;
            const term = mappingSearch.value.toLowerCase();
            const result = {};
            for (const [key, val] of Object.entries(globalMappings.value)) {
                if (key.toLowerCase().includes(term) || val.toLowerCase().includes(term)) {
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
                newMappingVal.value = "";
            }
        };

        const switchMappingType = async (type) => {
            activeMappingType.value = type;
            await fetchMappings();
            mappingStatusMessage.value = "";
            newMappingKey.value = "";
            newMappingVal.value = "";
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
            singleFileSuffix.value = "";
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

        const analyzeFiles = async () => {
            if (selectedFiles.value.length === 0) return;
            isProcessing.value = true;
            statusMessage.value = "Scanning template structure...";
            allMissingHeaders.value = [];
            fileTokens.value = [];

            try {
                const headerSet = new Set();
                const warningSet = new Set();

                for (const file of selectedFiles.value) {
                    const formData = new FormData();
                    formData.append('file', file);

                    const res = await fetch('/api/template/analyze', { method: 'POST', body: formData });
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

                    // Collect proactive warnings
                    if (data.warnings && data.warnings.length > 0) {
                        data.warnings.forEach(w => warningSet.add(w));
                    }
                }

                allMissingHeaders.value = Array.from(headerSet);
                proactiveWarnings.value = Array.from(warningSet);

                if (allMissingHeaders.value.length > 0) {
                    statusMessage.value = "Unknown headers found.";
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
                    const suffixes = ['_KH', '_VN'];
                    const baseName = filePrefix.value;

                    for (let i = 0; i < Math.min(fileTokens.value.length, 2); i++) {
                        const suffixedPrefix = `${baseName}${suffixes[i]}`;
                        statusMessage.value = `Generating ${suffixedPrefix}...`;

                        const res = await fetch('/api/template/generate', {
                            method: 'POST',
                            headers: { 'Content-Type': 'application/json' },
                            body: JSON.stringify({
                                file_prefix: suffixedPrefix,
                                user_mappings: finalMappings,
                                temp_filename: fileTokens.value[i].filename,
                                bundle_dir_name: baseName
                            })
                        });
                        const data = await res.json();

                        if (!res.ok) {
                            throw new Error(data.error || `Generation failed for ${suffixedPrefix}`);
                        }

                        generatedPrefixes.value.push(suffixedPrefix);
                        bundlePath.value = data.bundle_path || '';
                    }

                    currentStep.value = 3;

                } else {
                    // --- SINGLE MODE: 1 file ---
                    const effectivePrefix = `${filePrefix.value}${singleFileSuffix.value}`;
                    const useBundleDir = singleFileSuffix.value ? filePrefix.value : "";

                    const res = await fetch('/api/template/generate', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({
                            file_prefix: effectivePrefix,
                            user_mappings: finalMappings,
                            temp_filename: fileTokens.value[0].filename,
                            bundle_dir_name: useBundleDir
                        })
                    });
                    const data = await res.json();

                    if (!res.ok) {
                        throw new Error(data.error || "Generation failed");
                    }

                    generatedPrefixes.value.push(effectivePrefix);
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
            singleFileSuffix.value = "";
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
            proactiveWarnings.value = [];
        };

        return {
            currentStep,
            selectedFiles,
            singleFileSuffix,
            isDualMode,
            isProcessing,
            statusMessage,
            statusType,
            handleFileUpload,
            analyzeFiles,
            generateTemplate,
            resetFlow,
            filePrefix,
            allMissingHeaders,
            userMappings,
            confirmedHeaders,
            toggleMapping,
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
            addNewMapping,
            proactiveWarnings
        };
    }
};
