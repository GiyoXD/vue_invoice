import { ref, reactive } from 'vue';

const SYSTEM_HEADERS = [
    "col_po", "col_item", "col_desc", "col_qty_pcs", "col_qty_sf",
    "col_unit_price", "col_amount", "col_net", "col_gross", "col_cbm",
    "col_pallet", "col_remarks", "col_static", "col_dc"
];

export default {
    template: `
        <div class="template-extractor-view fade-in">
            <h1>New Template Extractor</h1>
            
            <!-- STEP 1: UPLOAD -->
            <div class="card" v-if="currentStep === 1">
                <h2>1. Analyze Invoice Source</h2>
                <p style="color: #94a3b8; margin-bottom: 1rem;">
                    Upload a sample invoice file. The system will detect unmapped headers.
                </p>
                
                <input type="file" @change="handleFileUpload" accept=".xlsx, .xls" />
                
                <button class="btn" @click="analyzeFile" :disabled="!selectedFile || isProcessing">
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
                </div>

                <div v-if="missingHeaders.length === 0" class="status-box success">
                    ✅ All headers recognized automatically!
                </div>

                <div v-else class="mapping-grid" style="display: grid; gap: 1rem; margin-top: 1rem;">
                    <div v-for="(header, index) in missingHeaders" :key="index" style="background: rgba(255,255,255,0.03); padding: 1rem; border-radius: 6px;">
                        <div style="font-weight: bold; margin-bottom: 0.5rem; color: #fbbf24;">"{{ header.text }}"</div>
                        <div style="display: flex; gap: 0.5rem; align-items: center;">
                            <select v-model="userMappings[header.text]" class="input-field" style="width: 100%;" :disabled="confirmedHeaders.includes(header.text)">
                                <option value="" disabled selected>Select a field...</option>
                                <option v-for="opt in systemHeaders" :value="opt">{{ opt }}</option>
                            </select>
                            <button 
                                class="btn-sm" 
                                :class="confirmedHeaders.includes(header.text) ? 'btn-danger' : 'btn-success'"
                                @click="toggleMapping(header.text)"
                                style="font-size: 0.8rem; padding: 0.3rem 0.6rem; min-width: 60px;">
                                {{ confirmedHeaders.includes(header.text) ? 'Remove' : 'Add' }}
                            </button>
                        </div>
                        <div style="font-size: 0.8rem; opacity: 0.6; margin-top: 4px;">
                            Suggested: {{ header.suggestion }}
                        </div>
                    </div>
                </div>

                <div class="flex-row" style="margin-top: 2rem; gap: 1rem; display: flex;">
                    <button class="nav-btn" @click="currentStep = 1">Back</button>
                    <button class="btn" @click="generateTemplate" :disabled="isProcessing || !filePrefix">
                        {{ isProcessing ? 'Generating...' : 'Create Template' }}
                    </button>
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
                </div>
                <p style="color: #94a3b8; margin-bottom: 2rem;">
                    You can now go to the Generator and process invoices for this company.
                </p>
                <button class="btn" @click="resetFlow">Process Another</button>
            </div>
        </div>
    `,
    setup() {
        const currentStep = ref(1);
        const selectedFile = ref(null);
        const isProcessing = ref(false);
        const statusMessage = ref("");
        const statusType = ref("info");

        // Data
        const tempFilename = ref("");
        const missingHeaders = ref([]);
        const filePrefix = ref("");
        const userMappings = reactive({});
        const confirmedHeaders = ref([]);
        const bundlePath = ref("");

        const handleFileUpload = (e) => {
            selectedFile.value = e.target.files[0];
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

        const analyzeFile = async () => {
            if (!selectedFile.value) return;
            isProcessing.value = true;
            statusMessage.value = "Uploading and analyzing structure...";

            const formData = new FormData();
            formData.append('file', selectedFile.value);

            try {
                const res = await fetch('/api/template/analyze', { method: 'POST', body: formData });
                const data = await res.json();

                if (res.ok) {
                    tempFilename.value = data.temp_filename;
                    missingHeaders.value = data.missing_headers;
                    filePrefix.value = data.suggested_prefix;

                    // Pre-fill mappings with suggestions or unknown
                    missingHeaders.value.forEach(h => {
                        // Default to suggestion if in list, else col_item? No, let user choose.
                        // Or try to match suggestion.
                        userMappings[h.text] = SYSTEM_HEADERS.includes(h.suggestion) ? h.suggestion : '';
                    });

                    currentStep.value = 2;
                    statusMessage.value = "";
                } else {
                    throw new Error(data.error || "Analysis failed");
                }
            } catch (e) {
                statusType.value = "error";
                statusMessage.value = e.message;
            } finally {
                isProcessing.value = false;
            }
        };

        const generateTemplate = async () => {
            if (!filePrefix.value) {
                alert("Please enter a prefix");
                return;
            }
            isProcessing.value = true;
            statusMessage.value = "Generating bundle configuration...";
            statusType.value = "info";

            try {

                // Filter out 'col_unknown' mappings to avoid polluting the config
                // AND ensure only explicitly confirmed headers are sent
                const filteredMappings = {};
                for (const [key, value] of Object.entries(userMappings)) {
                    if (value && value !== 'col_unknown' && confirmedHeaders.value.includes(key)) {
                        filteredMappings[key] = value;
                    }
                }

                const res = await fetch('/api/template/generate', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        file_prefix: filePrefix.value,
                        user_mappings: filteredMappings,
                        temp_filename: tempFilename.value
                    })
                });
                const data = await res.json();

                if (res.ok) {
                    bundlePath.value = data.bundle_path || '';
                    currentStep.value = 3;
                } else {
                    throw new Error(data.error || "Generation failed");
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
            selectedFile.value = null;
            filePrefix.value = "";
            missingHeaders.value = [];
            statusMessage.value = "";
            bundlePath.value = "";
        };

        return {
            currentStep,
            selectedFile,
            isProcessing,
            statusMessage,
            statusType,
            handleFileUpload,
            analyzeFile,
            generateTemplate,
            resetFlow,
            filePrefix,
            missingHeaders,
            userMappings,
            confirmedHeaders,
            toggleMapping,
            systemHeaders: SYSTEM_HEADERS,
            bundlePath
        };
    }
};
