import { ref, computed } from 'vue';

export default {
    emits: ['switch-view'], // Declare event to switch tabs
    template: `
        <div class="generator-view fade-in">
            <h1>Invoice Generator</h1>
            
            <div class="card">
                <h2>1. Upload Source Data</h2>
                <p style="color: #94a3b8; margin-bottom: 1rem;">Select your Excel file to begin processing.</p>
                
                <input type="file" @change="handleFileUpload" accept=".xlsx, .xls" />
                
                <button class="btn" @click="uploadFile" :disabled="!selectedFile || isUploading">
                    {{ isUploading ? 'Processing...' : 'Upload & Process' }}
                </button>

                <div v-if="uploadStatus && !uploadError" :class="['status-box', uploadStatus.type]">
                    {{ uploadStatus.message }}
                </div>

                <!-- ERROR PANEL FOR UPLOAD -->
                <div v-if="uploadError" class="error-panel">
                    <div class="error-header">
                        <span class="error-icon">⚠️</span>
                        <h3>Upload Failed</h3>
                    </div>
                    <span v-if="uploadError.step" class="error-step">{{ uploadError.step }}</span>
                    <div class="error-message">{{ uploadError.message }}</div>
                    
                    <div v-if="uploadError.traceback" 
                         class="traceback-toggle" 
                         :class="{ open: showUploadTraceback }"
                         @click="showUploadTraceback = !showUploadTraceback">
                        <span>📋 View Technical Details</span>
                        <span class="chevron">▼</span>
                    </div>
                    <div class="traceback-content" :class="{ open: showUploadTraceback }">
                        <pre>{{ uploadError.traceback }}</pre>
                    </div>
                    
                    <div class="error-actions">
                        <button class="btn-retry" @click="retryUpload">🔄 Try Again</button>
                        <button class="btn-copy-error" @click="copyError(uploadError)">📋 Copy Error</button>
                    </div>
                </div>
            </div>

            <div class="card" v-if="processingComplete" style="animation-delay: 0.2s">
                <h2>2. Invoice Details</h2>
                
                <!-- ASSET WARNING PANEL -->
                <div v-if="assetStatus && !assetStatus.ready" class="asset-warning">
                    <div class="warning-header">
                        <span class="warning-icon">📦</span>
                        <h3>Missing Blueprint Configuration</h3>
                    </div>
                    <div class="warning-message">{{ assetStatus.message }}</div>
                    <div class="warning-details">
                        <p><strong>Bundled Directory:</strong> <code>{{ assetStatus.bundled_dir }}</code></p>
                    </div>
                    <div class="warning-actions">
                        <button class="btn-create-template" @click="$emit('switch-view', 'extractor')">
                            ➕ Create New Template
                        </button>
                    </div>
                </div>
                
                <!-- ASSET READY STATUS -->
                <div v-if="assetStatus && assetStatus.ready" class="asset-ready">
                    <span class="ready-icon">✅</span>
                    <span class="ready-text">Blueprint found: using <strong>{{ assetConfigName }}</strong></span>
                </div>
                
                <div class="grid-form">
                    <div class="form-group">
                        <label>Invoice Number</label>
                        <input type="text" v-model="invoiceNo" class="input-field" />
                    </div>
                    <div class="form-group">
                        <label>Invoice Date</label>
                        <input type="date" v-model="invoiceDate" class="input-field" />
                    </div>
                    <div class="form-group">
                        <label>Invoice Ref (Optional)</label>
                        <input type="text" v-model="invoiceRef" class="input-field" />
                    </div>
                </div>

                <div class="form-group" style="margin-top: 1rem;">
                    <label>Generation Options</label>
                    <div style="display: flex; gap: 1.5rem; margin-top: 0.5rem; flex-wrap: wrap;">
                        <label style="display: flex; align-items: center; gap: 0.5rem; cursor: pointer;">
                            <input type="checkbox" v-model="includeStandard" accent-color="#2563eb" /> 
                            <span>Standard Invoice</span>
                        </label>
                        <label style="display: flex; align-items: center; gap: 0.5rem; cursor: pointer;">
                            <input type="checkbox" v-model="includeCustom" accent-color="#2563eb" /> 
                            <span>Custom Mode</span>
                        </label>
                        <label style="display: flex; align-items: center; gap: 0.5rem; cursor: pointer;">
                            <input type="checkbox" v-model="includeDAF" accent-color="#2563eb" /> 
                            <span>DAF Mode</span>
                        </label>
                    </div>
                </div>
                
                <button class="btn" @click="generateInvoice" :disabled="isGenerating || !assetStatus?.ready">
                    {{ isGenerating ? 'Generating...' : (assetStatus?.ready ? 'Generate Invoice' : 'Blueprint Required') }}
                </button>

                <div v-if="generationStatus && !generationError" :class="['status-box', generationStatus.type]">
                    {{ generationStatus.message }}
                </div>

                <!-- ERROR PANEL FOR GENERATION -->
                <div v-if="generationError" class="error-panel">
                    <div class="error-header">
                        <span class="error-icon">⚠️</span>
                        <h3>Generation Failed</h3>
                    </div>
                    <span v-if="generationError.step" class="error-step">{{ generationError.step }}</span>
                    <div class="error-message">{{ generationError.message }}</div>
                    
                    <div v-if="generationError.traceback" 
                         class="traceback-toggle" 
                         :class="{ open: showGenTraceback }"
                         @click="showGenTraceback = !showGenTraceback">
                        <span>📋 View Technical Details</span>
                        <span class="chevron">▼</span>
                    </div>
                    <div class="traceback-content" :class="{ open: showGenTraceback }">
                        <pre>{{ generationError.traceback }}</pre>
                    </div>
                    
                    <div class="error-actions">
                        <button class="btn-retry" @click="retryGeneration">🔄 Try Again</button>
                        <button class="btn-copy-error" @click="copyError(generationError)">📋 Copy Error</button>
                    </div>
                </div>
            </div>

            <!-- VALIDATION CARD -->
            <div class="card validation-card" v-if="validationData && !isGenerating && !generationError" style="animation-delay: 0.1s">
                <div class="validation-header">
                    <h3>✅ Invoice Generated Successfully</h3>
                    <span style="font-size: 0.875rem; opacity: 0.7">{{ validationData.timestamp }}</span>
                </div>

                <div v-if="summaryStats" class="stat-grid">
                    <div class="stat-item">
                        <span class="stat-label">Total Items</span>
                        <span class="stat-value">{{ summaryStats.total_pcs?.toLocaleString() || 0 }}</span>
                    </div>
                    <div class="stat-item">
                        <span class="stat-label">Total SQFT</span>
                        <span class="stat-value">{{ summaryStats.total_sqft?.toLocaleString(undefined, {maximumFractionDigits: 2}) || 0 }}</span>
                    </div>
                    <div class="stat-item">
                        <span class="stat-label">Total Pallets</span>
                        <span class="stat-value">{{ summaryStats.total_pallets || 0 }}</span>
                    </div>
                </div>

                <div v-if="weightStats" class="stat-grid">
                    <div class="stat-item">
                        <span class="stat-label">Net Weight</span>
                        <span class="stat-value">{{ weightStats.net?.toLocaleString() }} kg</span>
                    </div>
                    <div class="stat-item">
                        <span class="stat-label">Gross Weight</span>
                        <span class="stat-value">{{ weightStats.gross?.toLocaleString() }} kg</span>
                    </div>
                    <div class="stat-item">
                        <span class="stat-label">Total CBM</span>
                        <span class="stat-value">{{ weightStats.cbm?.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 3}) }} m³</span>
                    </div>
                </div>

                <div class="meta-details">
                    <p><strong>Config Used:</strong> {{ validationData.config_info?.data_source_detected || 'Unknown' }}</p>
                    <p><strong>Output Location:</strong></p>
                    <code>{{ validationData.output_dir_absolute }}</code>
                    
                    <button class="btn-small" @click="$emit('switch-view', 'inspector')" style="margin-top: 1rem; width: 100%;">🔍 Open in Data Inspector</button>
                </div>
            </div>
        </div>
    `,
    setup() {
        // --- Generator State ---
        const selectedFile = ref(null);
        const isUploading = ref(false);
        const uploadStatus = ref(null);
        const uploadError = ref(null);
        const showUploadTraceback = ref(false);

        const processingComplete = ref(false);
        const identifier = ref('');
        const jsonPath = ref('');

        const invoiceNo = ref('');
        const invoiceDate = ref(new Date().toISOString().split('T')[0]);
        const invoiceRef = ref('');

        // Options
        const includeStandard = ref(true);
        const includeCustom = ref(false);
        const includeDAF = ref(false);

        const isGenerating = ref(false);
        const generationStatus = ref(null);
        const generationError = ref(null);
        const showGenTraceback = ref(false);
        const validationData = ref(null); // Validation data from generation
        const assetStatus = ref(null); // Asset availability status from upload

        // --- Generator Actions ---
        const handleFileUpload = (event) => {
            selectedFile.value = event.target.files[0];
            uploadStatus.value = null;
            uploadError.value = null;
            showUploadTraceback.value = false;
            processingComplete.value = false;
            validationData.value = null;
            assetStatus.value = null;
        };

        /**
         * Uploads the selected file to the API and processes it.
         * Handles both success and error responses, populating the
         * appropriate state variables for UI display.
         */
        const uploadFile = async () => {
            if (!selectedFile.value) return;

            isUploading.value = true;
            uploadStatus.value = { type: 'info', message: 'Uploading and processing...' };
            uploadError.value = null;
            validationData.value = null;

            const formData = new FormData();
            formData.append('file', selectedFile.value);

            try {
                const response = await fetch('/api/upload', {
                    method: 'POST',
                    body: formData
                });

                const data = await response.json();

                if (response.ok) {
                    uploadStatus.value = { type: 'success', message: 'File processed successfully!' };
                    identifier.value = data.identifier;
                    jsonPath.value = data.json_path;
                    invoiceNo.value = data.default_inv_no || '';

                    // Capture asset status from API response
                    assetStatus.value = data.asset_status || null;

                    processingComplete.value = true;
                } else {
                    // Capture structured error from API
                    uploadError.value = {
                        message: data.error || 'Upload failed',
                        step: data.step || null,
                        traceback: data.traceback || null
                    };
                    uploadStatus.value = null;
                }
            } catch (error) {
                // Network/JS error
                uploadError.value = {
                    message: error.message || 'Network error occurred',
                    step: null,
                    traceback: null
                };
                uploadStatus.value = null;
            } finally {
                isUploading.value = false;
            }
        };

        /**
         * Triggers invoice generation with the provided metadata.
         * Handles both success and error responses.
         */
        const generateInvoice = async () => {
            isGenerating.value = true;
            generationStatus.value = { type: 'info', message: 'Generating invoice, please wait...' };
            generationError.value = null;
            validationData.value = null;

            try {
                const response = await fetch('/api/generate', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        identifier: identifier.value,
                        json_path: jsonPath.value,
                        invoice_no: invoiceNo.value,
                        invoice_date: invoiceDate.value,
                        invoice_no: invoiceNo.value,
                        invoice_date: invoiceDate.value,
                        invoice_ref: invoiceRef.value,
                        generate_standard: includeStandard.value,
                        generate_custom: includeCustom.value,
                        generate_daf: includeDAF.value
                    })
                });

                const data = await response.json();

                if (response.ok) {
                    generationStatus.value = { type: 'success', message: 'Invoice generated successfully!' };
                    if (data.metadata) {
                        validationData.value = data.metadata;
                    }
                } else {
                    // Capture structured error from API
                    generationError.value = {
                        message: data.error || 'Generation failed',
                        step: data.step || null,
                        traceback: data.traceback || null
                    };
                    generationStatus.value = null;
                }
            } catch (error) {
                // Network/JS error
                generationError.value = {
                    message: error.message || 'Network error occurred',
                    step: null,
                    traceback: null
                };
                generationStatus.value = null;
            } finally {
                isGenerating.value = false;
            }
        };

        /**
         * Retries the upload process after an error.
         */
        const retryUpload = () => {
            uploadError.value = null;
            showUploadTraceback.value = false;
            uploadFile();
        };

        /**
         * Retries the invoice generation after an error.
         */
        const retryGeneration = () => {
            generationError.value = null;
            showGenTraceback.value = false;
            generateInvoice();
        };

        /**
         * Copies error details to clipboard for debugging/sharing.
         * @param {Object} errorObj - The error object containing message and traceback.
         */
        const copyError = async (errorObj) => {
            const errorText = `Error: ${errorObj.message}\n\nStep: ${errorObj.step || 'N/A'}\n\nTraceback:\n${errorObj.traceback || 'No traceback available'}`;
            try {
                await navigator.clipboard.writeText(errorText);
                alert('Error copied to clipboard!');
            } catch (err) {
                console.error('Failed to copy error:', err);
            }
        };

        const summaryStats = computed(() => {
            return validationData.value?.database_export?.summary || null;
        });

        const weightStats = computed(() => {
            if (!validationData.value?.database_export?.packing_list_items) return null;
            const items = validationData.value.database_export.packing_list_items;
            let net = 0; let gross = 0; let cbm = 0;
            items.forEach(item => {
                try { net += parseFloat(item.net) || 0; } catch { }
                try { gross += parseFloat(item.gross) || 0; } catch { }
                try { cbm += parseFloat(item.cbm) || 0; } catch { }
            });
            return { net, gross, cbm };
        });

        /**
         * Computed: Extracts the config filename from the asset status path.
         */
        const assetConfigName = computed(() => {
            if (!assetStatus.value?.config_path) return 'Unknown';
            const path = assetStatus.value.config_path;
            return path.split(/[\\/]/).pop() || 'Unknown';
        });

        return {
            selectedFile,
            isUploading,
            uploadStatus,
            uploadError,
            showUploadTraceback,
            processingComplete,
            identifier,
            invoiceNo,
            invoiceDate,
            invoiceRef,
            includeStandard,
            includeCustom,
            includeDAF,
            handleFileUpload,
            uploadFile,
            isGenerating,
            generateInvoice,
            generationStatus,
            generationError,
            showGenTraceback,
            validationData,
            summaryStats,
            weightStats,
            retryUpload,
            retryGeneration,
            copyError,
            assetStatus,
            assetConfigName
        };
    }
};

