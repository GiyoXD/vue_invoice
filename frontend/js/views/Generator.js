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

                <!-- NORMALIZATION WARNINGS PANEL -->
                <div v-if="validationWarnings && validationWarnings.length > 0" class="warning-panel">
                    <div class="warning-header" style="display: flex; align-items: center; gap: 0.5rem; margin-bottom: 0.75rem;">
                        <span class="warning-icon" style="font-size: 1.25rem;">⚠️</span>
                        <h3 style="margin: 0; color: #b45309; font-size: 1rem;">Data Auto-Correction Notices</h3>
                    </div>
                    <ul style="margin: 0; padding-left: 1.5rem; color: #92400e; font-size: 0.9rem;">
                        <li v-for="(msg, idx) in validationWarnings" :key="idx" style="margin-bottom: 0.25rem;">
                            {{ msg }}
                        </li>
                    </ul>
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
                    <span v-if="hasVariants" style="margin-left: 0.5rem; padding: 0.15rem 0.5rem; background: rgba(234, 179, 8, 0.15); color: #facc15; border-radius: 4px; font-size: 0.75rem; font-weight: bold;">KH/VN variants detected</span>
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
                    
                    <!-- KH/VN Variant Options -->
                    <div v-if="hasVariants" style="display: flex; gap: 1.5rem; margin-top: 0.75rem; padding: 0.75rem; background: rgba(234, 179, 8, 0.05); border: 1px solid rgba(234, 179, 8, 0.15); border-radius: 6px; flex-wrap: wrap;">
                        <span style="color: #facc15; font-weight: bold; font-size: 0.85rem; align-self: center;">Variants:</span>
                        <label v-for="v in assetStatus.variants" :key="v.suffix" style="display: flex; align-items: center; gap: 0.5rem; cursor: pointer;">
                            <input type="checkbox" v-model="selectedVariants" :value="v.suffix" accent-color="#eab308" />
                            <span>{{ v.suffix.replace('_', '') }} version</span>
                        </label>
                    </div>
                </div>

                <div class="form-group" style="margin-top: 1rem;">
                    <label>Aggregation Adjustment (x)</label>
                    <input
                        type="number"
                        v-model="adjustmentInput"
                        step="any"
                        class="input-field"
                        placeholder="e.g. 100 or -50"
                    />
                    <p style="color: #94a3b8; font-size: 0.8rem; margin-top: 0.25rem;">
                        Evenly distributed across aggregation rows (col_amount).
                    </p>
                    <p v-if="adjustmentError" style="color: #f87171; font-size: 0.8rem; margin-top: 0.25rem;">
                        {{ adjustmentError }}
                    </p>
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
        const selectedVariants = ref([]);

        const adjustmentInput = ref('');
        const adjustmentError = ref('');

        const isGenerating = ref(false);
        const generationStatus = ref(null);
        const generationError = ref(null);
        const showGenTraceback = ref(false);
        const validationData = ref(null); // Validation data from generation
        const validationWarnings = ref([]); // Validation warnings from extraction step
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
            selectedVariants.value = [];
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
            validationWarnings.value = []; // Clear previous warnings

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

                    assetStatus.value = data.asset_status || null;

                    // Auto-select all available variants
                    if (data.asset_status?.variants?.length > 0) {
                        selectedVariants.value = data.asset_status.variants.map(v => v.suffix);
                    }

                    // Store any normalization warnings
                    if (data.warnings && data.warnings.length > 0) {
                        validationWarnings.value = data.warnings;
                        uploadStatus.value = { type: 'warning', message: 'File processed successfully, but with data corrections.' };
                    } else {
                        validationWarnings.value = [];
                    }

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
         * Validates the aggregation adjustment input.
         * Returns an object with isValid flag and numeric value (or null if empty).
         */
        const validateAdjustment = () => {
            const raw = adjustmentInput.value != null ? String(adjustmentInput.value).trim() : '';
            if (raw === '') {
                adjustmentError.value = '';
                return { isValid: true, value: null };
            }

            // Allow leading + / - , digits, and optional decimal portion
            const numPattern = /^[+-]?\d+(\.\d+)?$/;
            if (!numPattern.test(raw)) {
                adjustmentError.value = 'Please enter a valid number (e.g. 100 or -33.75).';
                return { isValid: false, value: null };
            }

            const parsed = Number(raw);
            if (!Number.isFinite(parsed)) {
                adjustmentError.value = 'Please enter a valid number.';
                return { isValid: false, value: null };
            }

            adjustmentError.value = '';
            return { isValid: true, value: parsed };
        };

        /**
         * Triggers invoice generation with the provided metadata.
         * Handles both success and error responses.
         */
        const generateInvoice = async () => {
            const { isValid, value: adjustmentValue } = validateAdjustment();
            if (!isValid) {
                generationStatus.value = null;
                generationError.value = {
                    message: 'Invalid aggregation adjustment. Please enter a valid number.',
                    step: 'Validation',
                    traceback: null
                };
                return;
            }

            isGenerating.value = true;
            generationStatus.value = { type: 'info', message: 'Generating invoice, please wait...' };
            generationError.value = null;
            validationData.value = null;

            try {
                const basePayload = {
                    identifier: identifier.value,
                    json_path: jsonPath.value,
                    invoice_no: invoiceNo.value,
                    invoice_date: invoiceDate.value,
                    invoice_ref: invoiceRef.value,
                    generate_standard: includeStandard.value,
                    generate_custom: includeCustom.value,
                    generate_daf: includeDAF.value,
                    generate_kh: selectedVariants.value.includes('_KH'),
                    generate_vn: selectedVariants.value.includes('_VN')
                };

                if (adjustmentValue !== null) {
                    basePayload.aggregation_adjustment = adjustmentValue;
                }

                const response = await fetch('/api/generate', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(basePayload)
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

        /**
         * Computed: Whether KH/VN variants are available.
         */
        const hasVariants = computed(() => {
            return (assetStatus.value?.variants?.length || 0) > 0;
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
            assetConfigName,
            hasVariants,
            selectedVariants,
            adjustmentInput,
            adjustmentError,
            validationWarnings
        };
    }
};

