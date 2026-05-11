import { ref, computed } from 'vue';

export default {
    emits: ['switch-view'], // Declare event to switch tabs
    template: `
        <div class="generator-view fade-in">
            <h1>Invoice Generator</h1>
            
            <div class="card">
                <h2>1. Upload Source Data</h2>
                <p class="text-secondary mb-4">Select your Excel file to begin processing.</p>
                
                <input type="file" @change="handleFileUpload" accept=".xlsx, .xls" />
                
                <button class="btn" @click="uploadFile" :disabled="!selectedFile || isUploading">
                    {{ isUploading ? 'Processing...' : 'Upload & Process' }}
                </button>

                <div v-if="uploadStatus && !uploadError" :class="['status-box', uploadStatus.type]">
                    {{ uploadStatus.message }}
                </div>

                <!-- NORMALIZATION WARNINGS PANEL -->
                <div v-if="validationWarnings && validationWarnings.length > 0" class="warning-panel">
                    <div class="warning-header flex items-center gap-2 mb-3">
                        <span class="warning-icon text-xl">⚠️</span>
                        <h3 class="m-0 text-amber-700 text-base">Data Auto-Correction Notices</h3>
                    </div>
                    <ul class="m-0 pl-6 text-amber-800 text-sm">
                        <li v-for="(msg, idx) in validationWarnings" :key="idx" class="mb-1">
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

            <div class="card delay-200" v-if="processingComplete">
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
                    <span v-if="hasVariants" class="ml-2 bg-yellow-15 text-yellow-400 rounded px-2 py-1 text-xs font-bold">KH/VN variants detected</span>
                </div>
                
                <div class="grid-form">
                    <div class="form-group relative">
                        <label>Invoice Number</label>
                        <div class="flex gap-2">
                            <input class="text-red-500 flex-1 input-field" type="text" v-model="invoiceNo" />
                            <button class="btn-small m-0 px-3" @click="lookupRefFromSheets" :disabled="isLookingUp || !invoiceNo" title="Lookup Ref No in Google Sheets">
                                {{ isLookingUp ? '...' : '🔍' }}
                            </button>
                        </div>
                    </div>
                    <div class="form-group">
                        <label>Invoice Date</label>
                        <input class="text-red-500 input-field" type="date" v-model="invoiceDate" />
                    </div>
                    <div class="form-group">
                        <label class="flex items-center justify-between">
                            <span>Invoice Ref (Optional)</span>
                            <span v-if="refSourceStatus" 
                                  class="text-xs font-bold px-2 py-0-5 rounded-md" 
                                  :class="refSourceStatus.type === 'found' ? 'bg-emerald-500-15 text-emerald-500' : 'bg-amber-500-15 text-amber-500'">
                                {{ refSourceStatus.message }}
                            </span>
                        </label>
                        <input class="text-red-500 input-field" type="text" v-model="invoiceRef" @input="refSourceStatus = null" />
                    </div>
                </div>

                <div class="form-group mt-4">
                    <label>Generation Options</label>
                    <div class="flex gap-6 mt-2 flex-wrap">
                        <label class="flex items-center gap-2 cursor-pointer">
                            <input type="checkbox" v-model="includeStandard" accent-color="#2563eb" /> 
                            <span>Standard Invoice</span>
                        </label>
                        <label class="flex items-center gap-2 cursor-pointer">
                            <input type="checkbox" v-model="includeCustom" accent-color="#2563eb" /> 
                            <span>Custom Mode</span>
                        </label>
                        <label class="flex items-center gap-2 cursor-pointer">
                            <input type="checkbox" v-model="includeDAF" accent-color="#2563eb" /> 
                            <span>DAF Mode</span>
                        </label>
                        <label class="flex items-center gap-2 cursor-pointer border-l border-gray-200 pl-6 ml-2">
                            <input type="checkbox" v-model="enableAutoFit" accent-color="#2563eb" /> 
                            <span>Auto-Fit Dimensions</span>
                        </label>
                    </div>
                    
                    <!-- KH/VN Variant Options -->
                    <div v-if="hasVariants" class="flex gap-6 mt-3 p-3 bg-yellow-5 border border-yellow-15 rounded-md flex-wrap">
                        <span class="text-yellow-400 font-bold text-sm self-center">Variants:</span>
                        <label v-for="v in assetStatus.variants" :key="v.suffix" class="flex items-center gap-2 cursor-pointer">
                            <input type="checkbox" v-model="selectedVariants" :value="v.suffix" accent-color="#eab308" />
                            <span>{{ v.suffix.replace('_', '') }} version</span>
                        </label>
                    </div>
                </div>

                <!-- NET WEIGHT PRICING MODE -->
                <div v-if="isNetMode" class="form-group mt-4 p-4 bg-emerald-5 border border-emerald-20 rounded-xl">
                    <label class="flex items-center gap-2">
                        <span class="text-emerald-500 font-bold">⚖️ Net Weight Mode</span>
                    </label>
                    <p class="text-secondary text-sm my-1 mb-3">
                        This template uses Net Weight as the pricing basis. Enter the unit price to calculate amounts.
                    </p>
                    <label>Unit Price (USD/kg)</label>
                    <input type="number" v-model="globalUnitPrice" step="0.01" min="0" placeholder="e.g. 1.25" class="input-field max-w-xs" />
                    <p class="text-gray-500 text-xs mt-1">
                        Amount = Net Weight × Unit Price
                    </p>
                </div>

                <div class="form-group mt-4">
                    <label>Aggregation Adjustments</label>
                    <div v-for="(adj, index) in priceAdjustments" :key="index" class="flex gap-2 mb-2">
                        <input
                            type="text"
                            v-model="adj.description"
                            class="input-field flex-2"
                        />
                        <input
                            type="number"
                            v-model="adj.value"
                            step="any"
                            class="input-field flex-1"
                        />
                        <button class="btn-small w-auto m-0 px-3" @click="removeAdjustment(index)">✕</button>
                    </div>
                    <button class="btn-small w-full mt-1" @click="addAdjustment">+ Add Adjustment</button>
                    <p v-if="adjustmentError" class="text-red-400 text-sm mt-2">
                        {{ adjustmentError }}
                    </p>
                    <p class="text-secondary text-sm mt-2">
                        These will be evenly distributed across aggregation rows (col_amount).
                    </p>
                </div>
                <!-- GOOGLE SHEETS SETTINGS (ONLINE MODE) -->
                <div class="form-group my-6">
                    <div @click="showGoogleSheetsSettings = !showGoogleSheetsSettings" class="cursor-pointer flex items-center gap-2 text-muted text-sm font-bold select-none">
                        <span class="transition-transform inline-block" :class="{ 'rotate-90': showGoogleSheetsSettings }">▶</span>
                        <span>🌐 Google Sheets Sync Settings <span :class="isOnlineMode ? 'text-emerald-500' : 'text-secondary'">{{ isOnlineMode ? '(Enabled)' : '(Disabled)' }}</span></span>
                    </div>
                    
                    <div v-if="showGoogleSheetsSettings" class="mt-3 p-4 bg-slate-500-50 border border-slate-400-20 rounded-xl">
                        <label class="flex items-center gap-2 cursor-pointer">
                            <input type="checkbox" v-model="isOnlineMode" accent-color="#10b981" /> 
                            <span class="font-bold text-primary uppercase text-sm tracking-wider">Enable Online Sync</span>
                        </label>
                        <div v-if="isOnlineMode" class="mt-3 flex gap-2 flex-wrap">
                            <input type="text" v-model="googleSheetId" placeholder="Spreadsheet ID (Optional if set in backend)" class="input-field flex-1 min-w-64" />
                            <input type="text" v-model="googleSheetName" placeholder="Sheet Name (e.g. 2026)" class="input-field w-40" />
                        </div>
                        <p class="text-secondary text-xs mt-2">
                            If checked, Ref No will be auto-fetched/incremented, and new invoices will be saved to the sheet.
                        </p>
                    </div>
                </div>
                
                <button class="btn" @click="generateInvoice" :disabled="isGenerating || !assetStatus?.ready">
                    {{ isGenerating ? 'Generating...' : (assetStatus?.ready ? 'Generate Invoice' : 'Blueprint Required') }}
                </button>

                <div v-if="generationStatus && !generationError" :class="['status-box', generationStatus.type]">
                    {{ generationStatus.message }}
                </div>

                <!-- Google Sheets Sync Button (Visible after successful generation) -->
                <div v-if="isOnlineMode && generationStatus && generationStatus.type === 'success' && !isGenerating" class="mt-4 p-4 border border-slate-400-20 rounded-xl bg-slate-500-50">
                    <div class="flex items-center justify-between">
                        <div>
                            <h4 class="m-0 text-primary">Google Sheets Sync</h4>
                            <p class="mt-1 mb-0 text-sm text-secondary">Push invoice data to Google Sheets.</p>
                        </div>
                        <button class="btn m-0 w-auto px-4 py-2 bg-emerald-500" @click="() => exportToSheets(false)" :disabled="isSyncing || showConflictConfirm">
                            {{ isSyncing ? 'Syncing...' : 'Push to Sheets' }}
                        </button>
                    </div>
                    <!-- Conflict Confirmation Panel (replaces window.confirm) -->
                    <div v-if="showConflictConfirm" class="mt-3 p-3 bg-amber-10 border border-amber-30 rounded-md">
                        <p class="m-0 mb-2 text-yellow-400 text-sm font-bold">⚠️ {{ conflictMessage }}</p>
                        <div class="flex gap-2">
                            <button class="btn m-0 w-auto py-1 px-4 bg-red-500 text-sm" @click="confirmOverride" :disabled="isSyncing">
                                {{ isSyncing ? 'Overriding...' : 'Override' }}
                            </button>
                            <button class="btn m-0 w-auto py-1 px-4 bg-slate-600 text-sm" @click="cancelOverride" :disabled="isSyncing">
                                Cancel
                            </button>
                        </div>
                    </div>
                    <div v-if="syncStatus" :class="['status-box', syncStatus.type]" class="mt-3 mb-0 p-2 text-sm">
                        {{ syncStatus.message }}
                    </div>
                </div>

                <!-- ERROR PANEL FOR GENERATION -->
                <div v-if="generationError" class="error-panel">
                    <div class="error-header">
                        <span class="error-icon">⚠️</span>
                        <h3>Generation Failed</h3>
                    </div>
                    <span v-if="generationError.step" class="error-step">{{ generationError.step }}</span>
                    <div class="error-message">{{ generationError.message }}</div>
                    
                    <!-- DETAILED ERROR LIST -->
                    <div v-if="generationError.details && generationError.details.length" class="error-details-list mt-4 pt-4 border-t border-red-400-20">
                        <h4 class="text-sm text-red-400 mb-2 uppercase tracking-wide">Root Causes:</h4>
                        <ul class="m-0 pl-5 text-sm text-red-300">
                            <li v-for="(detail, idx) in generationError.details" :key="idx" class="mb-1">
                                {{ detail }}
                            </li>
                        </ul>
                    </div>
                    
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
            <div class="card validation-card delay-100" v-if="validationData && !isGenerating && !generationError">
                <div class="validation-header">
                    <h3>✅ Invoice Generated Successfully</h3>
                    <span class="text-sm opacity-70">{{ validationData.timestamp }}</span>
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
        const refSourceStatus = ref(null); // { type: 'found'|'new', message: '' }

        // Options
        const includeStandard = ref(true);
        const includeCustom = ref(false);
        const includeDAF = ref(false);
        const selectedVariants = ref([]);
        const enableAutoFit = ref(true);

        const priceAdjustments = ref([]); // List of { description: '', value: '' }
        const adjustmentError = ref('');
        const globalUnitPrice = ref(''); // For 'net' pricing mode

        // Computed: is this template using net weight pricing?
        const isNetMode = computed(() => assetStatus.value?.pricing_mode === 'net');

        const isGenerating = ref(false);
        const generationStatus = ref(null);
        const generationError = ref(null);
        const showGenTraceback = ref(false);
        const validationData = ref(null); // Validation data from generation
        const validationWarnings = ref([]); // Validation warnings from extraction step
        const assetStatus = ref(null); // Asset availability status from upload

        // Google Sheets Export
        const isOnlineMode = ref(true);
        const showGoogleSheetsSettings = ref(false);
        const googleSheetId = ref(''); // Leave empty to use backend default (DEFAULT_SPREADSHEET_ID)
        const googleSheetName = ref('2026');

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
            priceAdjustments.value = [];
            refSourceStatus.value = null;
        };

        const addAdjustment = () => {
            priceAdjustments.value.push({ description: '', value: '' });
        };

        const removeAdjustment = (index) => {
            priceAdjustments.value.splice(index, 1);
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
        const validateAdjustments = () => {
            const validSet = [];
            for (const adj of priceAdjustments.value) {
                const desc = (adj.description || '').trim();
                const valRaw = String(adj.value || '').trim();

                if (desc === '' && valRaw === '') continue; // Skip empty rows

                if (valRaw === '') {
                    adjustmentError.value = 'Please enter a value for all adjustments.';
                    return { isValid: false, list: [] };
                }

                const parsed = Number(valRaw);
                if (isNaN(parsed) || !Number.isFinite(parsed)) {
                    adjustmentError.value = `Invalid value for "${desc || 'adjustment'}".`;
                    return { isValid: false, list: [] };
                }

                validSet.push([desc || 'Adjustment', parsed]);
            }

            adjustmentError.value = '';
            return { isValid: true, list: validSet };
        };

        /**
         * Triggers invoice generation with the provided metadata.
         * Handles both success and error responses.
         */
        const generateInvoice = async () => {
            const { isValid, list: validAdjustments } = validateAdjustments();
            if (!isValid) {
                generationStatus.value = null;
                generationError.value = {
                    message: adjustmentError.value,
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
                // Online Mode: Auto-resolve Ref No before generation if empty
                if (isOnlineMode.value && !invoiceRef.value) {
                    generationStatus.value = { type: 'info', message: 'Resolving Ref No from Google Sheets...' };
                    await lookupRefFromSheets(true); // silent lookup
                }

                const basePayload = {
                    identifier: identifier.value,
                    json_path: jsonPath.value,
                    invoice_no: invoiceNo.value,
                    invoice_date: invoiceDate.value,
                    invoice_ref: invoiceRef.value,
                    generate_standard: includeStandard.value,
                    generate_custom: includeCustom.value,
                    generate_daf: includeDAF.value,
                    generate_kh: true,  // KH is the default variant
                    generate_vn: selectedVariants.value.includes('_VN'),
                    auto_fit: enableAutoFit.value
                };

                // Net weight pricing mode: include global unit price
                if (isNetMode.value && globalUnitPrice.value) {
                    basePayload.global_unit_price = parseFloat(globalUnitPrice.value);
                }

                if (validAdjustments.length > 0) {
                    basePayload.price_adjustment = validAdjustments;
                }

                const response = await fetch('/api/generate', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(basePayload)
                });

                const data = await response.json();

                if (response.ok) {
                    generationStatus.value = { type: 'success', message: 'Invoice generated successfully! Download starting...' };
                    if (data.metadata) {
                        validationData.value = data.metadata;
                    }
                    // Warn if metadata read failed on the backend (Sheets sync will have wrong values)
                    if (data.metadata_error) {
                        console.error('[Generate] metadata_error:', data.metadata_error);
                        generationStatus.value = { type: 'warning', message: `Invoice generated, but metadata could not be loaded: ${data.metadata_error}. Google Sheets sync may push incorrect values.` };
                    }
                    if (data.files && data.files.length > 0) {
                        data.files.forEach(f => {
                            const mimeType = f.mime_type || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

                            // Convert Base64 to Blob for robust downloading of large files
                            const binaryString = window.atob(f.content);
                            const bytes = new Uint8Array(binaryString.length);
                            for (let i = 0; i < binaryString.length; i++) {
                                bytes[i] = binaryString.charCodeAt(i);
                            }
                            const blob = new Blob([bytes], { type: mimeType });
                            const url = URL.createObjectURL(blob);

                            const link = document.createElement('a');
                            link.href = url;
                            link.download = f.filename;
                            document.body.appendChild(link);
                            link.click();
                            document.body.removeChild(link);
                            URL.revokeObjectURL(url);
                        });
                    }
                    
                    // Reset sync status when a new invoice is generated
                    syncStatus.value = null;
                    
                } else {
                    // Capture structured error from API
                    generationError.value = {
                        message: data.error || 'Generation failed',
                        details: data.details || [],
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

        const isSyncing = ref(false);
        const syncStatus = ref(null);

        const showConflictConfirm = ref(false);
        const conflictMessage = ref('');

        /**
         * Export generated data to Google Sheets
         */
        const exportToSheets = async (forceOverride = false) => {
            isSyncing.value = true;
            syncStatus.value = { type: 'info', message: 'Syncing to Google Sheets...' };
            try {
                // Fetch directly from the grand_total footer data
                const grandTotal = validationData.value?.footer_data?.grand_total || {};

                // --- DEBUG: Log what we found in validationData ---
                console.log('[Sheets Export] validationData keys:', validationData.value ? Object.keys(validationData.value) : 'NULL');
                console.log('[Sheets Export] grand_total:', grandTotal);

                const pallets = grandTotal.col_pallet_count ?? (summaryStats.value?.total_pallets || 0);
                let gross = grandTotal.col_gross ?? (weightStats.value?.gross || 0);
                
                // Format gross to drop trailing zeroes if it's a string (e.g. "8290.5000" -> "8290.5")
                if (typeof gross === 'string' && !isNaN(parseFloat(gross))) {
                    gross = parseFloat(gross).toString();
                }

                console.log('[Sheets Export] pallets:', pallets, '| gross:', gross);

                // Guard: warn user if values look empty/zero
                if (!pallets && !gross) {
                    syncStatus.value = { type: 'error', message: 'Cannot sync: pallet and gross weight data is missing. Please regenerate the invoice first.' };
                    isSyncing.value = false;
                    return;
                }

                const palletStr = `${pallets} PALLETS: ${gross}`;
                console.log('[Sheets Export] palletStr:', palletStr);

                const payload = {
                    invoice_no: invoiceNo.value || identifier.value,
                    ref_no: invoiceRef.value || '',
                    invoice_date: invoiceDate.value,
                    pallet_str: palletStr,
                    force_override: forceOverride === true,
                    worksheet_name: googleSheetName.value || '2026'
                };

                if (googleSheetId.value && googleSheetId.value.trim() !== '') {
                    payload.spreadsheet_id = googleSheetId.value.trim();
                }

                console.log('[Sheets Export] payload:', payload);

                const response = await fetch('/api/sheets/export', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(payload)
                });
                
                const data = await response.json();
                console.log('[Sheets Export] response:', data);
                
                if (!response.ok) {
                    console.error("Sheets sync failed:", data);
                    syncStatus.value = { type: 'error', message: data.error || 'Failed to sync to Google Sheets.' };
                } else if (data.action === 'conflict') {
                    // Show inline confirmation instead of blocking window.confirm()
                    syncStatus.value = null;
                    conflictMessage.value = data.message;
                    showConflictConfirm.value = true;
                } else {
                    syncStatus.value = { type: 'success', message: data.message || 'Successfully synced to Google Sheets!' };
                }
            } catch (error) {
                console.error("Sheets network error:", error);
                syncStatus.value = { type: 'error', message: 'Network error while syncing.' };
            } finally {
                isSyncing.value = false;
            }
        };

        /**
         * Computed: Whether KH/VN variants are available.
         */
        const hasVariants = computed(() => {
            return (assetStatus.value?.variants?.length || 0) > 0;
        });

        const confirmOverride = async () => {
            showConflictConfirm.value = false;
            await exportToSheets(true);
        };

        const cancelOverride = () => {
            showConflictConfirm.value = false;
            syncStatus.value = { type: 'info', message: 'Sync cancelled.' };
        };

        const isLookingUp = ref(false);

        /**
         * Look up Ref No from Google Sheets
         */
        const lookupRefFromSheets = async (silent = false) => {
            if (!invoiceNo.value) return;
            isLookingUp.value = true;
            try {
                const payload = {
                    invoice_no: invoiceNo.value,
                    worksheet_name: googleSheetName.value || '2026'
                };
                if (googleSheetId.value && googleSheetId.value.trim() !== '') {
                    payload.spreadsheet_id = googleSheetId.value.trim();
                }

                const response = await fetch('/api/sheets/resolve_ref', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(payload)
                });

                if (response.ok) {
                    const data = await response.json();
                    invoiceRef.value = data.ref_no;
                    refSourceStatus.value = {
                        type: data.found ? 'found' : 'new',
                        message: data.found ? '✓ Pulled from Sheets' : '✨ Auto-generated'
                    };
                } else {
                    const err = await response.json();
                    if (!silent) alert(`Lookup Failed: ${err.error || 'Check console'}`);
                }
            } catch (error) {
                if (!silent) alert(`Network Error: ${error.message}`);
            } finally {
                isLookingUp.value = false;
            }
        };

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
            enableAutoFit,
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
            adjustmentError,
            validationWarnings,
            priceAdjustments,
            addAdjustment,
            removeAdjustment,
            globalUnitPrice,
            isNetMode,
            googleSheetId,
            googleSheetName,
            isOnlineMode,
            showGoogleSheetsSettings,
            isLookingUp,
            lookupRefFromSheets,
            isSyncing,
            syncStatus,
            exportToSheets,
            refSourceStatus,
            showConflictConfirm,
            conflictMessage,
            confirmOverride,
            cancelOverride
        };
    }
};

