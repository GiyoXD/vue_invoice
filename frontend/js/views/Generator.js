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
                    <div class="form-group" style="position: relative;">
                        <label>Invoice Number</label>
                        <div style="display: flex; gap: 0.5rem;">
                            <input style="color: red; flex: 1;" type="text" v-model="invoiceNo" class="input-field" />
                            <button class="btn-small" @click="lookupRefFromSheets" :disabled="isLookingUp || !invoiceNo" style="margin: 0; padding: 0 0.75rem;" title="Lookup Ref No in Google Sheets">
                                {{ isLookingUp ? '...' : '🔍' }}
                            </button>
                        </div>
                    </div>
                    <div class="form-group">
                        <label>Invoice Date</label>
                        <input style="color: red;" type="date" v-model="invoiceDate" class="input-field" />
                    </div>
                    <div class="form-group">
                        <label style="display: flex; align-items: center; justify-content: space-between;">
                            <span>Invoice Ref (Optional)</span>
                            <span v-if="refSourceStatus" :style="{ fontSize: '0.75rem', fontWeight: 'bold', padding: '0.1rem 0.4rem', borderRadius: '4px', background: refSourceStatus.type === 'found' ? 'rgba(16, 185, 129, 0.15)' : 'rgba(245, 158, 11, 0.15)', color: refSourceStatus.type === 'found' ? '#10b981' : '#f59e0b' }">
                                {{ refSourceStatus.message }}
                            </span>
                        </label>
                        <input style="color: red;" type="text" v-model="invoiceRef" @input="refSourceStatus = null" class="input-field" />
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
                        <label style="display: flex; align-items: center; gap: 0.5rem; cursor: pointer; border-left: 1px solid #e2e8f0; padding-left: 1.5rem; margin-left: 0.5rem;">
                            <input type="checkbox" v-model="enableAutoFit" accent-color="#2563eb" /> 
                            <span>Auto-Fit Dimensions</span>
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

                <!-- NET WEIGHT PRICING MODE -->
                <div v-if="isNetMode" class="form-group" style="margin-top: 1rem; padding: 1rem; background: rgba(16, 185, 129, 0.05); border: 1px solid rgba(16, 185, 129, 0.2); border-radius: 8px;">
                    <label style="display: flex; align-items: center; gap: 0.5rem;">
                        <span style="color: #10b981; font-weight: bold;">⚖️ Net Weight Mode</span>
                    </label>
                    <p style="color: #94a3b8; font-size: 0.8rem; margin: 0.25rem 0 0.75rem 0;">
                        This template uses Net Weight as the pricing basis. Enter the unit price to calculate amounts.
                    </p>
                    <label>Unit Price (USD/kg)</label>
                    <input type="number" v-model="globalUnitPrice" step="0.01" min="0" class="input-field"
                           placeholder="e.g. 1.25" style="max-width: 200px;" />
                    <p style="color: #6b7280; font-size: 0.75rem; margin-top: 0.25rem;">
                        Amount = Net Weight × Unit Price
                    </p>
                </div>

                <div class="form-group" style="margin-top: 1rem;">
                    <label>Aggregation Adjustments</label>
                    <div v-for="(adj, index) in priceAdjustments" :key="index" style="display: flex; gap: 0.5rem; margin-bottom: 0.5rem;">
                        <input
                            type="text"
                            v-model="adj.description"
                            class="input-field"
                            placeholder="Reason (e.g. Shipping Discount)"
                            style="flex: 2;"
                        />
                        <input
                            type="number"
                            v-model="adj.value"
                            step="any"
                            class="input-field"
                            placeholder="Value"
                            style="flex: 1;"
                        />
                        <button class="btn-small" @click="removeAdjustment(index)" style="width: auto; margin-top: 0; padding: 0 0.75rem;">✕</button>
                    </div>
                    <button class="btn-small" @click="addAdjustment" style="width: 100%; margin-top: 0.25rem;">+ Add Adjustment</button>
                    <p v-if="adjustmentError" style="color: #f87171; font-size: 0.8rem; margin-top: 0.5rem;">
                        {{ adjustmentError }}
                    </p>
                    <p style="color: #94a3b8; font-size: 0.8rem; margin-top: 0.5rem;">
                        These will be evenly distributed across aggregation rows (col_amount).
                    </p>
                </div>
                <!-- GOOGLE SHEETS SETTINGS (ONLINE MODE) -->
                <div class="form-group" style="margin-top: 1.5rem; margin-bottom: 1.5rem;">
                    <div @click="showGoogleSheetsSettings = !showGoogleSheetsSettings" style="cursor: pointer; display: flex; align-items: center; gap: 0.5rem; color: #64748b; font-size: 0.9rem; font-weight: bold; user-select: none;">
                        <span style="transition: transform 0.2s; display: inline-block;" :style="{ transform: showGoogleSheetsSettings ? 'rotate(90deg)' : 'rotate(0deg)' }">▶</span>
                        <span>🌐 Google Sheets Sync Settings <span :style="{ color: isOnlineMode ? '#10b981' : '#94a3b8' }">{{ isOnlineMode ? '(Enabled)' : '(Disabled)' }}</span></span>
                    </div>
                    
                    <div v-if="showGoogleSheetsSettings" style="margin-top: 0.75rem; padding: 1rem; background: rgba(30, 41, 59, 0.5); border: 1px solid rgba(148, 163, 184, 0.2); border-radius: 8px;">
                        <label style="display: flex; align-items: center; gap: 0.5rem; cursor: pointer;">
                            <input type="checkbox" v-model="isOnlineMode" accent-color="#10b981" /> 
                            <span style="font-weight: bold; color: #e2e8f0; text-transform: uppercase; font-size: 0.85rem; letter-spacing: 0.05em;">Enable Online Sync</span>
                        </label>
                        <div v-if="isOnlineMode" style="margin-top: 0.75rem; display: flex; gap: 0.5rem; flex-wrap: wrap;">
                            <input type="text" v-model="googleSheetId" placeholder="Spreadsheet ID (Optional if set in backend)" class="input-field" style="flex: 1; min-width: 250px;" />
                            <input type="text" v-model="googleSheetName" placeholder="Sheet Name (e.g. 2026)" class="input-field" style="width: 150px;" />
                        </div>
                        <p style="color: #94a3b8; font-size: 0.75rem; margin-top: 0.5rem;">
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
                <div v-if="isOnlineMode && generationStatus && generationStatus.type === 'success' && !isGenerating" style="margin-top: 1rem; padding: 1rem; border: 1px solid rgba(148, 163, 184, 0.2); border-radius: 8px; background: rgba(30, 41, 59, 0.5);">
                    <div style="display: flex; align-items: center; justify-content: space-between;">
                        <div>
                            <h4 style="margin: 0; color: #e2e8f0;">Google Sheets Sync</h4>
                            <p style="margin: 0.25rem 0 0 0; font-size: 0.8rem; color: #94a3b8;">Push invoice data to Google Sheets.</p>
                        </div>
                        <button class="btn" @click="() => exportToSheets(false)" :disabled="isSyncing || showConflictConfirm" style="margin: 0; width: auto; padding: 0.5rem 1rem; background: #10b981;">
                            {{ isSyncing ? 'Syncing...' : 'Push to Sheets' }}
                        </button>
                    </div>
                    <!-- Conflict Confirmation Panel (replaces window.confirm) -->
                    <div v-if="showConflictConfirm" style="margin-top: 0.75rem; padding: 0.75rem; background: rgba(245, 158, 11, 0.1); border: 1px solid rgba(245, 158, 11, 0.3); border-radius: 6px;">
                        <p style="margin: 0 0 0.5rem 0; color: #fbbf24; font-size: 0.9rem; font-weight: bold;">⚠️ {{ conflictMessage }}</p>
                        <div style="display: flex; gap: 0.5rem;">
                            <button class="btn" @click="confirmOverride" :disabled="isSyncing" style="margin: 0; width: auto; padding: 0.4rem 1rem; background: #ef4444; font-size: 0.85rem;">
                                {{ isSyncing ? 'Overriding...' : 'Override' }}
                            </button>
                            <button class="btn" @click="cancelOverride" :disabled="isSyncing" style="margin: 0; width: auto; padding: 0.4rem 1rem; background: #475569; font-size: 0.85rem;">
                                Cancel
                            </button>
                        </div>
                    </div>
                    <div v-if="syncStatus" :class="['status-box', syncStatus.type]" style="margin-top: 0.75rem; margin-bottom: 0; padding: 0.5rem; font-size: 0.85rem;">
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
                    <div v-if="generationError.details && generationError.details.length" class="error-details-list" style="margin-top: 1rem; padding-top: 1rem; border-top: 1px solid rgba(248, 113, 113, 0.2);">
                        <h4 style="font-size: 0.85rem; color: #f87171; margin-bottom: 0.5rem; text-transform: uppercase; letter-spacing: 0.025em;">Root Causes:</h4>
                        <ul style="margin: 0; padding-left: 1.25rem; font-size: 0.85rem; color: #fca5a5;">
                            <li v-for="(detail, idx) in generationError.details" :key="idx" style="margin-bottom: 0.25rem;">
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

