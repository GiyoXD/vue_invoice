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

                <div v-if="uploadStatus" :class="['status-box', uploadStatus.type]">
                    {{ uploadStatus.message }}
                </div>
            </div>

            <div class="card" v-if="processingComplete" style="animation-delay: 0.2s">
                <h2>2. Invoice Details</h2>
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
                
                <button class="btn" @click="generateInvoice" :disabled="isGenerating">
                    {{ isGenerating ? 'Generating...' : 'Generate Invoice' }}
                </button>

                <div v-if="generationStatus" :class="['status-box', generationStatus.type]">
                    {{ generationStatus.message }}
                </div>
            </div>

            <!-- VALIDATION CARD -->
            <div class="card validation-card" v-if="validationData && !isGenerating" style="animation-delay: 0.1s">
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

        const processingComplete = ref(false);
        const identifier = ref('');
        const jsonPath = ref('');

        const invoiceNo = ref('');
        const invoiceDate = ref(new Date().toISOString().split('T')[0]);
        const invoiceRef = ref('');

        const isGenerating = ref(false);
        const generationStatus = ref(null);
        const validationData = ref(null); // Validation data from generation

        // --- Generator Actions ---
        const handleFileUpload = (event) => {
            selectedFile.value = event.target.files[0];
            uploadStatus.value = null;
            processingComplete.value = false;
            validationData.value = null;
        };

        const uploadFile = async () => {
            if (!selectedFile.value) return;

            isUploading.value = true;
            uploadStatus.value = { type: 'info', message: 'Uploading and processing...' };
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
                    uploadStatus.value = { type: 'success', message: 'File processed & configuration generated!' };
                    identifier.value = data.identifier;
                    jsonPath.value = data.json_path;

                    invoiceNo.value = data.default_inv_no || '';

                    processingComplete.value = true;
                } else {
                    throw new Error(data.error || 'Upload failed');
                }
            } catch (error) {
                uploadStatus.value = { type: 'error', message: error.message };
            } finally {
                isUploading.value = false;
            }
        };

        const generateInvoice = async () => {
            isGenerating.value = true;
            generationStatus.value = { type: 'info', message: 'Generating invoice, please wait...' };
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
                        invoice_ref: invoiceRef.value
                    })
                });

                const data = await response.json();

                if (response.ok) {
                    generationStatus.value = { type: 'success', message: 'Invoice generated successfully!' };
                    if (data.metadata) {
                        validationData.value = data.metadata;
                    }
                } else {
                    throw new Error(data.error || 'Generation failed');
                }
            } catch (error) {
                generationStatus.value = { type: 'error', message: error.message };
            } finally {
                isGenerating.value = false;
            }
        };

        const summaryStats = computed(() => {
            return validationData.value?.database_export?.summary || null;
        });

        const weightStats = computed(() => {
            if (!validationData.value?.database_export?.packing_list_items) return null;
            const items = validationData.value.database_export.packing_list_items;
            let net = 0; let gross = 0;
            items.forEach(item => {
                try { net += parseFloat(item.net) || 0; } catch { }
                try { gross += parseFloat(item.gross) || 0; } catch { }
            });
            return { net, gross };
        });

        return {
            selectedFile,
            isUploading,
            uploadStatus,
            processingComplete,
            identifier,
            invoiceNo,
            invoiceDate,
            invoiceRef,
            handleFileUpload,
            uploadFile,
            isGenerating,
            generateInvoice,
            generationStatus,
            validationData,
            summaryStats,
            weightStats
        };
    }
};
