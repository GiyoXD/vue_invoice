import { useGeneratorStore } from '../stores/generatorStore.js';
import { storeToRefs } from 'pinia';

export default {
    name: 'InvoiceDetailsForm',
    emits: ['switch-view'],
    template: `
        <div class="bg-slate-800/80 backdrop-blur-md border border-slate-700/50 shadow-xl rounded-2xl p-8 mb-8" v-if="processingComplete">
            <h2 class="text-2xl font-bold text-slate-100 mb-6">2. Invoice Details</h2>
            
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
            <div v-if="assetStatus && assetStatus.ready" class="flex items-center mb-6 text-emerald-400 bg-emerald-500/10 border border-emerald-500/30 p-4 rounded-xl">
                <span class="mr-2">✅</span>
                <span class="text-slate-200">Blueprint found: using <strong class="text-emerald-400">{{ assetConfigName }}</strong></span>
                <span v-if="hasVariants" class="ml-2 bg-yellow-500/10 text-yellow-400 rounded px-2 py-1 text-xs font-bold border border-yellow-500/20">KH/VN variants detected</span>
            </div>
            
            <div class="flex flex-col gap-6">
                <div class="relative">
                    <label class="block text-slate-400 font-medium mb-2">Invoice Number</label>
                    <div class="flex gap-2">
                        <input class="flex-1 w-full bg-slate-900 border border-slate-700 rounded-lg px-4 py-2 text-slate-100 focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 transition-all" type="text" v-model="invoiceNo" />
                        <button class="px-4 py-2 bg-slate-700 hover:bg-slate-600 text-slate-200 rounded-lg shadow-sm transition-colors disabled:opacity-50 disabled:cursor-not-allowed" @click="() => lookupRefFromSheets(false)" :disabled="isLookingUp || !invoiceNo" title="Lookup Ref No in Google Sheets">
                            {{ isLookingUp ? '...' : '🔍' }}
                        </button>
                    </div>
                </div>
                <div>
                    <label class="block text-slate-400 font-medium mb-2">Invoice Date</label>
                    <input class="w-full bg-slate-900 border border-slate-700 rounded-lg px-4 py-2 text-slate-100 focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 transition-all" type="date" v-model="invoiceDate" />
                </div>
                <div>
                    <label class="flex items-center justify-between text-slate-400 font-medium mb-2">
                        <span>Invoice Ref (Optional)</span>
                        <span v-if="refSourceStatus" 
                              class="text-xs font-bold px-2 py-0.5 rounded-md" 
                              :class="refSourceStatus.type === 'found' ? 'bg-emerald-500/20 text-emerald-400 border border-emerald-500/30' : 'bg-amber-500/20 text-amber-400 border border-amber-500/30'">
                            {{ refSourceStatus.message }}
                        </span>
                    </label>
                    <input class="w-full bg-slate-900 border border-slate-700 rounded-lg px-4 py-2 text-slate-100 focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 transition-all" type="text" v-model="invoiceRef" @input="refSourceStatus = null" />
                </div>
            </div>

            <div class="mt-8">
                <label class="block text-slate-400 font-medium mb-3">Generation Options</label>
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
                    <label class="flex items-center gap-2 cursor-pointer border-l border-slate-700 pl-4">
                        <input type="checkbox" v-model="splitSheets" accent-color="#10b981" /> 
                        <span class="text-emerald-400 font-medium">Split Sheets into Separate Files</span>
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
                <input type="number" v-model="globalUnitPrice" step="0.01" min="0" placeholder="e.g. 1.25" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-4 py-2 text-slate-100 focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 transition-all max-w-xs" />
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
                        placeholder="Description"
                        class="w-full bg-slate-900 border border-slate-700 rounded-lg px-4 py-2 text-slate-100 focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 transition-all flex-[2] min-w-[150px]"
                    />
                    <input
                        type="number"
                        v-model="adj.value"
                        step="any"
                        placeholder="Amount"
                        class="w-full bg-slate-900 border border-slate-700 rounded-lg px-4 py-2 text-slate-100 focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 transition-all flex-[1] min-w-[120px]"
                    />
                    <button class="px-3 py-2 bg-slate-700 hover:bg-slate-600 text-slate-200 rounded-lg shadow-sm transition-colors w-auto" @click="removeAdjustment(index)">✕</button>
                </div>
                <button class="w-full px-4 py-2 mt-1 border border-slate-600 hover:bg-slate-700 text-slate-300 rounded-lg transition-colors" @click="addAdjustment">+ Add Adjustment</button>
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
                        <input type="text" v-model="googleSheetId" placeholder="Spreadsheet ID (Optional if set in backend)" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-4 py-2 text-slate-100 focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 transition-all flex-1 min-w-[250px]" />
                        <input type="text" v-model="googleSheetName" placeholder="Sheet Name (e.g. 2026)" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-4 py-2 text-slate-100 focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 transition-all w-40" />
                    </div>
                    <p class="text-secondary text-xs mt-2">
                        If checked, Ref No will be auto-fetched/incremented, and new invoices will be saved to the sheet.
                    </p>
                </div>
            </div>
            
            <button class="w-full px-6 py-3 mt-4 bg-gradient-to-r from-emerald-500 to-teal-500 hover:from-emerald-400 hover:to-teal-400 text-white font-bold rounded-xl shadow-lg shadow-emerald-500/20 transition-all transform hover:-translate-y-0.5 disabled:opacity-50 disabled:cursor-not-allowed disabled:transform-none" @click="generateInvoice" :disabled="isGenerating || !assetStatus?.ready">
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
                    <button class="px-6 py-2 bg-emerald-500 hover:bg-emerald-400 text-white font-medium rounded-lg shadow-md transition-colors disabled:opacity-50 disabled:cursor-not-allowed" @click="() => exportToSheets(false)" :disabled="isSyncing || showConflictConfirm">
                        {{ isSyncing ? 'Syncing...' : 'Push to Sheets' }}
                    </button>
                </div>
                <!-- Conflict Confirmation Panel (replaces window.confirm) -->
                <div v-if="showConflictConfirm" class="mt-3 p-3 bg-amber-10 border border-amber-30 rounded-md">
                    <p class="m-0 mb-2 text-yellow-400 text-sm font-bold">⚠️ {{ conflictMessage }}</p>
                    <div class="flex gap-2">
                        <button class="px-4 py-2 bg-red-500 hover:bg-red-400 text-white text-sm font-medium rounded-lg transition-colors disabled:opacity-50" @click="confirmOverride" :disabled="isSyncing">
                            {{ isSyncing ? 'Overriding...' : 'Override' }}
                        </button>
                        <button class="px-4 py-2 bg-slate-600 hover:bg-slate-500 text-white text-sm font-medium rounded-lg transition-colors disabled:opacity-50" @click="cancelOverride" :disabled="isSyncing">
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
    `,
    setup() {
        const store = useGeneratorStore();
        const {
            processingComplete,
            assetStatus,
            assetConfigName,
            hasVariants,
            invoiceNo,
            invoiceDate,
            invoiceRef,
            refSourceStatus,
            includeStandard,
            includeCustom,
            includeDAF,
            splitSheets,
            selectedVariants,
            isNetMode,
            globalUnitPrice,
            priceAdjustments,
            adjustmentError,
            isOnlineMode,
            showGoogleSheetsSettings,
            googleSheetId,
            googleSheetName,
            isGenerating,
            generationStatus,
            generationError,
            showGenTraceback,
            isSyncing,
            showConflictConfirm,
            conflictMessage,
            syncStatus,
            isLookingUp
        } = storeToRefs(store);

        return {
            processingComplete,
            assetStatus,
            assetConfigName,
            hasVariants,
            invoiceNo,
            invoiceDate,
            invoiceRef,
            refSourceStatus,
            includeStandard,
            includeCustom,
            includeDAF,
            splitSheets,
            selectedVariants,
            isNetMode,
            globalUnitPrice,
            priceAdjustments,
            adjustmentError,
            isOnlineMode,
            showGoogleSheetsSettings,
            googleSheetId,
            googleSheetName,
            isGenerating,
            generationStatus,
            generationError,
            showGenTraceback,
            isSyncing,
            showConflictConfirm,
            conflictMessage,
            syncStatus,
            isLookingUp,

            lookupRefFromSheets: store.lookupRefFromSheets,
            addAdjustment: store.addAdjustment,
            removeAdjustment: store.removeAdjustment,
            generateInvoice: store.generateInvoice,
            exportToSheets: store.exportToSheets,
            confirmOverride: store.confirmOverride,
            cancelOverride: store.cancelOverride,
            retryGeneration: store.retryGeneration,
            copyError: store.copyError
        };
    }
};
