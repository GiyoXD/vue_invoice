import { useGeneratorStore } from '../../stores/generatorStore.js';
import { storeToRefs } from 'pinia';

export default {
    name: 'FileUploader',
    template: `
        <div class="bg-slate-800/80 backdrop-blur-md border border-slate-700/50 shadow-2xl rounded-2xl p-8 mb-8 relative overflow-hidden group transition-all duration-300 hover:shadow-blue-500/10 hover:border-blue-500/30">
            <div class="absolute inset-0 bg-gradient-to-br from-blue-500/5 to-emerald-500/5 opacity-0 group-hover:opacity-100 transition-opacity duration-500 pointer-events-none"></div>
            <div class="relative z-10">
                <h2 class="text-2xl font-bold text-slate-100 mb-2">1. Upload Source Data</h2>
                <p class="text-slate-400 mb-6">Select your Excel file to begin processing.</p>
                
                <div class="flex items-center gap-4">
                    <input type="file" @change="onFileChange" accept="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel" class="block w-full text-sm text-slate-400 file:mr-4 file:py-2.5 file:px-4 file:rounded-full file:border-0 file:text-sm file:font-semibold file:bg-blue-500/10 file:text-blue-400 hover:file:bg-blue-500/20 transition-all cursor-pointer" />

                    
                    <button class="whitespace-nowrap px-6 py-2.5 bg-gradient-to-r from-blue-500 to-blue-600 hover:from-blue-400 hover:to-blue-500 text-white font-medium rounded-full shadow-lg shadow-blue-500/30 transition-all transform hover:-translate-y-0.5 disabled:opacity-50 disabled:cursor-not-allowed disabled:transform-none" @click="uploadFile" :disabled="!selectedFile || isUploading">
                        {{ isUploading ? 'Processing...' : 'Upload & Process' }}
                    </button>
                </div>
            </div>

            <div v-if="uploadStatus && !uploadError" :class="['status-box', uploadStatus.type, 'mt-6']">
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
                    <button v-if="uploadError.message && uploadError.message.includes('Weight Integrity Error')" class="px-4 py-2 bg-amber-600 hover:bg-amber-500 text-white font-medium rounded-lg shadow transition-colors ml-2" @click="ignoreTareAndRetry">⚠️ Bypass Weight Check</button>
                    <button v-if="uploadError.message && uploadError.message.includes('CBM')" class="px-4 py-2 bg-amber-600 hover:bg-amber-500 text-white font-medium rounded-lg shadow transition-colors ml-2" @click="ignoreCbmAndRetry">⚠️ Bypass CBM Check</button>
                    <button class="btn-copy-error" @click="copyError(uploadError)">📋 Copy Error</button>
                </div>
            </div>
        </div>
    `,
    setup() {
        const store = useGeneratorStore();
        const {
            selectedFile,
            isUploading,
            uploadStatus,
            uploadError,
            showUploadTraceback,
            validationWarnings
        } = storeToRefs(store);

        const onFileChange = (event) => {
            const file = event.target.files[0];
            store.setRawFile(file);
        };

        return {
            selectedFile,
            isUploading,
            uploadStatus,
            uploadError,
            showUploadTraceback,
            validationWarnings,
            onFileChange,
            uploadFile: store.uploadFile,
            retryUpload: store.retryUpload,
            ignoreTareAndRetry: store.ignoreTareAndRetry,
            ignoreCbmAndRetry: store.ignoreCbmAndRetry,
            copyError: store.copyError
        };
    }
};
