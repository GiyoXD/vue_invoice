import { useTemplateExtractorStore } from '../../stores/templateExtractorStore.js';

export default {
    name: 'ExtractorUploadStep',
    template: `
        <div class="bg-slate-800/80 backdrop-blur-md border border-slate-700/50 shadow-2xl rounded-2xl p-8 mb-8" v-if="store.currentStep === 1">
            <h2>1. Analyze Invoice Source</h2>
            <p class="text-secondary mb-4">
                Upload a sample invoice file. Upload <strong>2 files</strong> to auto-create KH + VN versions.
            </p>
            
            <input type="file" @change="store.handleFileUpload" accept=".xlsx, .xls" multiple class="block w-full text-sm text-slate-400 file:mr-4 file:py-2.5 file:px-4 file:rounded-full file:border-0 file:text-sm file:font-semibold file:bg-blue-500/10 file:text-blue-400 hover:file:bg-blue-500/20 transition-all cursor-pointer" />
            
            <!-- Show selected files with KH/VN labels -->
            <div v-if="store.selectedFiles.length > 0" class="mt-4">
                <div v-for="(file, idx) in store.selectedFiles" :key="idx" 
                     class="flex items-center gap-3 py-2 px-3 mb-2 bg-white/5 rounded-md border border-white/10">
                    <span v-if="store.selectedFiles.length === 2" 
                          class="px-2 py-1 rounded text-xs font-bold border"
                          :class="idx === 0 ? 'bg-blue-400/20 text-blue-400 border-blue-400/30' : 'bg-yellow-400/20 text-yellow-400 border-yellow-400/30'">
                        {{ idx === 0 ? 'KH' : 'VN' }}
                    </span>
                    <span class="text-primary">📄 {{ file.name }}</span>
                </div>
                
                <!-- Single-file suffix selector -->
                <div v-if="store.selectedFiles.length === 1" class="mt-3 flex items-center gap-3">
                    <label class="text-secondary text-sm">Version suffix:</label>
                    <select v-model="store.singleFileSuffix" class="bg-slate-900 border border-slate-700 rounded-lg px-4 py-2 text-slate-100 focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 transition-all w-40">
                        <option value="KH">KH version</option>
                        <option value="VN">VN version</option>
                    </select>
                </div>

                <!-- Ignore missing description check -->
                <div class="mt-4 flex items-center gap-2">
                    <input type="checkbox" id="ignore-missing-desc" v-model="store.ignoreMissingDescription" class="rounded bg-slate-950 border-slate-700 text-blue-500 focus:ring-blue-500 focus:ring-offset-slate-900 w-4 h-4 cursor-pointer" />
                    <label for="ignore-missing-desc" class="text-sm text-slate-300 select-none cursor-pointer flex items-center gap-1.5">
                        ⚠️ Ignore missing description error (Bypass DES column checks)
                    </label>
                </div>

                <div v-if="store.selectedFiles.length > 2" class="status-box error mt-2">
                    ⚠️ Maximum 2 files allowed. Only the first 2 will be used.
                </div>
            </div>
            
            <button class="w-full px-6 py-3 mt-4 bg-gradient-to-r from-blue-500 to-blue-600 hover:from-blue-400 hover:to-blue-500 text-white font-medium rounded-full shadow-lg shadow-blue-500/30 transition-all transform hover:-translate-y-0.5 disabled:opacity-50 disabled:cursor-not-allowed disabled:transform-none" @click="store.analyzeFiles" :disabled="store.selectedFiles.length === 0 || store.isProcessing">
                {{ store.isProcessing ? 'Analyzing...' : 'Analyze & Extract' }}
            </button>
            
             <div v-if="store.statusMessage" :class="['status-box', store.statusType]">
                <div class="flex flex-col gap-2">
                    <span class="break-words text-xs whitespace-pre-wrap leading-relaxed">{{ store.statusMessage }}</span>
                    <button v-if="store.statusType === 'error' && (store.statusMessage.includes('Missing Description') || store.statusMessage.includes('description') || store.statusMessage.includes('ValueError'))" 
                            @click="store.forceAnalyze" 
                            class="mt-2 px-4 py-2.5 bg-amber-500 hover:bg-amber-400 text-slate-950 text-xs font-bold rounded-lg transition-all self-start flex items-center gap-1.5 shadow-md transform hover:-translate-y-0.5 cursor-pointer">
                        ⚠️ Force Analyze & Bypass This Error
                    </button>
                </div>
            </div>
        </div>
    `,
    setup() {
        const store = useTemplateExtractorStore();
        return { store };
    }
};
