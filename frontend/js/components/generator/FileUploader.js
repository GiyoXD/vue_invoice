import { useGeneratorStore } from '../../stores/generatorStore.js';
import { storeToRefs } from 'pinia';
import { ref, computed, watch, onMounted } from 'vue';

export default {
    name: 'FileUploader',
    template: `
        <div class="bg-slate-800 border border-slate-700/80 rounded-xl mb-6 text-slate-200 shadow-sm overflow-hidden">
            <!-- Header & Working Folder Bar -->
            <div class="p-5 pb-4 border-b border-slate-700/80 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                <div>
                    <h2 class="text-base font-bold text-slate-100 tracking-tight">1. Source Data</h2>
                    <p class="text-xs text-slate-400 mt-1">Select an Excel file from the working folder or upload a new one.</p>
                </div>

                <!-- Folder Path & Change Toggle -->
                <div class="flex items-center gap-2 bg-slate-900/90 border border-slate-700/70 rounded-lg px-3 py-1.5 text-xs self-start sm:self-auto">
                    <span class="text-slate-400 font-medium shrink-0">Folder:</span>
                    <div v-if="!isEditingFolder" class="flex items-center gap-2 overflow-hidden">
                        <span class="font-mono text-slate-200 text-[11px] truncate max-w-[180px] sm:max-w-xs bg-slate-800/80 px-2 py-0.5 rounded border border-slate-700/50" :title="sourceFolderPath || 'Default uploads folder'">
                            {{ sourceFolderPath || 'Loading...' }}
                        </span>
                        <button 
                            type="button"
                            @click="startEditingFolder" 
                            class="text-slate-400 hover:text-blue-400 px-2 py-0.5 rounded hover:bg-slate-800 transition-colors text-[11px] font-medium" 
                            title="Change Source Folder">
                            Edit
                        </button>
                    </div>
                    <div v-else class="flex items-center gap-1.5">
                        <input 
                            type="text" 
                            v-model="tempFolderPath" 
                            placeholder="C:\\path\\to\\folder" 
                            class="bg-slate-800 border border-blue-500/70 rounded px-2 py-1 text-slate-100 font-mono text-xs w-48 sm:w-64 focus:outline-none focus:ring-1 focus:ring-blue-500" 
                            @keyup.enter="saveFolder" 
                            @keyup.esc="cancelEditingFolder"
                        />
                        <button 
                            type="button" 
                            @click="saveFolder" 
                            class="px-2.5 py-1 bg-blue-600 hover:bg-blue-500 text-white rounded text-xs font-semibold shadow transition-colors">
                            Save
                        </button>
                        <button 
                            type="button" 
                            @click="cancelEditingFolder" 
                            class="px-2 py-1 bg-slate-800 hover:bg-slate-700 text-slate-300 border border-slate-600/60 rounded text-xs transition-colors">
                            Cancel
                        </button>
                    </div>
                </div>
            </div>

            <!-- Main Content Area: Tabs + Selector/Upload -->
            <div class="p-5 space-y-4">
                <!-- Segmented Tab Bar -->
                <div class="inline-flex items-center p-1 bg-slate-900/90 border border-slate-700/80 rounded-lg">
                    <button 
                        type="button" 
                        @click="activeTab = 'selector'" 
                        :class="[
                            'px-3.5 py-1.5 text-xs font-semibold rounded-md transition-all', 
                            activeTab === 'selector' 
                                ? 'bg-blue-600 text-white' 
                                : 'text-slate-400 hover:text-slate-200 hover:bg-slate-800/60'
                        ]">
                        Folder Files ({{ availableFiles.length }})
                    </button>
                    <button 
                        type="button" 
                        @click="activeTab = 'upload'" 
                        :class="[
                            'px-3.5 py-1.5 text-xs font-semibold rounded-md transition-all', 
                            activeTab === 'upload' 
                                ? 'bg-blue-600 text-white' 
                                : 'text-slate-400 hover:text-slate-200 hover:bg-slate-800/60'
                        ]">
                        Upload from PC
                    </button>
                </div>

                <!-- TAB 1: Folder Files (Selector) -->
                <div v-if="activeTab === 'selector'" class="space-y-3">
                    <!-- Search Bar with Clear Button -->
                    <div class="relative">
                        <input 
                            type="text" 
                            v-model="fileSearchQuery" 
                            @keyup.enter="handleEnterKey"
                            placeholder="Search files in folder..." 
                            class="w-full px-3 py-2 text-xs sm:text-sm bg-slate-900/90 border border-slate-700/80 rounded-lg text-slate-100 placeholder-slate-500 focus:border-blue-500 focus:ring-1 focus:ring-blue-500 focus:outline-none font-mono transition-all"
                        />
                        <button 
                            v-if="fileSearchQuery" 
                            @click="clearSearch" 
                            type="button"
                            class="absolute inset-y-0 right-0 flex items-center pr-3 text-slate-400 hover:text-slate-200 text-xs font-mono"
                            title="Clear search">
                            Clear
                        </button>
                    </div>

                    <!-- Scrollable List Box -->
                    <div class="border border-slate-700/80 rounded-lg bg-slate-900/90 max-h-60 overflow-y-auto p-1.5 space-y-1 custom-scrollbar">
                        <div v-if="isLoadingFiles" class="p-8 text-center text-xs text-slate-400">
                            Scanning files in working folder...
                        </div>
                        <div v-else-if="filteredFiles.length === 0" class="p-8 text-center text-xs text-slate-400">
                            No files matching "<span class="text-slate-200 font-mono">{{ fileSearchQuery }}</span>" found.
                        </div>
                        <div 
                            v-else
                            v-for="fileItem in filteredFiles" 
                            :key="fileItem.filename"
                            @click="selectFileRow(fileItem)"
                            @dblclick="processFileDirectly(fileItem.filename)"
                            :class="[
                                'px-3 py-2 cursor-pointer flex items-center justify-between text-xs rounded-md transition-all select-none border',
                                isSelected(fileItem) 
                                    ? 'bg-blue-600/15 border-blue-500/50 text-blue-100' 
                                    : 'border-transparent hover:bg-slate-800/80 text-slate-300 hover:border-slate-700/60'
                            ]">
                            <div class="min-w-0 pr-3">
                                <span class="truncate font-mono font-medium block" :title="fileItem.filename">
                                    {{ fileItem.filename }}
                                </span>
                            </div>
                            <div class="flex items-center gap-3 shrink-0 text-[11px] font-mono">
                                <span class="hidden sm:inline text-slate-500 text-[10px]">{{ formatDate(fileItem.modified_at) }}</span>
                                <span class="text-slate-400">
                                    {{ formatSize(fileItem.size_bytes) }}
                                </span>
                            </div>
                        </div>
                    </div>
                </div>

                <!-- TAB 2: Upload Dropzone -->
                <div v-else class="space-y-3">
                    <input 
                        type="file" 
                        ref="fileInputRef"
                        @click="$event.target.value = ''"
                        @change="onFileChange" 
                        accept=".xlsx, .xls, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel" 
                        class="hidden" 
                    />

                    <!-- Dashed Dropzone Box -->
                    <div 
                        @click="triggerFileInput"
                        @dragover.prevent="isDragging = true"
                        @dragenter.prevent="isDragging = true"
                        @dragleave.prevent="isDragging = false"
                        @drop.prevent="onDrop"
                        :class="[
                            'border-2 border-dashed rounded-lg p-8 text-center cursor-pointer transition-all bg-slate-900/40 hover:bg-slate-900/80 flex flex-col items-center justify-center gap-2',
                            isDragging ? 'border-blue-500 bg-blue-900/20' : 'border-slate-700/80 hover:border-blue-500/70'
                        ]">

                        <!-- File selected card -->
                        <div v-if="localUploadedFile" class="w-full max-w-lg bg-slate-800 border border-slate-700 rounded-lg p-4 flex flex-col sm:flex-row items-center justify-between gap-3 shadow-md" @click.stop>
                            <div class="min-w-0 text-left w-full sm:w-auto">
                                <p class="text-xs font-semibold text-slate-100 font-mono truncate" :title="localUploadedFile.name">
                                    {{ localUploadedFile.name }}
                                </p>
                                <p class="text-[11px] text-slate-400 font-mono mt-0.5">
                                    {{ formatSize(localUploadedFile.size) }}
                                </p>
                            </div>
                            <div class="flex items-center gap-2 shrink-0 w-full sm:w-auto justify-end">
                                <button 
                                    type="button" 
                                    @click="triggerFileInput"
                                    class="px-3 py-1.5 text-xs font-medium text-slate-300 hover:text-white bg-slate-700 hover:bg-slate-600 border border-slate-600 rounded-md transition-colors">
                                    Change File
                                </button>
                            </div>
                        </div>

                        <!-- Empty dropzone default view -->
                        <div v-else class="flex flex-col items-center gap-1.5 pointer-events-none py-4">
                            <p class="text-xs sm:text-sm font-medium text-slate-200 pointer-events-none">
                                Click to browse or drag & drop Excel file (.xlsx, .xls)
                            </p>
                            <p class="text-[11px] text-slate-400 pointer-events-none">
                                Files are automatically saved to your working folder
                            </p>
                        </div>
                    </div>
                </div>
            </div>

            <!-- Action Toolbar Footer -->
            <div class="p-4 bg-slate-900/70 border-t border-slate-700/70 rounded-b-xl flex flex-wrap items-center justify-between gap-3">
                <!-- Left: Selected File Info Badge -->
                <div class="flex items-center gap-2 min-w-0 text-xs">
                    <span class="text-slate-400 font-medium shrink-0">Selected:</span>
                    <span v-if="activeFilename" class="font-mono text-slate-100 font-semibold truncate max-w-[200px] sm:max-w-xs bg-slate-800/90 px-2.5 py-1 rounded border border-slate-700/80" :title="activeFilename">
                        {{ activeFilename }}
                    </span>
                    <span v-else class="text-slate-500 italic">None</span>
                    <span v-if="processingComplete" class="shrink-0 px-2.5 py-0.5 rounded bg-emerald-950/80 border border-emerald-600/60 text-emerald-400 text-[10px] font-semibold">
                        Processed
                    </span>
                </div>

                <!-- Right: Action Buttons -->
                <div class="flex items-center flex-wrap gap-2">
                    <!-- Process File -->
                    <button 
                        type="button" 
                        @click="handleProcessAction" 
                        :disabled="!activeFilename || isUploading"
                        class="bg-blue-600 hover:bg-blue-500 active:bg-blue-700 text-white px-4 py-2 rounded-lg text-xs font-semibold transition-all disabled:opacity-40 disabled:cursor-not-allowed">
                        {{ isUploading ? 'Processing...' : 'Process File' }}
                    </button>

                    <!-- Open in Excel -->
                    <button 
                        type="button"
                        @click="openSelectedInExcel(selectedExistingFile ? selectedExistingFile.filename : activeFilename)"
                        :disabled="!selectedExistingFile || hasRawFile || isOpeningFile"
                        class="bg-slate-800 hover:bg-slate-700 text-slate-200 px-3.5 py-2 rounded-lg text-xs font-semibold border border-slate-700 transition-colors disabled:opacity-40 disabled:cursor-not-allowed"
                        title="Open file directly in Excel">
                        {{ isOpeningFile ? 'Opening...' : 'Open in Excel' }}
                    </button>

                    <!-- Refresh Files -->
                    <button 
                        type="button"
                        @click="fetchSourceFiles"
                        :disabled="isLoadingFiles"
                        class="bg-slate-800 hover:bg-slate-700 text-slate-200 px-3.5 py-2 rounded-lg text-xs font-semibold border border-slate-700 transition-colors disabled:opacity-40 disabled:cursor-not-allowed"
                        title="Scan folder for new files">
                        {{ isLoadingFiles ? 'Scanning...' : 'Refresh' }}
                    </button>

                    <!-- Reset -->
                    <button 
                        type="button"
                        @click="handleReset"
                        class="bg-rose-950/30 hover:bg-rose-950/60 text-rose-300 hover:text-rose-200 px-3 py-2 rounded-lg text-xs font-semibold border border-rose-900/50 hover:border-rose-800/80 transition-colors"
                        title="Clear selection and reset state">
                        Reset
                    </button>
                </div>
            </div>

            <!-- Status Box -->
            <div v-if="uploadStatus && !uploadError" :class="['status-box', uploadStatus.type, 'mx-5 mb-5']">
                {{ uploadStatus.message }}
            </div>

            <!-- NORMALIZATION WARNINGS PANEL -->
            <div v-if="validationWarnings && validationWarnings.length > 0" class="warning-panel mx-5 mb-5">
                <div class="warning-header mb-2">
                    <h3 class="m-0 text-amber-400 font-bold text-sm">Data Auto-Correction Notices</h3>
                </div>
                <ul class="m-0 pl-5 text-slate-200 text-xs">
                    <li v-for="(msg, idx) in validationWarnings" :key="idx" class="mb-1">
                        {{ msg }}
                    </li>
                </ul>
            </div>

            <!-- ERROR PANEL FOR UPLOAD -->
            <div v-if="uploadError" class="error-panel mx-5 mb-5">
                <div class="error-header">
                    <h3 class="text-sm font-bold text-rose-400">Upload / Processing Failed</h3>
                </div>
                <span v-if="uploadError.step" class="error-step text-xs">{{ uploadError.step }}</span>
                <div class="error-message text-xs" style="white-space: pre-wrap;">{{ uploadError.message }}</div>
                
                <div v-if="uploadError.traceback" 
                     class="traceback-toggle text-xs cursor-pointer py-1 text-slate-400 hover:text-slate-300" 
                     :class="{ open: showUploadTraceback }"
                     @click="showUploadTraceback = !showUploadTraceback">
                     <span>{{ showUploadTraceback ? 'Hide Technical Details' : 'View Technical Details' }}</span>
                </div>
                <div v-if="uploadError.traceback" class="traceback-content text-xs" :class="{ open: showUploadTraceback }">
                    <pre class="bg-slate-950 p-3 rounded text-slate-300 overflow-x-auto">{{ uploadError.traceback }}</pre>
                </div>
                
                <div class="error-actions mt-3 flex flex-wrap gap-2">
                    <button class="btn-retry text-xs px-3 py-1.5 rounded" @click="retryUpload">Try Again</button>
                    <button v-if="uploadError.message && (uploadError.message.includes('Weight Integrity Error') || uploadError.message.includes('Tare mismatch') || uploadError.message.includes('Gross <= Net') || uploadError.message.includes('Missing Gross Weight') || uploadError.message.includes('Missing Net Weight'))" class="px-3 py-1.5 bg-amber-600 hover:bg-amber-500 text-white font-medium rounded text-xs transition-colors" @click="ignoreTareAndRetry">Bypass Weight Check</button>
                    <button v-if="uploadError.message && uploadError.message.includes('CBM')" class="px-3 py-1.5 bg-amber-600 hover:bg-amber-500 text-white font-medium rounded text-xs transition-colors" @click="ignoreCbmAndRetry">Bypass CBM Check</button>
                    <button class="btn-copy-error text-xs px-3 py-1.5 rounded" @click="copyError(uploadError)">Copy Error</button>
                </div>
            </div>
        </div>
    `,
    setup() {
        const store = useGeneratorStore();
        const {
            selectedFile,
            hasRawFile,
            isUploading,
            uploadStatus,
            uploadError,
            showUploadTraceback,
            validationWarnings,
            sourceFolderPath,
            availableFiles,
            isLoadingFiles,
            fileSearchQuery,
            selectedExistingFile,
            isOpeningFile,
            processingComplete
        } = storeToRefs(store);

        const isEditingFolder = ref(false);
        const tempFolderPath = ref('');
        const activeTab = ref('upload');
        const fileInputRef = ref(null);
        const isDragging = ref(false);
        const localUploadedFile = ref(null);

        const filteredFiles = computed(() => {
            const q = (fileSearchQuery.value || '').trim().toLowerCase();
            if (!q) return availableFiles.value;
            return availableFiles.value.filter(f => f.filename.toLowerCase().includes(q));
        });

        const activeFilename = computed(() => {
            return selectedExistingFile.value?.filename || selectedFile.value?.name || '';
        });

        const isSelected = (fileItem) => {
            return selectedExistingFile.value?.filename === fileItem.filename;
        };

        const formatSize = (bytes) => {
            if (!bytes && bytes !== 0) return '';
            if (bytes < 1024) return `${bytes} B`;
            if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
            return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
        };

        const formatDate = (isoStr) => {
            if (!isoStr) return '';
            try {
                const d = new Date(isoStr);
                if (isNaN(d.getTime())) return isoStr;
                return d.toLocaleString(undefined, { 
                    month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit'
                });
            } catch {
                return isoStr;
            }
        };

        const startEditingFolder = () => {
            tempFolderPath.value = sourceFolderPath.value;
            isEditingFolder.value = true;
        };

        const cancelEditingFolder = () => {
            isEditingFolder.value = false;
        };

        const saveFolder = async () => {
            if (!tempFolderPath.value.trim()) return;
            const success = await store.updateSourceFolder(tempFolderPath.value);
            if (success) {
                isEditingFolder.value = false;
            }
        };

        const selectFileRow = (fileItem) => {
            localUploadedFile.value = null;
            store.setRawFile(null);
            selectedExistingFile.value = fileItem;
            selectedFile.value = { name: fileItem.filename };
        };

        const processFileDirectly = (filename) => {
            if (!filename) return;
            localUploadedFile.value = null;
            store.processExistingFile(filename);
        };

        const handleEnterKey = () => {
            const q = (fileSearchQuery.value || '').trim();
            if (q) {
                if (filteredFiles.value.length > 0) {
                    store.processExistingFile(filteredFiles.value[0].filename);
                } else {
                    store.processExistingFile(q);
                }
            } else if (selectedExistingFile.value) {
                store.processExistingFile(selectedExistingFile.value.filename);
            }
        };

        const triggerFileInput = () => {
            fileInputRef.value?.click();
        };

        const onDrop = (event) => {
            isDragging.value = false;
            const file = event.dataTransfer?.files?.[0];
            if (file) {
                const name = (file.name || '').toLowerCase();
                if (!name.endsWith('.xlsx') && !name.endsWith('.xls')) {
                    return;
                }
                localUploadedFile.value = file;
                store.setRawFile(file);
            }
        };

        const onFileChange = (event) => {
            const file = event.target.files?.[0];
            if (file) {
                localUploadedFile.value = file;
                store.setRawFile(file);
            }
        };

        const handleProcessAction = () => {
            if (hasRawFile.value || localUploadedFile.value) {
                store.uploadFile();
            } else if (selectedExistingFile.value) {
                store.processExistingFile(selectedExistingFile.value.filename);
            } else if (activeFilename.value) {
                store.processExistingFile(activeFilename.value);
            }
        };

        const handleReset = () => {
            localUploadedFile.value = null;
            store.resetGeneratorState();
        };

        const clearSearch = () => {
            fileSearchQuery.value = '';
        };

        watch(hasRawFile, (val) => {
            if (!val) {
                localUploadedFile.value = null;
            }
        });

        onMounted(async () => {
            await store.fetchSourceFolder();
            await store.fetchSourceFiles();
        });

        return {
            selectedFile,
            hasRawFile,
            isUploading,
            uploadStatus,
            uploadError,
            showUploadTraceback,
            validationWarnings,
            sourceFolderPath,
            availableFiles,
            isLoadingFiles,
            fileSearchQuery,
            selectedExistingFile,
            isOpeningFile,
            processingComplete,
            isEditingFolder,
            tempFolderPath,
            activeTab,
            fileInputRef,
            isDragging,
            localUploadedFile,
            filteredFiles,
            activeFilename,
            isSelected,
            formatSize,
            formatDate,
            startEditingFolder,
            cancelEditingFolder,
            saveFolder,
            selectFileRow,
            processFileDirectly,
            handleEnterKey,
            triggerFileInput,
            onDrop,
            onFileChange,
            handleProcessAction,
            handleReset,
            clearSearch,
            openSelectedInExcel: store.openSelectedInExcel,
            fetchSourceFiles: store.fetchSourceFiles,
            resetGeneratorState: store.resetGeneratorState,
            retryUpload: store.retryUpload,
            ignoreTareAndRetry: store.ignoreTareAndRetry,
            ignoreCbmAndRetry: store.ignoreCbmAndRetry,
            copyError: store.copyError
        };
    }
};
