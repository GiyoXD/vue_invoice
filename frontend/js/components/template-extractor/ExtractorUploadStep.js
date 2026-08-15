import { ref, computed, onMounted } from 'vue';
import { useTemplateExtractorStore } from '../../stores/templateExtractorStore.js';

export default {
    name: 'ExtractorUploadStep',
    template: `
        <div class="bg-slate-800 border border-slate-700/80 rounded-xl mb-6 text-slate-200 shadow-sm overflow-hidden" v-if="store.currentStep === 1">
            <!-- Header & Working Folder Bar -->
            <div class="p-5 pb-4 border-b border-slate-700/80 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                <div>
                    <h2 class="text-base font-bold text-slate-100 tracking-tight">1. Analyze Invoice Source</h2>
                    <p class="text-xs text-slate-400 mt-1">Select an Excel template from the working folder or upload from PC to extract blueprint configuration.</p>
                </div>

                <!-- Folder Path & Change Toggle -->
                <div class="flex items-center gap-2 bg-slate-900/90 border border-slate-700/70 rounded-lg px-3 py-1.5 text-xs self-start sm:self-auto">
                    <span class="text-slate-400 font-medium shrink-0">Folder:</span>
                    <div v-if="!isEditingFolder" class="flex items-center gap-2 overflow-hidden">
                        <span class="font-mono text-slate-200 text-[11px] truncate max-w-[180px] sm:max-w-xs bg-slate-800/80 px-2 py-0.5 rounded border border-slate-700/50" :title="store.sourceFolderPath || 'Default uploads folder'">
                            {{ store.sourceFolderPath || 'Loading...' }}
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
                        Folder Files ({{ store.availableFiles.length }})
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
                            v-model="store.fileSearchQuery" 
                            @keyup.enter="handleEnterKey"
                            placeholder="Search files in folder..." 
                            class="w-full px-3 py-2 text-xs sm:text-sm bg-slate-900/90 border border-slate-700/80 rounded-lg text-slate-100 placeholder-slate-500 focus:border-blue-500 focus:ring-1 focus:ring-blue-500 focus:outline-none font-mono transition-all"
                        />
                        <button 
                            v-if="store.fileSearchQuery" 
                            @click="clearSearch" 
                            type="button"
                            class="absolute inset-y-0 right-0 flex items-center pr-3 text-slate-400 hover:text-slate-200 text-xs font-mono"
                            title="Clear search">
                            Clear
                        </button>
                    </div>

                    <!-- Scrollable List Box -->
                    <div class="border border-slate-700/80 rounded-lg bg-slate-900/90 max-h-60 overflow-y-auto p-1.5 space-y-1 custom-scrollbar">
                        <div v-if="store.isLoadingFiles" class="p-8 text-center text-xs text-slate-400">
                            Scanning files in working folder...
                        </div>
                        <div v-else-if="filteredFiles.length === 0" class="p-8 text-center text-xs text-slate-400">
                            No files matching "<span class="text-slate-200 font-mono">{{ store.fileSearchQuery }}</span>" found.
                        </div>
                        <div 
                            v-else
                            v-for="fileItem in filteredFiles" 
                            :key="fileItem.filename"
                            @click="selectFileRow(fileItem)"
                            @dblclick="analyzeFileDirectly(fileItem)"
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

                        <!-- File selected card in dropzone -->
                        <div v-if="store.rawFile" class="w-full max-w-lg bg-slate-800 border border-slate-700 rounded-lg p-4 flex flex-col sm:flex-row items-center justify-between gap-3 shadow-md" @click.stop>
                            <div class="min-w-0 text-left w-full sm:w-auto">
                                <p class="text-xs font-semibold text-slate-100 font-mono truncate" :title="store.rawFile.name">
                                    {{ store.rawFile.name }}
                                </p>
                                <p class="text-[11px] text-slate-400 font-mono mt-0.5">
                                    {{ formatSize(store.rawFile.size) }}
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
                                Supports .xlsx, .xls spreadsheets
                            </p>
                        </div>
                    </div>
                </div>

                <!-- Below Tabs: Selection Info & Configuration Controls -->
                <div v-if="activeFilename" class="bg-slate-900/60 border border-slate-700/80 rounded-xl p-5 space-y-4 mt-4">
                    <div class="flex items-center justify-between pb-3 border-b border-slate-700/60">
                        <div class="flex items-center gap-2.5">
                            <span class="text-xs font-bold uppercase tracking-wider text-slate-400">Selected Source File</span>
                            <span class="px-2 py-0.5 rounded-full text-xs font-mono font-bold bg-blue-500/10 text-blue-400 border border-blue-500/20">
                                {{ store.selectedExistingFile ? 'Working Folder' : 'Uploaded File' }}
                            </span>
                        </div>
                    </div>

                    <!-- File Card -->
                    <div class="flex items-center justify-between gap-3 p-3.5 bg-slate-800/90 border border-slate-700/80 rounded-xl">
                        <div class="min-w-0">
                            <p class="font-mono text-sm font-semibold text-slate-100 truncate" :title="activeFilename">
                                {{ activeFilename }}
                            </p>
                            <p class="text-[11px] text-slate-400 font-mono mt-0.5">
                                {{ selectedFileSize ? formatSize(selectedFileSize) : 'Excel Spreadsheet' }}
                            </p>
                        </div>
                        <div class="shrink-0">
                            <span class="px-2.5 py-1 rounded-md text-xs font-bold border bg-blue-500/15 text-blue-400 border-blue-500/30 tracking-wide uppercase">
                                {{ store.singleFileSuffix }} Version
                            </span>
                        </div>
                    </div>

                    <!-- Single-file suffix selector -->
                    <div class="p-3.5 bg-slate-800/70 border border-slate-700/60 rounded-xl flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                        <div class="flex items-center gap-2">
                            <span class="text-xs text-slate-400 font-medium">Single file mode:</span>
                            <span class="text-xs text-slate-300">Choose template target version</span>
                        </div>
                        <div class="flex items-center gap-2 self-start sm:self-auto">
                            <label class="text-xs text-slate-400 font-semibold uppercase tracking-wider shrink-0">Version:</label>
                            <select v-model="store.singleFileSuffix" class="bg-slate-900 border border-slate-700 rounded-lg px-3 py-1.5 text-xs font-semibold text-slate-100 focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 transition-all cursor-pointer">
                                <option value="KH">KH Version (Cambodia)</option>
                                <option value="VN">VN Version (Vietnam)</option>
                            </select>
                        </div>
                    </div>

                    <!-- Ignore missing description check -->
                    <div>
                        <label for="ignore-missing-desc" class="flex items-center gap-2.5 p-3 bg-slate-900/50 hover:bg-slate-900/80 border border-slate-700/70 hover:border-slate-600 rounded-xl cursor-pointer transition-colors group">
                            <input type="checkbox" id="ignore-missing-desc" v-model="store.ignoreMissingDescription" class="rounded bg-slate-950 border-slate-700 text-blue-500 focus:ring-blue-500 focus:ring-offset-slate-900 w-4 h-4 cursor-pointer" />
                            <span class="text-xs text-slate-300 group-hover:text-slate-200 select-none">
                                Ignore missing description error <span class="text-slate-400">(Bypass DES column validation checks)</span>
                            </span>
                        </label>
                    </div>

                    <div v-if="store.uploadWarning" class="status-box warning mt-2">
                        {{ store.uploadWarning }}
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
                </div>

                <!-- Right: Action Buttons -->
                <div class="flex items-center flex-wrap gap-2">
                    <!-- Analyze Button -->
                    <button 
                        type="button" 
                        @click="store.analyzeFiles" 
                        :disabled="!activeFilename || store.isProcessing"
                        class="bg-blue-600 hover:bg-blue-500 active:bg-blue-700 text-white px-4 py-2 rounded-lg text-xs font-semibold shadow-md transition-all disabled:opacity-40 disabled:cursor-not-allowed">
                        {{ store.isProcessing ? 'Analyzing Template...' : 'Analyze & Extract Template' }}
                    </button>

                    <!-- Open in Excel -->
                    <button 
                        type="button"
                        @click="store.openSelectedInExcel(activeFilename)"
                        :disabled="!store.selectedExistingFile || store.rawFile || store.isOpeningFile"
                        class="bg-slate-800 hover:bg-slate-700 text-slate-200 px-3.5 py-2 rounded-lg text-xs font-semibold border border-slate-700 transition-colors disabled:opacity-40 disabled:cursor-not-allowed"
                        title="Open file directly in Excel">
                        {{ store.isOpeningFile ? 'Opening...' : 'Open in Excel' }}
                    </button>

                    <!-- Refresh Files -->
                    <button 
                        type="button"
                        @click="store.fetchSourceFiles"
                        :disabled="store.isLoadingFiles"
                        class="bg-slate-800 hover:bg-slate-700 text-slate-200 px-3.5 py-2 rounded-lg text-xs font-semibold border border-slate-700 transition-colors disabled:opacity-40 disabled:cursor-not-allowed"
                        title="Scan folder for new files">
                        {{ store.isLoadingFiles ? 'Scanning...' : 'Refresh' }}
                    </button>

                    <!-- Reset -->
                    <button 
                        type="button"
                        @click="store.resetFlow"
                        class="bg-rose-950/30 hover:bg-rose-950/60 text-rose-300 hover:text-rose-200 px-3 py-2 rounded-lg text-xs font-semibold border border-rose-900/50 hover:border-rose-800/80 transition-colors"
                        title="Clear selection and reset state">
                        Reset
                    </button>
                </div>
            </div>

            <!-- Status & Error Box -->
            <div v-if="store.statusMessage" :class="['status-box', store.statusType, 'mx-5 mb-5 mt-4']">
                <div class="flex flex-col gap-2 w-full">
                    <div class="flex items-start gap-2">
                        <span class="break-words text-xs whitespace-pre-wrap leading-relaxed">{{ store.statusMessage }}</span>
                    </div>
                    <button v-if="store.statusType === 'error' && (store.statusMessage.includes('Missing Description') || store.statusMessage.includes('description') || store.statusMessage.includes('ValueError'))" 
                            @click="store.forceAnalyze" 
                            type="button"
                            class="mt-2 px-4 py-2 bg-amber-500 hover:bg-amber-400 text-slate-950 text-xs font-bold rounded-lg transition-all self-start flex items-center gap-1.5 shadow-md transform hover:-translate-y-0.5 cursor-pointer">
                        Force Analyze & Bypass This Error
                    </button>
                </div>
            </div>
        </div>
    `,
    setup() {
        const store = useTemplateExtractorStore();
        const isEditingFolder = ref(false);
        const tempFolderPath = ref('');
        const activeTab = ref('upload');
        const fileInputRef = ref(null);
        const isDragging = ref(false);

        const filteredFiles = computed(() => {
            const q = (store.fileSearchQuery || '').trim().toLowerCase();
            if (!q) return store.availableFiles;
            return store.availableFiles.filter(f => f.filename.toLowerCase().includes(q));
        });

        const activeFilename = computed(() => {
            return store.selectedExistingFile?.filename || store.selectedFile?.name || store.rawFile?.name || '';
        });

        const selectedFileSize = computed(() => {
            return store.selectedExistingFile?.size_bytes || store.selectedFile?.size || store.rawFile?.size || null;
        });

        const isSelected = (fileItem) => {
            return store.selectedExistingFile?.filename === fileItem.filename;
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
            tempFolderPath.value = store.sourceFolderPath;
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
            store.selectExistingFile(fileItem);
        };

        const analyzeFileDirectly = (fileItem) => {
            if (!fileItem) return;
            store.selectExistingFile(fileItem);
            store.analyzeFiles();
        };

        const handleEnterKey = () => {
            const q = (store.fileSearchQuery || '').trim();
            if (q) {
                if (filteredFiles.value.length > 0) {
                    store.selectExistingFile(filteredFiles.value[0]);
                    store.analyzeFiles();
                }
            } else if (store.selectedExistingFile) {
                store.analyzeFiles();
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
                store.setRawFile(file);
            }
        };

        const onFileChange = (event) => {
            const file = event.target.files?.[0];
            if (file) {
                store.setRawFile(file);
            }
        };

        const clearSearch = () => {
            store.fileSearchQuery = '';
        };

        onMounted(async () => {
            if (!store.sourceFolderPath) {
                await store.fetchSourceFolder();
            }
            if (store.availableFiles.length === 0) {
                await store.fetchSourceFiles();
            }
        });

        return {
            store,
            isEditingFolder,
            tempFolderPath,
            activeTab,
            fileInputRef,
            isDragging,
            filteredFiles,
            activeFilename,
            selectedFileSize,
            isSelected,
            formatSize,
            formatDate,
            startEditingFolder,
            cancelEditingFolder,
            saveFolder,
            selectFileRow,
            analyzeFileDirectly,
            handleEnterKey,
            triggerFileInput,
            onDrop,
            onFileChange,
            clearSearch
        };
    }
};
