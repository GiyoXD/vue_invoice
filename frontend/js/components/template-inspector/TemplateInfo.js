import { useTemplateInspectorStore } from '../../stores/templateInspectorStore.js';
import { ref, watch, computed } from 'vue';

export default {
    name: 'TemplateInfo',
    template: `
        <div class="template-info-section">
            <div class="flex items-center justify-between p-4 bg-blue-500/10 border border-blue-500/20 rounded-xl mb-4 flex-shrink-0">
                <div>
                    <strong class="text-blue-400">Viewing:</strong> <span class="text-slate-200">{{ store.selectedTemplateName }}</span> <br>
                    <span class="text-sm text-slate-400 opacity-80">Source: {{ store.currentTemplateFingerprint?.source_file }}</span>
                </div>
                <button class="px-4 py-2 bg-red-500 hover:bg-red-400 text-white font-medium rounded-lg shadow-sm transition-colors text-sm" @click="store.deleteTemplate" title="Delete Template">
                    Delete Template
                </button>
            </div>

            <!-- Client Notes Section -->
            <div class="mb-4 bg-slate-900/50 border border-slate-700 rounded-xl p-4">
                <div class="flex justify-between items-center mb-2">
                    <h4 class="m-0 text-sm text-slate-400 flex items-center gap-2">
                        <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" class="lucide lucide-notebook-pen"><path d="M11 2H9a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2v-4"/><path d="m16 2 4 4-8 8H8v-4l8-8Z"/><path d="M15 5 19 9"/></svg>
                        Client Notes / Remarks
                    </h4>
                    <button v-if="!isEditingNotes" class="px-3 py-1 bg-slate-700 hover:bg-slate-600 text-slate-200 rounded text-xs transition-colors" @click="isEditingNotes = true">Edit</button>
                    <div v-else class="flex gap-1">
                        <button class="px-3 py-1 bg-slate-600 hover:bg-slate-500 text-white rounded text-xs transition-colors" @click="cancelEditNotes">Cancel</button>
                        <button class="px-3 py-1 bg-blue-600 hover:bg-blue-500 text-white rounded text-xs transition-colors" @click="saveNotes" :disabled="isSavingNotes">
                            {{ isSavingNotes ? 'Saving...' : 'Save' }}
                        </button>
                    </div>
                </div>
                <div v-if="!isEditingNotes">
                    <div v-if="store.templateNotes" class="text-sm text-primary whitespace-pre-wrap leading-relaxed">{{ store.templateNotes }}</div>
                    <div v-else class="text-sm text-muted italic">No notes for this client yet. Click Edit to add.</div>
                </div>
                <div v-else>
                    <textarea v-model="editingNotesText" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-4 py-2 text-slate-100 focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 transition-all text-sm min-h-[100px]" placeholder="Enter things to remember for this client..."></textarea>
                </div>
            </div>

            <!-- Client Profile Section -->
            <div v-if="store.clientProfile" class="mb-4 bg-slate-900/50 border border-slate-700 rounded-xl p-4">
                <h4 class="mb-3 text-sm text-slate-400 flex items-center gap-2">
                    <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/></svg>
                    Client Profile
                </h4>
                <div style="display: grid; grid-template-columns: 120px 1fr; gap: 0.5rem 1rem;" class="text-sm">
                    <div class="text-muted font-medium">Company:</div>
                    <div class="text-primary">{{ store.clientProfile.fullname || '—' }}</div>

                    <div class="text-muted font-medium">Address:</div>
                    <div class="text-primary whitespace-pre-line">{{ store.clientProfile.address || '—' }}</div>

                    <div class="text-muted font-medium">Contact:</div>
                    <div class="text-primary whitespace-pre-line">{{ store.clientProfile.contact || '—' }}</div>

                    <div class="text-muted font-medium">Shipping:</div>
                    <div class="text-primary">{{ store.clientProfile.shipping || '—' }}</div>
                </div>
            </div>

            <!-- Table Information Section -->
            <div v-if="store.currentTemplate && store.currentTemplate.table_info" class="mb-4 bg-slate-900/50 border border-slate-700 rounded-xl p-4">
                <h4 class="mb-2 text-sm text-slate-400 flex items-center gap-2">
                    <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" class="lucide lucide-table-properties"><path d="M15 2H9a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V4a2 2 0 0 0-2-2Z"/><path d="M9 10h12"/><path d="M9 14h12"/><path d="M9 18h12"/><path d="M9 6h12"/><path d="M11 2v20"/></svg>
                    Table Information
                </h4>
                <div style="display: grid; grid-template-columns: auto 1fr; gap: 0.5rem 1rem;" class="text-sm">
                    <div class="text-muted font-medium">Fallback Desc (Standard):</div>
                    <div class="text-primary">{{ store.currentTemplate.table_info.fallback_description?.standard || 'None' }}</div>

                    <div class="text-muted font-medium">Fallback Desc (DAF):</div>
                    <div class="text-primary">{{ store.currentTemplate.table_info.fallback_description?.daf || 'None' }}</div>

                    <div class="text-muted font-medium">HS Code:</div>
                    <div class="text-primary">{{ store.currentTemplate.table_info.hs_code || 'None' }}</div>
                </div>
            </div>
        </div>
    `,
    setup() {
        const store = useTemplateInspectorStore();
        const isEditingNotes = ref(false);
        const isSavingNotes = ref(false);
        const editingNotesText = ref("");

        watch(() => store.currentTemplate, (newVal) => {
            if (newVal) {
                editingNotesText.value = newVal.notes || "";
                isEditingNotes.value = false;
            }
        }, { immediate: true });

        const saveNotes = async () => {
            isSavingNotes.value = true;
            const success = await store.saveNotes(editingNotesText.value);
            if (success) {
                isEditingNotes.value = false;
            }
            isSavingNotes.value = false;
        };

        const cancelEditNotes = () => {
            isEditingNotes.value = false;
            editingNotesText.value = store.templateNotes;
        };

        return {
            store,
            isEditingNotes,
            isSavingNotes,
            editingNotesText,
            saveNotes,
            cancelEditNotes
        };
    }
};
