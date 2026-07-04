import { useTemplateInspectorStore } from '../../stores/templateInspectorStore.js';
import { computed } from 'vue';

export default {
    name: 'TemplateSidebar',
    template: `
        <div class="w-[300px] bg-slate-800/80 backdrop-blur-md border border-slate-700/50 shadow-2xl rounded-2xl p-6 flex flex-col flex-shrink-0">
            <h3 class="text-xl font-bold text-slate-100 mb-4 mt-0">Available Templates</h3>
            
            <div class="mb-4">
                <input type="text" v-model="store.searchQuery" placeholder="Search templates..." class="w-full bg-slate-900 border border-slate-700 rounded-lg px-4 py-2 text-slate-100 focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 transition-all" />
                <button class="w-full px-4 py-2 bg-slate-700 hover:bg-slate-600 text-slate-200 rounded-lg shadow-sm transition-colors mt-4 flex-shrink-0" @click="store.fetchTemplates">Refresh List</button>
            </div>

            <div v-if="store.filteredTemplates.length === 0" class="text-slate-400 text-base">No templates found.</div>
            <div class="flex-1 overflow-y-auto flex flex-col gap-2 pr-2">
                <div v-for="t in store.filteredTemplates" :key="t.name + (t.bundle_name || '') + (t.source_file || '')" 
                     class="p-3 bg-slate-900/50 border border-slate-700 rounded-lg cursor-pointer transition-all hover:border-blue-500 hover:bg-slate-800" :class="{ 'border-blue-500 bg-slate-800': store.selectedTemplateName === t.name }"
                     @click="store.loadTemplate(t)">
                    <div class="text-emerald-400 font-bold text-sm mb-1 truncate drop-shadow-sm">
                        {{ t.name }}
                        <span v-if="t.bundle_name && t.name !== t.bundle_name" class="text-xs text-red-500 font-bold ml-1">({{ t.bundle_name }})</span>
                    </div>
                    <div class="text-slate-400 text-xs mb-2 truncate">Source: {{ t.source_file }}</div>
                    <div class="text-slate-500 text-xs">Updated: {{ formatTime(t.modified) }}</div>
                </div>
            </div>
        </div>
    `,
    setup() {
        const store = useTemplateInspectorStore();

        const formatTime = (ts) => {
            if (!ts) return '';
            const d = new Date(ts);
            const pad = (n) => String(n).padStart(2, '0');
            return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
        };

        return {
            store,
            formatTime
        };
    }
};
