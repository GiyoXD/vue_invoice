import { useInspectorStore } from '../../stores/inspectorStore.js';

export default {
    name: 'HistorySidebar',
    template: `
        <div class="w-[300px] bg-slate-800/80 backdrop-blur-md border border-slate-700/50 shadow-2xl rounded-2xl p-6 flex flex-col flex-shrink-0">
            <h3 class="text-xl font-bold text-slate-100 mb-4 mt-0">Recent Runs</h3>
            <button class="w-full px-4 py-2 bg-slate-700 hover:bg-slate-600 text-slate-200 rounded-lg shadow-sm transition-colors mt-4 flex-shrink-0" @click="store.fetchHistory(false)">Refresh List</button>
            <div v-if="store.historyList.length === 0" class="text-secondary text-sm mt-4">No history found.</div>
            <div class="flex-1 overflow-y-auto flex flex-col gap-2 pr-2 mt-4">
                <div v-for="run in store.historyList" :key="run.filename" 
                     class="p-3 bg-slate-900/50 border border-slate-700 rounded-lg cursor-pointer transition-all hover:border-blue-500 hover:bg-slate-800" :class="run.type" @click="store.loadHistoryItem(run)">
                    <div class="text-emerald-400 font-bold text-sm mb-1 truncate drop-shadow-sm" :title="run.output_file">{{ run.output_file }}</div>
                    <div class="text-slate-400 text-xs mb-2">{{ formatTime(run.timestamp) }}</div>
                    <div class="text-slate-500 text-xs flex items-center">
                        <span class="text-emerald-500 font-bold mr-1">{{ run.item_count }}</span> items • {{ run.status }}
                    </div>
                </div>
            </div>
        </div>
    `,
    setup() {
        const store = useInspectorStore();
        
        const formatTime = (ts) => {
            if (!ts) return '';
            return new Date(ts).toLocaleString();
        };

        return { store, formatTime };
    }
};
