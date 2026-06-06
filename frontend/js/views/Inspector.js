import { onMounted } from 'vue';
import { useInspectorStore } from '../stores/inspectorStore.js';
import HistorySidebar from '../components/inspector/HistorySidebar.js';
import TotalsSummary from '../components/inspector/TotalsSummary.js';
import DataTable from '../components/inspector/DataTable.js';

export default {
    name: 'InspectorView',
    components: {
        HistorySidebar,
        TotalsSummary,
        DataTable
    },
    template: `
        <div class="max-w-[1600px] mx-auto py-8 fade-in h-screen flex flex-col">
            <h1 class="text-4xl font-extrabold tracking-tight text-transparent bg-clip-text bg-gradient-to-r from-blue-400 to-emerald-400 drop-shadow-md flex-shrink-0">Data Inspector & Registry</h1>
            
            <div class="flex gap-6 mt-8 flex-grow min-h-0">
                <history-sidebar></history-sidebar>

                <div class="flex-1 bg-slate-800/80 backdrop-blur-md border border-slate-700/50 shadow-2xl rounded-2xl p-6 flex flex-col min-w-0">
                    <totals-summary></totals-summary>

                    <div v-if="!store.inspectorData" class="text-center p-8 text-slate-500">
                        <p>Select a run from the left 👈 to inspect invoice data.</p>
                    </div>     

                    <data-table></data-table>
                </div>
            </div>
        </div>
    `,
    setup() {
        const store = useInspectorStore();

        onMounted(() => {
            store.fetchHistory(true);
        });

        return { store };
    }
};
