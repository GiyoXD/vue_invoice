import { ref } from 'vue';

export default {
    name: 'ExportDataView',
    template: `
        <div class="max-w-5xl mx-auto py-8 fade-in">
            <div class="mb-8">
                <h1 class="text-4xl font-extrabold tracking-tight text-transparent bg-clip-text bg-gradient-to-r from-blue-400 to-emerald-400 drop-shadow-md mb-2">Data Export Registry</h1>
                <p class="text-slate-400">Select a time interval to export stored invoice data to CSV.</p>
            </div>

            <div class="bg-slate-800/80 backdrop-blur-md border border-slate-700/50 shadow-2xl rounded-2xl p-8 mb-8">
                <div class="flex flex-col md:flex-row gap-6 items-end">
                    <div class="flex-1 w-full">
                        <label for="start-date">Start Date</label>
                        <input type="date" id="start-date" v-model="startDate" class="w-full mt-2 bg-slate-900 border border-slate-700 rounded-lg px-4 py-2 text-slate-100 focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 transition-all">
                    </div>
                    
                    <div class="flex-1 w-full">
                        <label for="end-date">End Date</label>
                        <input type="date" id="end-date" v-model="endDate" class="w-full mt-2 bg-slate-900 border border-slate-700 rounded-lg px-4 py-2 text-slate-100 focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 transition-all">
                    </div>

                    <div class="w-full md:w-auto">
                        <button @click="exportData" :disabled="exporting" class="w-full px-8 py-2.5 bg-gradient-to-r from-blue-500 to-blue-600 hover:from-blue-400 hover:to-blue-500 text-white font-medium rounded-lg shadow-lg shadow-blue-500/30 transition-all transform hover:-translate-y-0.5 disabled:opacity-50 disabled:cursor-not-allowed disabled:transform-none">
                            <span v-if="!exporting">Export to CSV</span>
                            <span v-else>Exporting...</span>
                        </button>
                    </div>
                </div>

                <div v-if="error" class="error-message text-red-500 mt-4 text-center">
                    {{ error }}
                </div>
                
                <div v-if="success" class="success-message">
                    Export completed successfully.
                </div>
            </div>

            <!-- Data Preview Section (Peek) -->
            <div class="bg-slate-800/80 backdrop-blur-md border border-slate-700/50 shadow-2xl rounded-2xl p-8 mb-8 mt-8">
                <h3 class="text-2xl font-bold text-slate-100 mb-2">Recent Invoices (Peek)</h3>
                <div class="table-container mt-4 overflow-x-auto">
                    <table class="w-full border-collapse text-sm text-slate-300">
                        <thead class="bg-slate-900/50 border-b border-slate-700 text-slate-400">
                            <tr>
                                <th class="p-3 text-left font-medium">Filename</th>
                                <th class="p-3 text-left font-medium">Accepted At</th>
                                <th class="p-3 text-center font-medium">Items</th>
                                <th class="p-3 text-right font-medium">Total SQFT</th>
                                <th class="p-3 text-right font-medium">Total Net</th>
                                <th class="p-3 text-right font-medium">Total Pallets</th>
                                <th class="p-3 text-right font-medium">Total Amount</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr v-for="item in recentInvoices" :key="item.id" class="border-b border-slate-700/50 hover:bg-slate-700/20 transition-colors">
                                <td class="p-3">{{ item.filename }}</td>
                                <td class="p-3">{{ formatDate(item.timestamp) }}</td>
                                <td class="p-3 text-center">{{ item.item_count }}</td>
                                <td class="p-3 text-right font-mono">{{ item.total_sqft?.toLocaleString(undefined, {minimumFractionDigits: 2}) || '0.00' }}</td>
                                <td class="p-3 text-right font-mono">{{ item.total_net?.toLocaleString(undefined, {minimumFractionDigits: 2}) || '0.00' }} kg</td>
                                <td class="p-3 text-right font-mono">{{ item.total_pallets?.toLocaleString(undefined, {minimumFractionDigits: 2}) || '0.00' }}</td>
                                <td class="p-3 text-right font-mono">$ {{ item.total_amount?.toLocaleString(undefined, {minimumFractionDigits: 2}) || '0.00' }}</td>
                            </tr>
                            <tr v-if="recentInvoices.length === 0">
                                <td colspan="7" class="p-8 text-center text-slate-500">No invoices found in registry.</td>
                            </tr>
                        </tbody>
                    </table>
                </div>
                <div class="mt-4 text-right">
                    <button @click="fetchRecentInvoices" class="px-4 py-2 bg-slate-700 hover:bg-slate-600 text-slate-200 rounded-lg shadow-sm transition-colors text-sm">
                        Refresh List
                    </button>
                </div>
            </div>
        </div>
    `,
    setup() {
        const startDate = ref('');
        const endDate = ref('');
        const exporting = ref(false);
        const error = ref(null);
        const success = ref(false);
        const recentInvoices = ref([]);

        const fetchRecentInvoices = async () => {
            try {
                const response = await fetch('/api/registry/list');
                if (response.ok) {
                    recentInvoices.value = await response.json();
                }
            } catch (err) {
                console.error('Failed to fetch recent invoices:', err);
            }
        };

        const formatDate = (dateStr) => {
            if (!dateStr) return '-';
            const d = new Date(dateStr);
            return d.toLocaleString();
        };

        const exportData = async () => {
            exporting.value = true;
            error.value = null;
            success.value = false;

            try {
                let url = '/api/registry/export';
                const params = new URLSearchParams();
                if (startDate.value) params.append('start_date', startDate.value);
                if (endDate.value) params.append('end_date', endDate.value);
                
                if (params.toString()) {
                    url += '?' + params.toString();
                }

                const response = await fetch(url);
                if (!response.ok) {
                    const errData = await response.json();
                    throw new Error(errData.error || 'Failed to export data');
                }

                // Trigger download
                const blob = await response.blob();
                const downloadUrl = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = downloadUrl;
                
                const contentDisposition = response.headers.get('Content-Disposition');
                let filename = 'invoice_export.csv';
                if (contentDisposition && contentDisposition.indexOf('filename=') !== -1) {
                    filename = contentDisposition.split('filename=')[1];
                }
                
                a.download = filename;
                document.body.appendChild(a);
                a.click();
                a.remove();
                window.URL.revokeObjectURL(downloadUrl);
                
                success.value = true;
                setTimeout(() => { success.value = false; }, 3000);
            } catch (err) {
                console.error('Export error:', err);
                error.value = err.message;
            } finally {
                exporting.value = false;
            }
        };

        // Initial fetch
        fetchRecentInvoices();

        return {
            startDate,
            endDate,
            exporting,
            error,
            success,
            recentInvoices,
            exportData,
            fetchRecentInvoices,
            formatDate
        };
    }
};
