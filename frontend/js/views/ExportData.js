import { ref } from 'vue';

export default {
    name: 'ExportDataView',
    template: `
        <div class="view-container">
            <div class="header-section">
                <h1>Data Export Registry</h1>
                <p class="subtitle">Select a time interval to export stored invoice data to CSV.</p>
            </div>

            <div class="card export-card">
                <div class="export-controls">
                    <div class="control-group">
                        <label for="start-date">Start Date</label>
                        <input type="date" id="start-date" v-model="startDate" class="date-input">
                    </div>
                    
                    <div class="control-group">
                        <label for="end-date">End Date</label>
                        <input type="date" id="end-date" v-model="endDate" class="date-input">
                    </div>

                    <div class="action-group">
                        <button @click="exportData" :disabled="exporting" class="action-btn primary-btn">
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
            <div class="card preview-card mt-8 bg-slate-800-40 border border-white-5">
                <h3>Recent Invoices (Peek)</h3>
                <div class="table-container mt-4 overflow-x-auto">
                    <table class="w-full border-collapse text-sm">
                        <thead class="bg-white-5">
                            <tr>
                                <th class="p-3 text-left">Filename</th>
                                <th class="p-3 text-left">Accepted At</th>
                                <th class="p-3 text-center">Items</th>
                                <th class="p-3 text-right">Total SQFT</th>
                                <th class="p-3 text-right">Total Net</th>
                                <th class="p-3 text-right">Total Pallets</th>
                                <th class="p-3 text-right">Total Amount</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr v-for="item in recentInvoices" :key="item.id" class="border-b border-white-5">
                                <td class="p-3">{{ item.filename }}</td>
                                <td class="p-3">{{ formatDate(item.timestamp) }}</td>
                                <td class="p-3 text-center">{{ item.item_count }}</td>
                                <td class="p-3 text-right font-mono">{{ item.total_sqft?.toLocaleString(undefined, {minimumFractionDigits: 2}) || '0.00' }}</td>
                                <td class="p-3 text-right font-mono">{{ item.total_net?.toLocaleString(undefined, {minimumFractionDigits: 2}) || '0.00' }} kg</td>
                                <td class="p-3 text-right font-mono">{{ item.total_pallets?.toLocaleString(undefined, {minimumFractionDigits: 2}) || '0.00' }}</td>
                                <td class="p-3 text-right font-mono">$ {{ item.total_amount?.toLocaleString(undefined, {minimumFractionDigits: 2}) || '0.00' }}</td>
                            </tr>
                            <tr v-if="recentInvoices.length === 0">
                                <td colspan="7" class="p-8 text-center text-white-40">No invoices found in registry.</td>
                            </tr>
                        </tbody>
                    </table>
                </div>
                <div class="mt-4 text-right">
                    <button @click="fetchRecentInvoices" class="action-btn text-xs py-2 px-3 bg-blue-100 text-blue-300 border border-blue-200">
                        Refresh List
                    </button>
                </div>
            </div>

            <div class="danger-zone mt-16 p-8 border border-red-500-20 rounded-xl bg-red-500-5">
                <h3 class="text-red-400 mt-0">Danger Zone</h3>
                <p class="text-red-400-80 text-sm">Resetting the database will permanently delete all stored invoice processing history and the master list table.</p>
                <div class="mt-6">
                    <button @click="confirmReset" class="action-btn danger-btn bg-red-600 text-white">
                        Reset Database Registry
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

        const confirmReset = async () => {
            if (confirm('CRITICAL ACTION: This will permanently delete ALL data in the registry and the master list. It will also delete the processed JSON files on disk. PROCEED?')) {
                try {
                    const response = await fetch('/api/registry/reset', { method: 'POST' });
                    if (response.ok) {
                        alert('Database and files have been successfully reset.');
                        fetchRecentInvoices();
                    } else {
                        const errData = await response.json();
                        alert('Reset failed: ' + errData.error);
                    }
                } catch (err) {
                    alert('Reset failed: ' + err.message);
                }
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
            formatDate,
            confirmReset
        };
    }
};
