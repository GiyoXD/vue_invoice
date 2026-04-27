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

                <div v-if="error" class="error-message" style="color: #ef4444; margin-top: 1rem; text-align: center;">
                    {{ error }}
                </div>
                
                <div v-if="success" class="success-message">
                    Export completed successfully.
                </div>
            </div>

            <!-- Data Preview Section (Peek) -->
            <div class="card preview-card" style="margin-top: 2rem; background: rgba(30, 41, 59, 0.4); border: 1px solid rgba(255, 255, 255, 0.05);">
                <h3>Recent Invoices (Peek)</h3>
                <div class="table-container" style="margin-top: 1rem; overflow-x: auto;">
                    <table style="width: 100%; border-collapse: collapse; font-size: 0.9rem;">
                        <thead style="background: rgba(255, 255, 255, 0.05);">
                            <tr>
                                <th style="padding: 0.75rem; text-align: left;">Filename</th>
                                <th style="padding: 0.75rem; text-align: left;">Accepted At</th>
                                <th style="padding: 0.75rem; text-align: center;">Items</th>
                                <th style="padding: 0.75rem; text-align: right;">Total SQFT</th>
                                <th style="padding: 0.75rem; text-align: right;">Total Net</th>
                                <th style="padding: 0.75rem; text-align: right;">Total Pallets</th>
                                <th style="padding: 0.75rem; text-align: right;">Total Amount</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr v-for="item in recentInvoices" :key="item.id" style="border-bottom: 1px solid rgba(255, 255, 255, 0.05);">
                                <td style="padding: 0.75rem;">{{ item.filename }}</td>
                                <td style="padding: 0.75rem;">{{ formatDate(item.timestamp) }}</td>
                                <td style="padding: 0.75rem; text-align: center;">{{ item.item_count }}</td>
                                <td style="padding: 0.75rem; text-align: right; font-family: monospace;">{{ item.total_sqft?.toLocaleString(undefined, {minimumFractionDigits: 2}) || '0.00' }}</td>
                                <td style="padding: 0.75rem; text-align: right; font-family: monospace;">{{ item.total_net?.toLocaleString(undefined, {minimumFractionDigits: 2}) || '0.00' }} kg</td>
                                <td style="padding: 0.75rem; text-align: right; font-family: monospace;">{{ item.total_pallets?.toLocaleString(undefined, {minimumFractionDigits: 2}) || '0.00' }}</td>
                                <td style="padding: 0.75rem; text-align: right; font-family: monospace;">$ {{ item.total_amount?.toLocaleString(undefined, {minimumFractionDigits: 2}) || '0.00' }}</td>
                            </tr>
                            <tr v-if="recentInvoices.length === 0">
                                <td colspan="7" style="padding: 2rem; text-align: center; color: rgba(255, 255, 255, 0.4);">No invoices found in registry.</td>
                            </tr>
                        </tbody>
                    </table>
                </div>
                <div style="margin-top: 1rem; text-align: right;">
                    <button @click="fetchRecentInvoices" class="action-btn" style="padding: 0.4rem 0.8rem; font-size: 0.8rem; background: rgba(59, 130, 246, 0.2); color: #60a5fa; border: 1px solid rgba(59, 130, 246, 0.3);">
                        Refresh List
                    </button>
                </div>
            </div>

            <div class="danger-zone" style="margin-top: 4rem; padding: 2rem; border: 1px solid rgba(239, 68, 68, 0.2); border-radius: 12px; background: rgba(239, 68, 68, 0.05);">
                <h3 style="color: #f87171; margin-top: 0;">Danger Zone</h3>
                <p style="color: rgba(248, 113, 113, 0.8); font-size: 0.9rem;">Resetting the database will permanently delete all stored invoice processing history and the master list table.</p>
                <div style="margin-top: 1.5rem;">
                    <button @click="confirmReset" class="action-btn danger-btn" style="background: #dc2626; color: white;">
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
