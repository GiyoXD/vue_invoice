import { ref } from 'vue';

export default {
    name: 'ExportForm',
    template: `
        <div class="bg-slate-800/80 backdrop-blur-md border border-slate-700/50 shadow-2xl rounded-2xl p-8 mb-8">
            <div class="flex flex-col md:flex-row gap-6 items-end">
                <div class="flex-1 w-full">
                    <label for="start-date" class="text-slate-300 font-medium">Start Date</label>
                    <input type="date" id="start-date" v-model="startDate" class="w-full mt-2 bg-slate-900 border border-slate-700 rounded-lg px-4 py-2 text-slate-100 focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 transition-all">
                </div>
                
                <div class="flex-1 w-full">
                    <label for="end-date" class="text-slate-300 font-medium">End Date</label>
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
    `,
    setup() {
        const startDate = ref('');
        const endDate = ref('');
        const exporting = ref(false);
        const error = ref(null);
        const success = ref(false);

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

        return {
            startDate,
            endDate,
            exporting,
            error,
            success,
            exportData
        };
    }
};
