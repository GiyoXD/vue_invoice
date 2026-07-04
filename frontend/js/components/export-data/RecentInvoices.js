import { ref } from 'vue';

export default {
    name: 'RecentInvoices',
    template: `
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
    `,
    setup() {
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

        // Initial fetch
        fetchRecentInvoices();

        return {
            recentInvoices,
            fetchRecentInvoices,
            formatDate
        };
    }
};
