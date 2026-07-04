import ExportForm from '../components/export-data/ExportForm.js';
import RecentInvoices from '../components/export-data/RecentInvoices.js';

export default {
    name: 'ExportDataView',
    components: {
        ExportForm,
        RecentInvoices
    },
    template: `
        <div class="max-w-5xl mx-auto py-8 fade-in">
            <div class="mb-8">
                <h1 class="text-4xl font-extrabold tracking-tight text-transparent bg-clip-text bg-gradient-to-r from-blue-400 to-emerald-400 drop-shadow-md mb-2">Data Export Registry</h1>
                <p class="text-slate-400">Select a time interval to export stored invoice data to CSV.</p>
            </div>

            <export-form></export-form>

            <recent-invoices></recent-invoices>
        </div>
    `,
    setup() {
        return {};
    }
};
