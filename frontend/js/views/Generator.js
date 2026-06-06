import FileUploader from '../components/generator/FileUploader.js';
import InvoiceDetailsForm from '../components/generator/InvoiceDetailsForm.js';
import ValidationStats from '../components/generator/ValidationStats.js';

export default {
    name: 'GeneratorView',
    components: {
        FileUploader,
        InvoiceDetailsForm,
        ValidationStats
    },
    emits: ['switch-view'],
    template: `
        <div class="generator-view fade-in max-w-5xl mx-auto py-8">
            <h1 class="text-4xl font-extrabold tracking-tight text-transparent bg-clip-text bg-gradient-to-r from-blue-400 to-emerald-400 mb-8 drop-shadow-md">Invoice Generator</h1>
            
            <file-uploader></file-uploader>
            
            <invoice-details-form @switch-view="(val) => $emit('switch-view', val)"></invoice-details-form>
            
            <validation-stats></validation-stats>
        </div>
    `,
    setup() {
        return {};
    }
};
