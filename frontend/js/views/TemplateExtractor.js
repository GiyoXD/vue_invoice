import { useTemplateExtractorStore } from '../stores/templateExtractorStore.js';
import ExtractorUploadStep from '../components/template-extractor/ExtractorUploadStep.js';
import ExtractorMappingStep from '../components/template-extractor/ExtractorMappingStep.js';
import ExtractorSuccessStep from '../components/template-extractor/ExtractorSuccessStep.js';
import GlobalMappingsManager from '../components/template-extractor/GlobalMappingsManager.js';

export default {
    name: 'TemplateExtractor',
    components: {
        ExtractorUploadStep,
        ExtractorMappingStep,
        ExtractorSuccessStep,
        GlobalMappingsManager
    },
    template: `
        <div class="template-extractor-view fade-in">
            <h1>New Template Extractor</h1>
            
            <!-- Step 1: Upload Source File -->
            <extractor-upload-step />

            <!-- Step 2: Map Headers/Footers -->
            <extractor-mapping-step />

            <!-- Step 3: Success Status -->
            <extractor-success-step />

            <!-- Global Mappings Management -->
            <global-mappings-manager />
        </div>
    `,
    setup() {
        const store = useTemplateExtractorStore();

        // Load initial data on view creation
        store.fetchOptions();
        store.fetchMappings();
        store.fetchFooterMappings();

        return { store };
    }
};
