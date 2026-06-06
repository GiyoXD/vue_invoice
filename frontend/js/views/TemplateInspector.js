import { onMounted } from 'vue';
import { useTemplateInspectorStore } from '../stores/templateInspectorStore.js';
import TemplateSidebar from '../components/template-inspector/TemplateSidebar.js';
import TemplateInfo from '../components/template-inspector/TemplateInfo.js';
import TemplateGrid from '../components/template-inspector/TemplateGrid.js';

export default {
    name: 'TemplateInspectorView',
    components: {
        TemplateSidebar,
        TemplateInfo,
        TemplateGrid
    },
    template: `
        <div class="h-full flex flex-col fade-in">
            <h1 class="text-4xl font-extrabold tracking-tight text-transparent bg-clip-text bg-gradient-to-r from-blue-400 to-emerald-400 drop-shadow-md flex-shrink-0">Template Inspector</h1>
            
            <div class="flex gap-6 mt-8 flex-grow min-h-0">
                <!-- Sidebar: Template List -->
                <template-sidebar></template-sidebar>

                <!-- Main: Details -->
                <div class="flex-1 bg-slate-800/80 backdrop-blur-md border border-slate-700/50 shadow-2xl rounded-2xl p-6 flex flex-col min-w-0">
                    <div v-if="!store.currentTemplate" class="text-center p-8 text-slate-400 my-auto">
                        <p class="text-lg">Select a template from the list to inspect.</p>
                    </div>

                    <div v-else class="template-viewer overflow-y-auto flex-1 pr-2 custom-scrollbar">
                        <!-- Client & Table Information -->
                        <template-info class="mb-6"></template-info>

                        <!-- Layout Grid Preview -->
                        <template-grid></template-grid>
                    </div>
                </div>
            </div>
        </div>
    `,
    setup() {
        const store = useTemplateInspectorStore();

        onMounted(() => {
            store.fetchTemplates();
        });

        return {
            store
        };
    }
};
