import { createApp, ref } from 'vue';
import { createPinia } from 'pinia';
import GeneratorView from './views/Generator.js?v=7';
import InspectorView from './views/Inspector.js?v=6';
import TemplateExtractorView from './views/TemplateExtractor.js?v=9';
import TemplateInspectorView from './views/TemplateInspector.js?v=10';
import LogViewerView from './views/LogViewer.js?v=6';
import ExportDataView from './views/ExportData.js?v=6';

const App = {
    components: {
        GeneratorView,
        InspectorView,
        TemplateExtractorView,
        TemplateInspectorView,
        LogViewerView,
        ExportDataView
    },
    template: `
        <div class="max-w-[1600px] mx-auto px-6 py-6 fade-in h-screen flex flex-col">
            <!-- Navigation -->
            <div class="flex gap-4 p-4 bg-slate-800/80 backdrop-blur-md border border-slate-700/50 shadow-2xl rounded-2xl mb-8 overflow-x-auto custom-scrollbar flex-shrink-0 justify-center">
                <button class="px-6 py-3 font-medium rounded-xl transition-all whitespace-nowrap" :class="currentView === 'home' ? 'bg-blue-600 text-white shadow-lg shadow-blue-500/30' : 'bg-slate-700/50 hover:bg-slate-600/80 text-slate-300'" @click="currentView = 'home'">⚡ Generator</button>
                <button class="px-6 py-3 font-medium rounded-xl transition-all whitespace-nowrap" :class="currentView === 'inspector' ? 'bg-blue-600 text-white shadow-lg shadow-blue-500/30' : 'bg-slate-700/50 hover:bg-slate-600/80 text-slate-300'" @click="currentView = 'inspector'">🔍 Inspector</button>
                <button class="px-6 py-3 font-medium rounded-xl transition-all whitespace-nowrap" :class="currentView === 'export' ? 'bg-blue-600 text-white shadow-lg shadow-blue-500/30' : 'bg-slate-700/50 hover:bg-slate-600/80 text-slate-300'" @click="currentView = 'export'">📦 Export</button>
                <button class="px-6 py-3 font-medium rounded-xl transition-all whitespace-nowrap" :class="currentView === 'template_inspector' ? 'bg-blue-600 text-white shadow-lg shadow-blue-500/30' : 'bg-slate-700/50 hover:bg-slate-600/80 text-slate-300'" @click="currentView = 'template_inspector'">📐 Templates</button>
                <button class="px-6 py-3 font-medium rounded-xl transition-all whitespace-nowrap" :class="currentView === 'extractor' ? 'bg-blue-600 text-white shadow-lg shadow-blue-500/30' : 'bg-slate-700/50 hover:bg-slate-600/80 text-slate-300'" @click="currentView = 'extractor'">✨ New Template</button>
                <button class="px-6 py-3 font-medium rounded-xl transition-all whitespace-nowrap" :class="currentView === 'logs' ? 'bg-blue-600 text-white shadow-lg shadow-blue-500/30' : 'bg-slate-700/50 hover:bg-slate-600/80 text-slate-300'" @click="currentView = 'logs'">📋 Logs</button>
            </div>

            <!-- HOME VIEW: Generator -->
            <div v-show="currentView === 'home'">
                <generator-view @switch-view="switchView"></generator-view>
            </div>

            <!-- INSPECTOR VIEW -->
            <div v-show="currentView === 'inspector'">
                <inspector-view ref="inspectorRef"></inspector-view>
            </div>

            <!-- EXPORT VIEW -->
            <div v-show="currentView === 'export'">
                <export-data-view></export-data-view>
            </div>
            
            <!-- EXTRACTOR VIEW -->
            <div v-show="currentView === 'extractor'">
                <template-extractor-view></template-extractor-view>
            </div>

            <!-- TEMPLATE INSPECTOR VIEW -->
            <div v-show="currentView === 'template_inspector'">
                <template-inspector-view></template-inspector-view>
            </div>

            <!-- LOG VIEWER -->
            <div v-show="currentView === 'logs'">
                <log-viewer-view></log-viewer-view>
            </div>

        </div>
    `,
    setup() {
        const currentView = ref('home');
        const inspectorRef = ref(null);

        const switchView = (viewName) => {
            currentView.value = viewName;
            // Optionally trigger refresh if switching to inspector
            if (viewName === 'inspector' && inspectorRef.value) {
                // inspectorRef.value.fetchHistory(); // If needed, but onMounted handles it
            }
        };

        return {
            currentView,
            switchView,
            inspectorRef
        };
    }
};

const app = createApp(App);
app.use(createPinia());
app.mount('#app');
