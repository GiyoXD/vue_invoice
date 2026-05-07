import { createApp, ref } from 'vue';
import GeneratorView from './views/Generator.js?v=5';
import InspectorView from './views/Inspector.js?v=5';
import TemplateExtractorView from './views/TemplateExtractor.js?v=5';
import TemplateInspectorView from './views/TemplateInspector.js?v=5';
import LogViewerView from './views/LogViewer.js?v=5';
import ExportDataView from './views/ExportData.js?v=5';

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
        <div class="container fade-in">
            <!-- Navigation -->
            <div class="nav-bar">
                <button class="nav-btn" :class="{ active: currentView === 'home' }" @click="currentView = 'home'">⚡ Generator</button>
                <button class="nav-btn" :class="{ active: currentView === 'inspector' }" @click="currentView = 'inspector'">🔍 Inspector</button>
                <button class="nav-btn" :class="{ active: currentView === 'export' }" @click="currentView = 'export'">📦 Export</button>
                <button class="nav-btn" :class="{ active: currentView === 'template_inspector' }" @click="currentView = 'template_inspector'">📐 Templates</button>
                <button class="nav-btn" :class="{ active: currentView === 'extractor' }" @click="currentView = 'extractor'">✨ New Template</button>
                <button class="nav-btn" :class="{ active: currentView === 'logs' }" @click="currentView = 'logs'">📋 Logs</button>
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

createApp(App).mount('#app');
