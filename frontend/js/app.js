import { createApp, ref } from 'vue';
import GeneratorView from './views/Generator.js';
import InspectorView from './views/Inspector.js';
import TemplateExtractorView from './views/TemplateExtractor.js?v=2';
import TemplateInspectorView from './views/TemplateInspector.js?v=revert_spill';

const App = {
    components: {
        GeneratorView,
        InspectorView,
        TemplateExtractorView,
        TemplateInspectorView
    },
    template: `
        <div class="container fade-in">
            <!-- Navigation -->
            <div class="nav-bar">
                <button class="nav-btn" :class="{ active: currentView === 'home' }" @click="currentView = 'home'">Generator</button>
                <button class="nav-btn" :class="{ active: currentView === 'inspector' }" @click="currentView = 'inspector'">Data Inspector</button>
                <button class="nav-btn" :class="{ active: currentView === 'template_inspector' }" @click="currentView = 'template_inspector'">Inspect Template</button>
                <button class="nav-btn" :class="{ active: currentView === 'extractor' }" @click="currentView = 'extractor'">New Template</button>
            </div>

            <!-- HOME VIEW: Generator -->
            <div v-show="currentView === 'home'">
                <generator-view @switch-view="switchView"></generator-view>
            </div>

            <!-- INSPECTOR VIEW -->
            <div v-show="currentView === 'inspector'">
                <inspector-view ref="inspectorRef"></inspector-view>
            </div>
            
            <!-- EXTRACTOR VIEW -->
            <div v-show="currentView === 'extractor'">
                <template-extractor-view></template-extractor-view>
            </div>

            <!-- TEMPLATE INSPECTOR VIEW -->
            <div v-show="currentView === 'template_inspector'">
                <template-inspector-view></template-inspector-view>
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
