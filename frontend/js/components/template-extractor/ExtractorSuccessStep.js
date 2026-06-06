import { useTemplateExtractorStore } from '../../stores/templateExtractorStore.js';

export default {
    name: 'ExtractorSuccessStep',
    template: `
        <div class="bg-slate-800/80 backdrop-blur-md border border-slate-700/50 shadow-2xl rounded-2xl p-8 mb-8 text-center delay-100 animate-in zoom-in" v-if="store.currentStep === 3">
            <div class="text-6xl mb-4">🎉</div>
            <h2>Template Created!</h2>
            <p class="text-secondary mb-4">
                The template <strong>{{ store.filePrefix }}</strong> has been configured successfully.
            </p>
            <div v-if="store.bundlePath" class="bg-emerald-500/10 p-4 rounded-xl mb-6 text-left">
                <p class="text-emerald-300 m-0 mb-2 text-sm">📁 Bundle created at:</p>
                <code class="text-emerald-500 text-xs break-all">{{ store.bundlePath }}</code>
                <div v-if="store.generatedPrefixes.length > 1" class="mt-2 pt-2 border-t border-emerald-500/20">
                    <p class="text-emerald-300 m-0 mb-1 text-xs">Contains:</p>
                    <div v-for="p in store.generatedPrefixes" :key="p" class="text-emerald-400 text-xs">
                        ✅ {{ p }}
                    </div>
                </div>
            </div>
            <p class="text-secondary mb-8">
                You can now go to the Generator and process invoices for this company.
            </p>
            <button class="w-full px-6 py-3 mt-4 bg-gradient-to-r from-blue-500 to-blue-600 hover:from-blue-400 hover:to-blue-500 text-white font-medium rounded-full shadow-lg shadow-blue-500/30 transition-all transform hover:-translate-y-0.5" @click="store.resetFlow">Process Another</button>
        </div>
    `,
    setup() {
        const store = useTemplateExtractorStore();
        return { store };
    }
};
