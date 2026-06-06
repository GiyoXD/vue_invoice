import { ref, watch, nextTick } from 'vue';

export default {
    name: 'LogDisplay',
    props: {
        lines: {
            type: Array,
            required: true
        },
        loading: {
            type: Boolean,
            default: false
        },
        autoScroll: {
            type: Boolean,
            default: true
        }
    },
    template: `
        <div class="flex-1 bg-slate-900/90 border border-slate-700/50 shadow-2xl rounded-2xl p-4 overflow-y-auto font-mono text-sm leading-relaxed min-h-0 custom-scrollbar" ref="logContainerRef">
            <div v-if="lines.length === 0 && !loading" class="text-slate-500 text-center py-8 italic">
                No log data. Run an invoice generation to see output here.
            </div>
            <div
                v-for="(line, index) in lines"
                :key="index"
                class="py-0.5 break-words"
                :class="[getLogLevelClass(line), 'text-slate-300']"
            >{{ line }}</div>
        </div>
    `,
    setup(props) {
        const logContainerRef = ref(null);

        const getLogLevelClass = (line) => {
            if (line.includes('| ERROR') || line.includes('| CRITICAL')) return '!text-red-400 font-bold';
            if (line.includes('| WARNING')) return '!text-yellow-400';
            if (line.includes('| DEBUG')) return '!text-slate-500';
            return '';
        };

        const scrollToBottom = () => {
            const container = logContainerRef.value;
            if (container) {
                container.scrollTop = container.scrollHeight;
            }
        };

        // Scroll to bottom when lines change, if autoScroll is enabled
        watch(() => props.lines, async () => {
            if (props.autoScroll) {
                await nextTick();
                scrollToBottom();
            }
        }, { deep: true });

        return {
            logContainerRef,
            getLogLevelClass,
            scrollToBottom
        };
    }
};
