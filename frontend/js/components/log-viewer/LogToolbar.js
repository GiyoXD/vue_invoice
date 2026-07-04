export default {
    name: 'LogToolbar',
    props: {
        searchQuery: {
            type: String,
            default: ''
        },
        filteredCount: {
            type: Number,
            required: true
        },
        totalCount: {
            type: Number,
            required: true
        },
        autoRefresh: {
            type: Boolean,
            required: true
        },
        autoScroll: {
            type: Boolean,
            required: true
        },
        loading: {
            type: Boolean,
            default: false
        }
    },
    emits: [
        'update:searchQuery',
        'update:autoRefresh',
        'update:autoScroll',
        'refresh',
        'clear'
    ],
    template: `
        <div class="bg-slate-800/80 backdrop-blur-md border border-slate-700/50 shadow-2xl rounded-2xl p-6 mb-6 flex flex-col sm:flex-row justify-between items-center gap-4 flex-shrink-0">
            <div class="flex items-center gap-4 w-full sm:w-auto">
                <input
                    type="text"
                    :value="searchQuery"
                    @input="$emit('update:searchQuery', $event.target.value)"
                    placeholder="Filter logs..."
                    class="bg-slate-900 border border-slate-700 rounded-lg px-4 py-2 text-slate-100 focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 transition-all flex-1 sm:w-64"
                />
                <span class="text-slate-400 text-sm whitespace-nowrap">{{ filteredCount }} / {{ totalCount }} lines</span>
            </div>
            <div class="flex items-center gap-4 flex-wrap w-full sm:w-auto justify-end">
                <label class="flex items-center gap-2 text-slate-300 text-sm cursor-pointer">
                    <input
                        type="checkbox"
                        :checked="autoRefresh"
                        @change="$emit('update:autoRefresh', $event.target.checked)"
                        class="accent-blue-500"
                    />
                    <span>Auto-refresh</span>
                </label>
                <label class="flex items-center gap-2 text-slate-300 text-sm cursor-pointer">
                    <input
                        type="checkbox"
                        :checked="autoScroll"
                        @change="$emit('update:autoScroll', $event.target.checked)"
                        class="accent-blue-500"
                    />
                    <span>Auto-scroll</span>
                </label>
                <button
                    class="px-4 py-2 bg-slate-700 hover:bg-slate-600 text-slate-200 rounded-lg shadow-sm transition-colors text-sm"
                    @click="$emit('refresh')"
                    :disabled="loading"
                >
                    {{ loading ? '...' : '↻ Refresh' }}
                </button>
                <button
                    class="px-4 py-2 bg-red-500/20 text-red-400 hover:bg-red-500/30 rounded-lg shadow-sm transition-colors text-sm border border-red-500/30"
                    @click="$emit('clear')"
                >
                    🗑 Clear
                </button>
            </div>
        </div>
    `
};
