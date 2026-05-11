import { ref, computed, onMounted, onUnmounted, nextTick, watch } from 'vue';

/**
 * LogViewer - Displays the current session log in a terminal-style viewer.
 *
 * Features:
 *   - Auto-refresh (polls /api/logs/current every 3s)
 *   - Log-level color coding (ERROR, WARNING, INFO, DEBUG)
 *   - Search/filter by keyword
 *   - Clear log button
 *   - Auto-scroll to bottom
 */
export default {
    template: `
        <div class="max-w-5xl mx-auto py-8 fade-in flex flex-col h-[calc(100vh-80px)]">
            <h1 class="text-4xl font-extrabold tracking-tight text-transparent bg-clip-text bg-gradient-to-r from-blue-400 to-emerald-400 drop-shadow-md flex-shrink-0 mb-6">Session Log</h1>

            <!-- Toolbar -->
            <div class="bg-slate-800/80 backdrop-blur-md border border-slate-700/50 shadow-2xl rounded-2xl p-6 mb-6 flex flex-col sm:flex-row justify-between items-center gap-4 flex-shrink-0">
                <div class="flex items-center gap-4 w-full sm:w-auto">
                    <input
                        type="text"
                        v-model="searchQuery"
                        placeholder="Filter logs..."
                        class="bg-slate-900 border border-slate-700 rounded-lg px-4 py-2 text-slate-100 focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 transition-all flex-1 sm:w-64"
                    />
                    <span class="text-slate-400 text-sm whitespace-nowrap">{{ filteredLines.length }} / {{ allLines.length }} lines</span>
                </div>
                <div class="flex items-center gap-4 flex-wrap w-full sm:w-auto justify-end">
                    <label class="flex items-center gap-2 text-slate-300 text-sm cursor-pointer">
                        <input type="checkbox" v-model="autoRefresh" class="accent-blue-500" />
                        <span>Auto-refresh</span>
                    </label>
                    <label class="flex items-center gap-2 text-slate-300 text-sm cursor-pointer">
                        <input type="checkbox" v-model="autoScroll" class="accent-blue-500" />
                        <span>Auto-scroll</span>
                    </label>
                    <button class="px-4 py-2 bg-slate-700 hover:bg-slate-600 text-slate-200 rounded-lg shadow-sm transition-colors text-sm" @click="fetchLog" :disabled="loading">
                        {{ loading ? '...' : '↻ Refresh' }}
                    </button>
                    <button class="px-4 py-2 bg-red-500/20 text-red-400 hover:bg-red-500/30 rounded-lg shadow-sm transition-colors text-sm border border-red-500/30" @click="clearLog">🗑 Clear</button>
                </div>
            </div>

            <!-- Log Container -->
            <div class="flex-1 bg-slate-900/90 border border-slate-700/50 shadow-2xl rounded-2xl p-4 overflow-y-auto font-mono text-sm leading-relaxed min-h-0 custom-scrollbar" ref="logContainerRef">
                <div v-if="allLines.length === 0 && !loading" class="text-slate-500 text-center py-8 italic">
                    No log data. Run an invoice generation to see output here.
                </div>
                <div
                    v-for="(line, index) in filteredLines"
                    :key="index"
                    class="py-0.5 break-words"
                    :class="[getLogLevelClass(line), 'text-slate-300']"
                >{{ line }}</div>
            </div>
        </div>
    `,
    setup() {
        const logContent = ref('');
        const searchQuery = ref('');
        const autoRefresh = ref(true);
        const autoScroll = ref(true);
        const loading = ref(false);
        const logContainerRef = ref(null);
        let refreshInterval = null;

        /**
         * Parse log content into individual lines.
         * @returns {string[]} Array of non-empty log lines.
         */
        const allLines = computed(() => {
            if (!logContent.value) return [];
            return logContent.value.split('\n').filter(line => line.trim() !== '');
        });

        /**
         * Filter lines by search query (case-insensitive).
         * @returns {string[]} Filtered array of log lines.
         */
        const filteredLines = computed(() => {
            if (!searchQuery.value) return allLines.value;
            const query = searchQuery.value.toLowerCase();
            return allLines.value.filter(line => line.toLowerCase().includes(query));
        });

        /**
         * Determine the CSS class for a log line based on its level.
         * @param {string} line - A single log line.
         * @returns {string} CSS class name.
         */
        const getLogLevelClass = (line) => {
            if (line.includes('| ERROR') || line.includes('| CRITICAL')) return '!text-red-400 font-bold';
            if (line.includes('| WARNING')) return '!text-yellow-400';
            if (line.includes('| DEBUG')) return '!text-slate-500';
            return '';
        };

        /**
         * Fetch current session log from the backend API.
         */
        const fetchLog = async () => {
            loading.value = true;
            try {
                const res = await fetch('/api/logs/current');
                if (res.ok) {
                    const data = await res.json();
                    logContent.value = data.content || '';

                    if (autoScroll.value) {
                        await nextTick();
                        scrollToBottom();
                    }
                }
            } catch (e) {
                console.error('Failed to fetch log:', e);
            } finally {
                loading.value = false;
            }
        };

        /**
         * Clear the session log via the backend API.
         */
        const clearLog = async () => {
            try {
                const res = await fetch('/api/logs/clear', { method: 'POST' });
                if (res.ok) {
                    logContent.value = '';
                }
            } catch (e) {
                console.error('Failed to clear log:', e);
            }
        };

        /**
         * Scroll the log container to the bottom.
         */
        const scrollToBottom = () => {
            const container = logContainerRef.value;
            if (container) {
                container.scrollTop = container.scrollHeight;
            }
        };

        /**
         * Start or stop the auto-refresh interval based on the toggle.
         */
        const startAutoRefresh = () => {
            stopAutoRefresh();
            if (autoRefresh.value) {
                refreshInterval = setInterval(fetchLog, 3000);
            }
        };

        const stopAutoRefresh = () => {
            if (refreshInterval) {
                clearInterval(refreshInterval);
                refreshInterval = null;
            }
        };

        // Watch autoRefresh toggle
        watch(autoRefresh, (newVal) => {
            if (newVal) {
                startAutoRefresh();
            } else {
                stopAutoRefresh();
            }
        });

        onMounted(() => {
            fetchLog();
            startAutoRefresh();
        });

        onUnmounted(() => {
            stopAutoRefresh();
        });

        return {
            logContent,
            searchQuery,
            autoRefresh,
            autoScroll,
            loading,
            logContainerRef,
            allLines,
            filteredLines,
            getLogLevelClass,
            fetchLog,
            clearLog
        };
    }
};
