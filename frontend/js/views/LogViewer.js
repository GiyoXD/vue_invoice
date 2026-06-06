import { ref, computed, onMounted, onUnmounted, watch } from 'vue';
import LogToolbar from '../components/log-viewer/LogToolbar.js';
import LogDisplay from '../components/log-viewer/LogDisplay.js';

/**
 * LogViewer - Orchestrates the sub-components to display the current session log.
 */
export default {
    name: 'LogViewer',
    components: {
        LogToolbar,
        LogDisplay
    },
    template: `
        <div class="max-w-5xl mx-auto py-8 fade-in flex flex-col h-[calc(100vh-80px)]">
            <h1 class="text-4xl font-extrabold tracking-tight text-transparent bg-clip-text bg-gradient-to-r from-blue-400 to-emerald-400 drop-shadow-md flex-shrink-0 mb-6">Session Log</h1>

            <log-toolbar
                v-model:searchQuery="searchQuery"
                v-model:autoRefresh="autoRefresh"
                v-model:autoScroll="autoScroll"
                :filteredCount="filteredLines.length"
                :totalCount="allLines.length"
                :loading="loading"
                @refresh="fetchLog"
                @clear="clearLog"
            ></log-toolbar>

            <log-display
                :lines="filteredLines"
                :loading="loading"
                :autoScroll="autoScroll"
            ></log-display>
        </div>
    `,
    setup() {
        const logContent = ref('');
        const searchQuery = ref('');
        const autoRefresh = ref(true);
        const autoScroll = ref(true);
        const loading = ref(false);
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
         * Fetch current session log from the backend API.
         */
        const fetchLog = async () => {
            loading.value = true;
            try {
                const res = await fetch('/api/logs/current');
                if (res.ok) {
                    const data = await res.json();
                    logContent.value = data.content || '';
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
            searchQuery,
            autoRefresh,
            autoScroll,
            loading,
            allLines,
            filteredLines,
            fetchLog,
            clearLog
        };
    }
};
