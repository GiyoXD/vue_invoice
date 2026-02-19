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
        <div class="log-viewer-view fade-in">
            <h1>Session Log</h1>

            <!-- Toolbar -->
            <div class="log-toolbar card">
                <div class="log-toolbar-left">
                    <input
                        type="text"
                        v-model="searchQuery"
                        placeholder="Filter logs..."
                        class="log-search-input"
                    />
                    <span class="log-line-count">{{ filteredLines.length }} / {{ allLines.length }} lines</span>
                </div>
                <div class="log-toolbar-right">
                    <label class="log-toggle">
                        <input type="checkbox" v-model="autoRefresh" />
                        <span>Auto-refresh</span>
                    </label>
                    <label class="log-toggle">
                        <input type="checkbox" v-model="autoScroll" />
                        <span>Auto-scroll</span>
                    </label>
                    <button class="btn-small log-btn-refresh" @click="fetchLog" :disabled="loading">
                        {{ loading ? '...' : '↻ Refresh' }}
                    </button>
                    <button class="btn-small log-btn-clear" @click="clearLog">🗑 Clear</button>
                </div>
            </div>

            <!-- Log Container -->
            <div class="log-container card" ref="logContainerRef">
                <div v-if="allLines.length === 0 && !loading" class="log-empty">
                    No log data. Run an invoice generation to see output here.
                </div>
                <div
                    v-for="(line, index) in filteredLines"
                    :key="index"
                    class="log-line"
                    :class="getLogLevelClass(line)"
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
            if (line.includes('| ERROR') || line.includes('| CRITICAL')) return 'log-error';
            if (line.includes('| WARNING')) return 'log-warning';
            if (line.includes('| DEBUG')) return 'log-debug';
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
