import { defineStore } from 'pinia';
import { ref, computed, watch } from 'vue';
import { recommendTruck } from '../utils/truck.js';

export const useInspectorStore = defineStore('inspector', () => {
    // State
    const uploadedMetadata = ref(null);
    const historyList = ref([]);
    const currentRun = ref(null);
    const existingInDb = ref(false);
    const isWideCargo = ref(false);

    // Computed
    const inspectorData = computed(() => {
        return uploadedMetadata.value;
    });

    const inspectorTotals = computed(() => {
        const data = inspectorData.value;
        const gt = data?.footer_data?.grand_total || {};

        let adjustmentSum = 0;
        if (data?.price_adjustment && Array.isArray(data.price_adjustment)) {
            data.price_adjustment.forEach(adj => {
                const val = Number(adj[1]);
                if (isNaN(val)) {
                    throw new Error(`Invalid price_adjustment value: "${adj[1]}" is not a number`);
                }
                adjustmentSum += val;
            });
        }

        return {
            pcs: gt.col_qty_pcs || 0,
            sqft: gt.col_qty_sf || 0,
            pallets: gt.col_pallet_count || 0,
            net: gt.col_net || 0,
            gross: gt.col_gross || 0,
            cbm: gt.col_cbm || 0,
            amount: Number(gt.col_amount || 0) + adjustmentSum
        };
    });

    const recommendedTruckInfo = computed(() => {
        const gross = inspectorTotals.value?.gross || 0;
        const cbm = inspectorTotals.value?.cbm || 0;
        const pallets = inspectorTotals.value?.pallets || 0;
        return recommendTruck(gross, cbm, pallets, isWideCargo.value);
    });

    const inspectorItems = computed(() => {
        const data = inspectorData.value;
        if (!data) return [];

        let items = [];

        if (data.price_adjustment && Array.isArray(data.price_adjustment)) {
            data.price_adjustment.forEach(adj => {
                items.push({
                    col_desc: adj[0],
                    col_amount: adj[1],
                    is_adjustment: true
                });
            });
        }

        const mainItems = (data.raw_data || []).flat();
        items = items.concat(mainItems);

        return items;
    });

    // Watchers
    watch(uploadedMetadata, (newData) => {
        if (newData) {
            const desc = String(newData.footer_data?.grand_total?.col_desc || '').toUpperCase();
            const file = String(newData.output_file || '').toUpperCase();
            if (desc.includes('LEATHER') || file.includes('JF') || file.includes('JLFTLT')) {
                isWideCargo.value = true;
            } else {
                isWideCargo.value = false;
            }
        }
    });

    // Actions
    const fetchHistory = async (autoLoadLatest = false) => {
        try {
            const res = await fetch('/api/history');
            if (res.ok) {
                historyList.value = await res.json();
                if (autoLoadLatest && historyList.value.length > 0) {
                    loadHistoryItem(historyList.value[0]);
                }
            }
        } catch (e) {
            console.error("Failed to fetch history", e);
        }
    };

    const loadHistoryItem = async (run) => {
        try {
            const res = await fetch(`/api/history/view?filename=${encodeURIComponent(run.filename)}&source=${run.type || 'run_log'}`);
            if (res.ok) {
                const data = await res.json();
                uploadedMetadata.value = data;
                currentRun.value = run;

                if (run.type === 'processed') {
                    try {
                        const checkRes = await fetch('/api/registry/check', {
                            method: 'POST',
                            headers: { 'Content-Type': 'application/json' },
                            body: JSON.stringify({ filename: run.filename })
                        });
                        if (checkRes.ok) {
                            const checkData = await checkRes.json();
                            existingInDb.value = checkData.exists;
                        } else {
                            existingInDb.value = false;
                        }
                    } catch (e) {
                        existingInDb.value = false;
                    }
                } else {
                    existingInDb.value = false;
                }
            } else {
                alert("Failed to load history item.");
            }
        } catch (e) {
            console.error("Error loading history item", e);
        }
    };

    const acceptCurrentRun = async () => {
        if (!currentRun.value || !currentRun.value.filename) return;

        try {
            const checkRes = await fetch('/api/registry/check', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ filename: currentRun.value.filename })
            });

            if (checkRes.ok) {
                const checkData = await checkRes.json();
                if (checkData.exists) {
                    if (!confirm(`⚠️ WARNING: The invoice "${currentRun.value.filename}" is ALREADY in the database.\n\nDo you want to REPLACE the existing data?`)) {
                        return;
                    }
                } else {
                    if (!confirm(`Accept and save "${currentRun.value.filename}" to database?`)) return;
                }
            } else {
                if (!confirm(`Accept and save "${currentRun.value.filename}" to database?`)) return;
            }
        } catch (e) {
            console.error("Failed to check registry", e);
            if (!confirm(`Accept and save "${currentRun.value.filename}" to database?`)) return;
        }

        try {
            const res = await fetch('/api/registry/accept', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ filename: currentRun.value.filename })
            });
            const result = await res.json();
            if (res.ok) {
                alert(result.message || "Saved successfully!");
                await fetchHistory();
                const acceptedItem = historyList.value.find(h => h.filename === currentRun.value.filename && h.type === 'accepted');
                if (acceptedItem) {
                    loadHistoryItem(acceptedItem);
                } else {
                    uploadedMetadata.value = null;
                    currentRun.value = null;
                }
            } else {
                alert("Error: " + result.error);
            }
        } catch (e) {
            console.error("Failed to accept run", e);
            alert("Failed to accept run");
        }
    };

    const rejectCurrentRun = async () => {
        if (!currentRun.value || !currentRun.value.filename) return;
        if (!confirm(`Reject and delete "${currentRun.value.filename}"? This cannot be undone.`)) return;

        try {
            const res = await fetch('/api/registry/reject', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ filename: currentRun.value.filename })
            });
            const result = await res.json();
            if (res.ok) {
                alert(result.message || "Deleted successfully!");
                uploadedMetadata.value = null;
                currentRun.value = null;
                await fetchHistory();
            } else {
                alert("Error: " + result.error);
            }
        } catch (e) {
            console.error("Failed to reject run", e);
            alert("Failed to reject run");
        }
    };

    const clearInspector = () => {
        uploadedMetadata.value = null;
    };

    const downloadExcel = (path) => {
        if (!path) return;
        window.location.href = `/api/download?path=${encodeURIComponent(path)}`;
    };

    return {
        uploadedMetadata,
        historyList,
        currentRun,
        existingInDb,
        isWideCargo,
        inspectorData,
        inspectorTotals,
        recommendedTruckInfo,
        inspectorItems,
        fetchHistory,
        loadHistoryItem,
        acceptCurrentRun,
        rejectCurrentRun,
        clearInspector,
        downloadExcel
    };
});
