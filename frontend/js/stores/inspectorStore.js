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
    const maxStackingLayers = ref(3);

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

    const detectedPalletDims = computed(() => {
        const items = inspectorItems.value;
        for (const item of items) {
            const cbmRaw = item.col_cbm_raw || item.col_cbm;
            if (typeof cbmRaw === 'string' && cbmRaw.trim()) {
                const parts = cbmRaw.split(/[xX*]/).map(p => parseFloat(p.trim())).filter(p => !isNaN(p));
                if (parts.length >= 2) {
                    let dim1 = parts[0];
                    let dim2 = parts[1];
                    if (dim1 > 10) dim1 = dim1 / 100;
                    if (dim2 > 10) dim2 = dim2 / 100;
                    if (dim1 > 10) dim1 = dim1 / 10;
                    if (dim2 > 10) dim2 = dim2 / 10;
                    
                    return {
                        length: Math.max(dim1, dim2),
                        width: Math.min(dim1, dim2)
                    };
                }
            }
        }
        return { length: 1.2, width: 1.0 };
    });

    const recommendedTruckInfo = computed(() => {
        const gross = inspectorTotals.value?.gross || 0;
        const cbm = inspectorTotals.value?.cbm || 0;
        const pallets = inspectorTotals.value?.pallets || 0;
        return recommendTruck(gross, cbm, pallets, isWideCargo.value, {
            maxStackingLayers: maxStackingLayers.value,
            palletLength: detectedPalletDims.value.length,
            palletWidth: detectedPalletDims.value.width
        });
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

        const mainItems = (data.multi_table || data.raw_data || []).flat();
        items = items.concat(mainItems);

        return items;
    });

    const detectWideCargoFromItems = (items) => {
        if (!Array.isArray(items)) return false;
        for (const item of items) {
            const cbmRaw = item.col_cbm_raw || item.col_cbm;
            if (typeof cbmRaw === 'string' && cbmRaw.trim()) {
                const parts = cbmRaw.split(/[xX*]/).map(p => parseFloat(p.trim())).filter(p => !isNaN(p));
                if (parts.length >= 2) {
                    let dim1 = parts[0];
                    let dim2 = parts[1];
                    
                    // Handle cm/mm conversions to meters if necessary
                    if (dim1 > 10) dim1 = dim1 / 100;
                    if (dim2 > 10) dim2 = dim2 / 100;
                    if (dim1 > 10) dim1 = dim1 / 10;
                    if (dim2 > 10) dim2 = dim2 / 10;
                    
                    // If any single horizontal dimension is >= 2.0 (e.g. 2.2m),
                    // or if both horizontal dimensions are > 1.1m (cannot fit side-by-side in 2.1m/2.2m truck):
                    if (dim1 >= 2.0 || dim2 >= 2.0 || (dim1 > 1.1 && dim2 > 1.1)) {
                        return true;
                    }
                }
            }
        }
        return false;
    };

    // Watchers
    watch(uploadedMetadata, (newData) => {
        if (newData) {
            const desc = String(newData.footer_data?.grand_total?.col_desc || '').toUpperCase();
            const file = String(newData.output_file || '').toUpperCase();
            let isWide = desc.includes('LEATHER') || file.includes('JF') || file.includes('JLFTLT');
            
            if (!isWide) {
                const items = (newData.multi_table || newData.raw_data || []).flat();
                if (detectWideCargoFromItems(items)) {
                    isWide = true;
                }
            }
            
            isWideCargo.value = isWide;
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
        maxStackingLayers,
        detectedPalletDims,
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
