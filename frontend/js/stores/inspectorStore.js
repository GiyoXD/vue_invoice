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
    const maxStackingLayers = ref(2);
    const loadingOrientation = ref('width');

    // Computed
    const inspectorData = computed(() => {
        return uploadedMetadata.value;
    });

    const normalizeDimensionToMeters = (val) => {
        if (val === null || val === undefined || isNaN(val)) return 0;
        const num = parseFloat(val);
        if (num <= 0 || isNaN(num)) return 0;
        return num; // input is always in meters
    };

    const extractDimsFromItem = (item) => {
        if (!item) return null;
        const candidates = [item.col_cbm_raw, item.cbm_raw, item.col_cbm, item.cbm];
        for (const candidate of candidates) {
            if (typeof candidate === 'string' && (candidate.includes('*') || /[xX]/.test(candidate))) {
                const parts = candidate.split(/[xX*]/).map(p => parseFloat(p.trim())).filter(p => !isNaN(p));
                if (parts.length >= 2) {
                    const dim1 = normalizeDimensionToMeters(parts[0]);
                    const dim2 = normalizeDimensionToMeters(parts[1]);
                    let dim3 = parts.length >= 3 ? normalizeDimensionToMeters(parts[2]) : null;
                    if (dim1 > 0 && dim2 > 0) {
                        return {
                            length: Math.max(dim1, dim2),
                            width: Math.min(dim1, dim2),
                            height: dim3 || null
                        };
                    }
                }
            }
        }
        return null;
    };

    const inspectorItems = computed(() => {
        const data = inspectorData.value;
        if (!data) return [];

        let items = [];

        if (data.price_adjustment && Array.isArray(data.price_adjustment)) {
            for (const adj of data.price_adjustment) {
                if (!adj) continue;
                const desc = Array.isArray(adj) ? adj[0] : (adj.col_desc ?? adj.description ?? 'Adjustment');
                const rawAmt = Array.isArray(adj) ? adj[1] : (adj.amount ?? adj.col_amount ?? 0);
                const numAmt = Number(rawAmt);
                if (isNaN(numAmt)) continue;
                items.push({
                    col_desc: desc,
                    col_amount: numAmt,
                    is_adjustment: true
                });
            }
        }

        const mainItems = (data.multi_table || []).flat();
        items = items.concat(mainItems);

        return items;
    });

    const inspectorTotals = computed(() => {
        const data = inspectorData.value;
        const gt = data?.footer_data?.grand_total || {};

        let adjustmentSum = 0;
        if (data?.price_adjustment && Array.isArray(data.price_adjustment)) {
            for (const adj of data.price_adjustment) {
                if (!adj) continue;
                const rawAmt = Array.isArray(adj) ? adj[1] : (adj.amount ?? adj.col_amount ?? 0);
                const numAmt = Number(rawAmt);
                if (isNaN(numAmt)) continue;
                adjustmentSum += numAmt;
            }
        }

        const items = inspectorItems.value;
        let pallets = Number(gt.col_pallet_count) || 0;
        if (!pallets && items.length > 0) {
            pallets = items.filter(it => !it.is_adjustment && extractDimsFromItem(it) !== null).length || items.filter(it => !it.is_adjustment).length;
        }

        const rawItems = (data?.multi_table || []).flat();
        const rawSum = rawItems.reduce((acc, it) => acc + (Number(it.col_amount) || 0), 0);
        const gtAmount = Number(gt.col_amount || 0);

        const isAlreadyAdjusted = Boolean(
            (adjustmentSum !== 0 && rawSum > 0 && Math.abs(gtAmount - (rawSum + adjustmentSum)) < 0.01) ||
            (adjustmentSum === 0)
        );

        return {
            pcs: gt.col_qty_pcs || 0,
            sqft: gt.col_qty_sf || 0,
            pallets: pallets,
            net: gt.col_net || 0,
            gross: gt.col_gross || 0,
            cbm: gt.col_cbm || 0,
            amount: isAlreadyAdjusted ? gtAmount : gtAmount + adjustmentSum
        };
    });

    const detectedPalletDims = computed(() => {
        const items = inspectorItems.value;
        let baseDims = null;
        let maxHeight = 0;
        for (const item of items) {
            const dims = extractDimsFromItem(item);
            if (dims) {
                if (!baseDims) baseDims = dims;
                if (dims.height && dims.height > maxHeight) {
                    maxHeight = dims.height;
                }
            }
        }
        if (!baseDims) return null;
        const { length, width } = baseDims;
        return { length, width, height: maxHeight || null };
    });

    const recommendedTruckInfo = computed(() => {
        const gross = inspectorTotals.value?.gross || 0;
        const cbm = inspectorTotals.value?.cbm || 0;
        const pallets = inspectorTotals.value?.pallets || 0;
        const dims = detectedPalletDims.value;

        if (!dims) {
            return {
                name: 'Missing Dimensions',
                maxWeight: 0,
                maxCbm: 0,
                maxPallets: 0,
                displayName: 'Missing Dims ⚠️',
                description: 'Cannot recommend truck: Pallet dimensions (L×W) not found in data',
                color: 'text-amber-400 bg-amber-500/10 border-amber-500/30'
            };
        }

        const items = inspectorItems.value;
        let maxPalletWeight = 0;
        let curPalletGross = 0;
        for (const item of items) {
            const count = Number(item.col_pallet_count);
            const grossVal = Number(item.col_gross || item.gross || 0);
            if (count > 0 && curPalletGross > 0) {
                if (curPalletGross > maxPalletWeight) maxPalletWeight = curPalletGross;
                curPalletGross = grossVal;
            } else {
                curPalletGross += grossVal;
            }
        }
        if (curPalletGross > maxPalletWeight) maxPalletWeight = curPalletGross;

        return recommendTruck(gross, cbm, pallets, isWideCargo.value, {
            maxStackingLayers: maxStackingLayers.value,
            loadingOrientation: loadingOrientation.value,
            palletLength: dims.length,
            palletWidth: dims.width,
            palletHeight: dims.height,
            maxPalletWeight
        });
    });

    const detectWideCargoFromItems = (items) => {
        if (!Array.isArray(items)) return false;
        for (const item of items) {
            const dims = extractDimsFromItem(item);
            if (dims) {
                if (dims.length >= 2.0 || dims.width >= 2.0 || (dims.length > 1.1 && dims.width > 1.1)) {
                    return true;
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
                const items = (newData.multi_table || []).flat();
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
        loadingOrientation,
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
