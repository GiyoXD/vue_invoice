import { defineStore } from 'pinia';
import { ref, computed, watch } from 'vue';
import { recommendTruck } from '../utils/truck.js';

export const useGeneratorStore = defineStore('generator', () => {
    // --- State ---
    const selectedFile = ref(null);
    let rawFile = null; // Native file object reference
    const isUploading = ref(false);
    const uploadStatus = ref(null);
    const uploadError = ref(null);
    const showUploadTraceback = ref(false);
    const ignoreTareError = ref(false);
    const ignoreCbmError = ref(false);

    const processingComplete = ref(false);
    const identifier = ref('');
    const jsonPath = ref('');

    const invoiceNo = ref('');
    const invoiceDate = ref(new Date().toISOString().split('T')[0]);
    const invoiceRef = ref('');
    const refSourceStatus = ref(null);

    // Options
    const includeStandard = ref(true);
    const includeCustom = ref(false);
    const includeDAF = ref(false);
    const selectedVariants = ref([]);
    const splitSheets = ref(false);
    const selectedTargets = ref([]);

    const priceAdjustments = ref([]);
    const adjustmentError = ref('');
    const globalUnitPrice = ref('');

    const isGenerating = ref(false);
    const generationStatus = ref(null);
    const generationError = ref(null);
    const showGenTraceback = ref(false);
    const validationData = ref(null);
    const validationWarnings = ref([]);
    const assetStatus = ref(null);

    // Google Sheets
    const isOnlineMode = ref(true);
    const showGoogleSheetsSettings = ref(false);
    const googleSheetId = ref('');
    const googleSheetName = ref('2026');

    const isSyncing = ref(false);
    const syncStatus = ref(null);
    const showConflictConfirm = ref(false);
    const conflictMessage = ref('');
    const isWideCargo = ref(false);
    const maxStackingLayers = ref(3);
    const isLookingUp = ref(false);

    // --- Computed properties ---
    const isNetMode = computed(() => assetStatus.value?.pricing_mode === 'net');

    const hasVariants = computed(() => (assetStatus.value?.variants?.length || 0) > 0);

    const assetConfigName = computed(() => {
        if (!assetStatus.value?.config_path) return 'Unknown';
        const path = assetStatus.value.config_path;
        return path.split(/[\\/]/).pop() || 'Unknown';
    });

    const summaryStats = computed(() => {
        const gt = validationData.value?.footer_data?.grand_total;
        if (gt) {
            return {
                total_pcs: Number(gt.col_qty_pcs) || 0,
                total_sqft: Number(gt.col_qty_sf) || 0,
                total_pallets: Number(gt.col_pallet_count) || 0
            };
        }
        return validationData.value?.database_export?.summary || null;
    });

    const weightStats = computed(() => {
        const gt = validationData.value?.footer_data?.grand_total;
        if (gt) {
            return {
                net: Number(gt.col_net) || 0,
                gross: Number(gt.col_gross) || 0,
                cbm: Number(gt.col_cbm) || 0
            };
        }
        if (!validationData.value?.database_export?.packing_list_items) return null;
        const items = validationData.value.database_export.packing_list_items;
        let net = 0; let gross = 0; let cbm = 0;
        items.forEach(item => {
            try { net += parseFloat(item.net) || 0; } catch { }
            try { gross += parseFloat(item.gross) || 0; } catch { }
            try { cbm += parseFloat(item.cbm) || 0; } catch { }
        });
        return { net, gross, cbm };
    });

    const detectedPalletDims = computed(() => {
        const data = validationData.value;
        if (!data) return { length: 1.2, width: 1.0 };
        
        const items = (data.multi_table || data.raw_data || []).flat();
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
        
        if (data.database_export?.packing_list_items) {
            for (const item of data.database_export.packing_list_items) {
                const cbmRaw = item.cbm_raw || item.col_cbm_raw || item.cbm || item.col_cbm;
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
        }
        
        return { length: 1.2, width: 1.0 };
    });

    const recommendedTruckInfo = computed(() => {
        const gross = weightStats.value?.gross || 0;
        const cbm = weightStats.value?.cbm || 0;
        const pallets = summaryStats.value?.total_pallets || 0;
        return recommendTruck(gross, cbm, pallets, isWideCargo.value, {
            maxStackingLayers: maxStackingLayers.value,
            palletLength: detectedPalletDims.value.length,
            palletWidth: detectedPalletDims.value.width
        });
    });

    const totalAmount = computed(() => {
        const gt = validationData.value?.footer_data?.grand_total;
        if (gt) {
            return Number(gt.col_amount || 0);
        }
        if (validationData.value?.database_export?.summary) {
            return Number(validationData.value.database_export.summary.total_amount || 0);
        }
        return 0;
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

    // --- Watchers ---
    watch(validationData, (newData) => {
        if (newData) {
            const desc = String(newData.footer_data?.grand_total?.col_desc || '').toUpperCase();
            const file = String(identifier.value || '').toUpperCase();
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

    // --- Actions ---
    const setRawFile = (file) => {
        rawFile = file || null;
        selectedFile.value = file ? { name: file.name } : null;

        // Reset state when new file is chosen
        uploadStatus.value = null;
        uploadError.value = null;
        showUploadTraceback.value = false;
        ignoreTareError.value = false;
        ignoreCbmError.value = false;
        processingComplete.value = false;
        validationData.value = null;
        assetStatus.value = null;
        selectedVariants.value = [];
        selectedTargets.value = [];
        priceAdjustments.value = [];
        refSourceStatus.value = null;
    };

    const addAdjustment = () => {
        priceAdjustments.value.push({ description: '', value: '' });
    };

    const removeAdjustment = (index) => {
        priceAdjustments.value.splice(index, 1);
    };

    const uploadFile = async () => {
        if (!rawFile) return;

        isUploading.value = true;
        uploadStatus.value = { type: 'info', message: 'Uploading and processing...' };
        uploadError.value = null;
        validationData.value = null;
        validationWarnings.value = [];

        const formData = new FormData();
        formData.append('file', rawFile);
        formData.append('ignore_tare', ignoreTareError.value);
        formData.append('ignore_cbm', ignoreCbmError.value);

        try {
            const response = await fetch('/api/upload', {
                method: 'POST',
                body: formData
            });

            const data = await response.json();

            if (response.ok) {
                uploadStatus.value = { type: 'success', message: 'File processed successfully!' };
                identifier.value = data.identifier;
                jsonPath.value = data.json_path;
                invoiceNo.value = data.default_inv_no || '';

                assetStatus.value = data.asset_status || null;

                if (data.asset_status?.variants?.length > 0) {
                    selectedVariants.value = data.asset_status.variants.map(v => v.suffix);
                    const suffixes = data.asset_status.variants.map(v => v.suffix);
                    if (suffixes.includes('_KH')) {
                        selectedTargets.value = ['KH_Standard'];
                    } else if (suffixes.includes('_VN')) {
                        selectedTargets.value = ['VN_Standard'];
                    } else {
                        const firstLoc = suffixes[0]?.replace('_', '') || 'Default';
                        selectedTargets.value = [`${firstLoc}_Standard`];
                    }
                } else {
                    selectedTargets.value = ['Default_Standard'];
                }

                if (data.warnings && data.warnings.length > 0) {
                    validationWarnings.value = data.warnings;
                    uploadStatus.value = { type: 'warning', message: 'File processed successfully, but with data corrections.' };
                } else {
                    validationWarnings.value = [];
                }

                processingComplete.value = true;
            } else {
                uploadError.value = {
                    message: data.error || 'Upload failed',
                    step: data.step || null,
                    traceback: data.traceback || null
                };
                uploadStatus.value = null;
            }
        } catch (error) {
            uploadError.value = {
                message: error.message || 'Network error occurred',
                step: null,
                traceback: null
            };
            uploadStatus.value = null;
        } finally {
            isUploading.value = false;
        }
    };

    const retryUpload = () => {
        uploadError.value = null;
        showUploadTraceback.value = false;
        uploadFile();
    };

    const ignoreTareAndRetry = () => {
        ignoreTareError.value = true;
        uploadFile();
    };

    const ignoreCbmAndRetry = () => {
        ignoreCbmError.value = true;
        uploadFile();
    };

    const validateAdjustments = () => {
        const validSet = [];
        for (const adj of priceAdjustments.value) {
            const desc = (adj.description || '').trim();
            const valRaw = String(adj.value || '').trim();

            if (desc === '' && valRaw === '') continue;

            if (valRaw === '') {
                adjustmentError.value = 'Please enter a value for all adjustments.';
                return { isValid: false, list: [] };
            }

            const parsed = Number(valRaw);
            if (isNaN(parsed) || !Number.isFinite(parsed)) {
                adjustmentError.value = `Invalid value for "${desc || 'adjustment'}".`;
                return { isValid: false, list: [] };
            }

            validSet.push([desc || 'Adjustment', parsed]);
        }

        adjustmentError.value = '';
        return { isValid: true, list: validSet };
    };

    const generateInvoice = async () => {
        const { isValid, list: validAdjustments } = validateAdjustments();
        if (!isValid) {
            generationStatus.value = null;
            generationError.value = {
                message: adjustmentError.value,
                step: 'Validation',
                traceback: null
            };
            return;
        }

        isGenerating.value = true;
        generationStatus.value = { type: 'info', message: 'Generating invoice, please wait...' };
        generationError.value = null;
        validationData.value = null;

        try {
            if (isOnlineMode.value && !invoiceRef.value) {
                generationStatus.value = { type: 'info', message: 'Resolving Ref No from Google Sheets...' };
                await lookupRefFromSheets(true);
            }

            const basePayload = {
                identifier: identifier.value,
                json_path: jsonPath.value,
                invoice_no: invoiceNo.value,
                invoice_date: invoiceDate.value,
                invoice_ref: invoiceRef.value,
                targets: selectedTargets.value,
                generate_standard: includeStandard.value,
                generate_custom: includeCustom.value,
                generate_daf: includeDAF.value,
                generate_kh: true,
                generate_vn: selectedVariants.value.includes('_VN'),
                split_sheets: splitSheets.value
            };

            if (isNetMode.value && globalUnitPrice.value) {
                basePayload.global_unit_price = parseFloat(globalUnitPrice.value);
            }

            if (validAdjustments.length > 0) {
                basePayload.price_adjustment = validAdjustments;
            }

            const response = await fetch('/api/generate', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(basePayload)
            });

            const data = await response.json();

            if (response.ok) {
                generationStatus.value = { type: 'success', message: 'Invoice generated successfully! Download starting...' };
                if (data.metadata) {
                    validationData.value = data.metadata;
                }
                if (data.metadata_error) {
                    console.error('[Generate] metadata_error:', data.metadata_error);
                    generationStatus.value = { type: 'warning', message: `Invoice generated, but metadata could not be loaded: ${data.metadata_error}. Google Sheets sync may push incorrect values.` };
                }
                if (data.files && data.files.length > 0) {
                    data.files.forEach(f => {
                        console.log('[Generate] Downloading file:', f);
                        const mimeType = f.mime_type || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
                        const binaryString = window.atob(f.content);
                        const bytes = new Uint8Array(binaryString.length);
                        for (let i = 0; i < binaryString.length; i++) {
                            bytes[i] = binaryString.charCodeAt(i);
                        }
                        const blob = new Blob([bytes], { type: mimeType });
                        const url = URL.createObjectURL(blob);

                        const link = document.createElement('a');
                        link.href = url;
                        link.download = f.filename || f.file_name || f.name;
                        document.body.appendChild(link);
                        link.click();
                        document.body.removeChild(link);
                        URL.revokeObjectURL(url);
                    });
                }
                syncStatus.value = null;
            } else {
                generationError.value = {
                    message: data.error || 'Generation failed',
                    details: data.details || [],
                    step: data.step || null,
                    traceback: data.traceback || null
                };
                generationStatus.value = null;
            }
        } catch (error) {
            generationError.value = {
                message: error.message || 'Network error occurred',
                step: null,
                traceback: null
            };
            generationStatus.value = null;
        } finally {
            isGenerating.value = false;
        }
    };

    const retryGeneration = () => {
        generationError.value = null;
        showGenTraceback.value = false;
        generateInvoice();
    };

    const copyError = async (errorObj) => {
        const errorText = `Error: ${errorObj.message}\n\nStep: ${errorObj.step || 'N/A'}\n\nTraceback:\n${errorObj.traceback || 'No traceback available'}`;
        try {
            await navigator.clipboard.writeText(errorText);
            alert('Error copied to clipboard!');
        } catch (err) {
            console.error('Failed to copy error:', err);
        }
    };

    const exportToSheets = async (forceOverride = false) => {
        isSyncing.value = true;
        syncStatus.value = { type: 'info', message: 'Syncing to Google Sheets...' };
        try {
            const grandTotal = validationData.value?.footer_data?.grand_total || {};
            const pallets = grandTotal.col_pallet_count ?? (summaryStats.value?.total_pallets || 0);
            let gross = grandTotal.col_gross ?? (weightStats.value?.gross || 0);
            const netWeight = grandTotal.col_net ?? (weightStats.value?.net || 0);
            const amount = grandTotal.col_amount ?? 0;

            if (typeof gross === 'string' && !isNaN(parseFloat(gross))) {
                gross = parseFloat(gross).toString();
            }

            if (!pallets && !netWeight && !amount) {
                syncStatus.value = { type: 'error', message: 'Cannot sync: pallet, weight, or amount data is missing. Please regenerate the invoice first.' };
                isSyncing.value = false;
                return;
            }

            const palletStr = `${pallets} PALLETS: ${gross}`;

            const payload = {
                invoice_no: invoiceNo.value || identifier.value,
                ref_no: invoiceRef.value || '',
                invoice_date: invoiceDate.value,
                pallet_str: palletStr,
                net_weight: parseFloat(netWeight) || 0,
                pallets: parseInt(pallets, 10) || 0,
                amount: parseFloat(amount) || 0,
                force_override: forceOverride === true,
                worksheet_name: googleSheetName.value || '2026'
            };

            if (googleSheetId.value && googleSheetId.value.trim() !== '') {
                payload.spreadsheet_id = googleSheetId.value.trim();
            }

            const response = await fetch('/api/sheets/export', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });

            const data = await response.json();

            if (!response.ok) {
                syncStatus.value = { type: 'error', message: data.error || 'Failed to sync to Google Sheets.' };
            } else if (data.action === 'conflict') {
                syncStatus.value = null;
                conflictMessage.value = data.message;
                showConflictConfirm.value = true;
            } else {
                syncStatus.value = { type: 'success', message: data.message || 'Successfully synced to Google Sheets!' };
            }
        } catch (error) {
            syncStatus.value = { type: 'error', message: 'Network error while syncing.' };
        } finally {
            isSyncing.value = false;
        }
    };

    const confirmOverride = async () => {
        showConflictConfirm.value = false;
        await exportToSheets(true);
    };

    const cancelOverride = () => {
        showConflictConfirm.value = false;
        syncStatus.value = { type: 'info', message: 'Sync cancelled.' };
    };

    const lookupRefFromSheets = async (silent = false) => {
        if (!invoiceNo.value) return;
        isLookingUp.value = true;
        try {
            const payload = {
                invoice_no: invoiceNo.value,
                worksheet_name: googleSheetName.value || '2026'
            };
            if (googleSheetId.value && googleSheetId.value.trim() !== '') {
                payload.spreadsheet_id = googleSheetId.value.trim();
            }

            const response = await fetch('/api/sheets/resolve_ref', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });

            if (response.ok) {
                const data = await response.json();
                invoiceRef.value = data.ref_no;
                refSourceStatus.value = {
                    type: data.found ? 'found' : 'new',
                    message: data.found ? '✓ Pulled from Sheets' : '✨ Auto-generated'
                };
            } else {
                const err = await response.json();
                if (!silent) alert(`Lookup Failed: ${err.error || 'Check console'}`);
            }
        } catch (error) {
            if (!silent) alert(`Network Error: ${error.message}`);
        } finally {
            isLookingUp.value = false;
        }
    };

    return {
        // state
        selectedFile,
        isUploading,
        uploadStatus,
        uploadError,
        showUploadTraceback,
        ignoreTareError,
        ignoreCbmError,
        processingComplete,
        identifier,
        jsonPath,
        invoiceNo,
        invoiceDate,
        invoiceRef,
        refSourceStatus,
        includeStandard,
        includeCustom,
        includeDAF,
        selectedVariants,
        splitSheets,
        selectedTargets,
        priceAdjustments,
        adjustmentError,
        globalUnitPrice,
        isGenerating,
        generationStatus,
        generationError,
        showGenTraceback,
        validationData,
        validationWarnings,
        assetStatus,
        isOnlineMode,
        showGoogleSheetsSettings,
        googleSheetId,
        googleSheetName,
        isSyncing,
        syncStatus,
        showConflictConfirm,
        conflictMessage,
        isWideCargo,
        maxStackingLayers,
        isLookingUp,

        // computed
        isNetMode,
        hasVariants,
        assetConfigName,
        summaryStats,
        weightStats,
        detectedPalletDims,
        recommendedTruckInfo,
        totalAmount,

        // actions
        setRawFile,
        addAdjustment,
        removeAdjustment,
        uploadFile,
        retryUpload,
        ignoreTareAndRetry,
        ignoreCbmAndRetry,
        generateInvoice,
        retryGeneration,
        copyError,
        exportToSheets,
        confirmOverride,
        cancelOverride,
        lookupRefFromSheets
    };
});
