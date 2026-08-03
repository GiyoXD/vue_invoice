import { defineStore } from 'pinia';
import { ref, reactive, computed } from 'vue';

export const useTemplateExtractorStore = defineStore('templateExtractor', () => {
    const currentStep = ref(1);
    const selectedFiles = ref([]);
    let rawFiles = []; // Non-reactive list for native File objects
    const singleFileSuffix = ref("KH");
    const isProcessing = ref(false);
    const statusMessage = ref("");
    const statusType = ref("info");
    const ignoreMissingDescription = ref(false);

    const showMappings = ref(false);
    const globalMappings = ref({});
    const mappingSearch = ref("");
    const isSavingMappings = ref(false);
    const mappingStatusMessage = ref("");
    const mappingStatusType = ref("info");
    const activeMappingType = ref("header_text_mappings");
    const newMappingKey = ref("");
    const newMappingVal = ref("");
    const newMappingType = ref("aggregation");
    const footerKeywords = ref([]);

    // Data
    const fileTokens = ref([]); // Array of { filename, missingHeaders }
    const allMissingHeaders = ref([]); // Deduplicated list across all files
    const allMissingFooters = ref([]); // Deduplicated footers
    const allUnrecognizedSheets = ref([]); // Unrecognized sheet names
    const filePrefix = ref("");
    const userMappings = reactive({});
    const confirmedHeaders = ref([]);
    const confirmedFooters = ref([]);
    const proactiveWarnings = ref([]); // warnings from analysis
    const bundlePath = ref("");
    const generatedPrefixes = ref([]);
    const pricingMode = ref('standard'); // 'standard' or 'net'

    const systemOptions = ref([]);

    /**
     * Returns true when user uploaded 2 files (KH + VN mode).
     */
    const isDualMode = computed(() => selectedFiles.value.length >= 2);

    // Load options
    const fetchOptions = async () => {
        try {
            const res = await fetch('/api/blueprint/options');
            if (res.ok) {
                systemOptions.value = await res.json();
            }
        } catch (e) {
            console.error("Failed to fetch options", e);
        }
    };

    const fetchMappings = async () => {
        try {
            const res = await fetch(`/api/blueprint/mappings?mapping_type=${activeMappingType.value}`);
            if (res.ok) {
                globalMappings.value = await res.json();
            }
        } catch (e) {
            console.error("Failed to fetch mappings", e);
        }
    };

    const fetchFooterMappings = async () => {
        try {
            const res = await fetch(`/api/blueprint/mappings?mapping_type=footer_label_mappings`);
            if (res.ok) {
                const data = await res.json();
                footerKeywords.value = Object.keys(data).map(k => k.toUpperCase());
            }
        } catch (e) {
            console.error("Failed to fetch footer mappings", e);
        }
    };

    const filteredMappings = computed(() => {
        if (!mappingSearch.value) return globalMappings.value;
        const term = mappingSearch.value.toLowerCase();
        const result = {};
        for (const [key, val] of Object.entries(globalMappings.value)) {
            if (key.toLowerCase().includes(term) || (typeof val === 'string' && val.toLowerCase().includes(term))) {
                result[key] = val;
            }
        }
        return result;
    });

    const updateMappingHeader = (oldKey, newKey) => {
        const trimmed = newKey.trim();
        if (oldKey === trimmed || !trimmed) return;
        if (globalMappings.value[trimmed]) {
            alert("Header mapping already exists.");
            return;
        }

        const newMappings = { ...globalMappings.value };
        newMappings[trimmed] = newMappings[oldKey];
        delete newMappings[oldKey];
        globalMappings.value = newMappings;
    };

    const updateMappingColId = (key, newColId) => {
        globalMappings.value = { ...globalMappings.value, [key]: newColId };
    };

    const deleteMapping = (key) => {
        if (confirm(`Are you sure you want to delete the mapping for "${key}"?`)) {
            const newMappings = { ...globalMappings.value };
            delete newMappings[key];
            globalMappings.value = newMappings;
        }
    };

    const addNewMapping = () => {
        if (newMappingKey.value && newMappingVal.value) {
            globalMappings.value = {
                ...globalMappings.value,
                [newMappingKey.value]: newMappingVal.value
            };
            newMappingKey.value = "";
            if (activeMappingType.value === 'sheet_mappings') {
                newMappingVal.value = "";
            } else {
                newMappingVal.value = activeMappingType.value === 'footer_label_mappings' ? 'Footer Keyword' : '';
            }
        }
    };

    const switchMappingType = async (type) => {
        activeMappingType.value = type;
        await fetchMappings();
        mappingStatusMessage.value = "";
        newMappingKey.value = "";
        
        if (type === 'sheet_mappings') {
            newMappingVal.value = "";
        } else {
            newMappingVal.value = type === 'footer_label_mappings' ? 'Footer Keyword' : '';
        }
        
        if (type === 'footer_label_mappings') {
            footerKeywords.value = Object.keys(globalMappings.value).map(k => k.toUpperCase());
        }
    };

    const saveMappings = async () => {
        isSavingMappings.value = true;
        mappingStatusMessage.value = "Saving mappings...";
        mappingStatusType.value = "info";
        try {
            const res = await fetch('/api/blueprint/mappings', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    mapping_type: activeMappingType.value,
                    mappings: globalMappings.value
                })
            });
            if (res.ok) {
                mappingStatusMessage.value = "Mappings saved successfully!";
                mappingStatusType.value = "success";
                
                if (activeMappingType.value === 'footer_label_mappings') {
                    footerKeywords.value = Object.keys(globalMappings.value).map(k => k.toUpperCase());
                }
                
                setTimeout(() => { mappingStatusMessage.value = ""; }, 3000);
            } else {
                const data = await res.json();
                throw new Error(data.error || "Save failed");
            }
        } catch (e) {
            mappingStatusType.value = "error";
            mappingStatusMessage.value = e.message;
        } finally {
            isSavingMappings.value = false;
        }
    };

    /**
     * Handles file input change. Accepts up to 2 files.
     */
    const handleFileUpload = (e) => {
        const files = Array.from(e.target.files).slice(0, 2);
        rawFiles = files;
        selectedFiles.value = files.map(f => ({ name: f.name }));
        singleFileSuffix.value = "KH";
        statusMessage.value = "";
    };

    const toggleMapping = (headerText) => {
        if (confirmedHeaders.value.includes(headerText)) {
            confirmedHeaders.value = confirmedHeaders.value.filter(h => h !== headerText);
        } else {
            if (!userMappings[headerText]) {
                alert("Please select a field first.");
                return;
            }
            confirmedHeaders.value.push(headerText);
        }
    };

    const toggleFooter = (footerText) => {
        if (confirmedFooters.value.includes(footerText)) {
            confirmedFooters.value = confirmedFooters.value.filter(f => f !== footerText);
        } else {
            confirmedFooters.value.push(footerText);
        }
    };

    const analyzeFiles = async () => {
        if (rawFiles.length === 0) return;
        isProcessing.value = true;
        statusMessage.value = "Scanning template structure...";
        allMissingHeaders.value = [];
        allMissingFooters.value = [];
        allUnrecognizedSheets.value = [];
        fileTokens.value = [];

        // Refresh footer mappings before analysis to prevent stale keywords
        await fetchFooterMappings();

        try {
            const headerSet = new Set();
            const footerSet = new Set();
            const warningSet = new Set();
            const unrecognizedSheetSet = new Set();

            for (const file of rawFiles) {
                const formData = new FormData();
                formData.append('file', file);

                const res = await fetch(`/api/template/analyze?ignore_missing_description=${ignoreMissingDescription.value}`, { method: 'POST', body: formData });
                const data = await res.json();

                if (!res.ok) {
                    throw new Error(data.error || `Analysis failed for ${file.name}`);
                }

                // Collect file token info
                fileTokens.value.push({
                    filename: data.temp_filename,
                    missingHeaders: (data.missing_headers || []).map(h => h.text)
                });

                // Collect unique missing headers across all files
                for (const h of (data.missing_headers || [])) {
                    headerSet.add(h.text);
                }
                
                // Collect missing footers with case-insensitive deduplication
                for (const f of (data.missing_footers || [])) {
                    const fUpper = f.toUpperCase().trim();
                    // Only add if not already in global mappings
                    if (!footerKeywords.value.includes(fUpper)) {
                        // Check if already in our current set (case-insensitive)
                        const exists = Array.from(footerSet).some(existing => existing.toUpperCase().trim() === fUpper);
                        if (!exists) {
                            footerSet.add(f);
                        }
                    }
                }

                // Collect unrecognized sheets
                if (data.unrecognized_sheets && data.unrecognized_sheets.length > 0) {
                    data.unrecognized_sheets.forEach(s => unrecognizedSheetSet.add(s));
                }

                // Collect proactive warnings
                if (data.warnings && data.warnings.length > 0) {
                    data.warnings.forEach(w => warningSet.add(w));
                }
            }

            allMissingHeaders.value = Array.from(headerSet);
            allMissingFooters.value = Array.from(footerSet);
            allUnrecognizedSheets.value = Array.from(unrecognizedSheetSet);
            proactiveWarnings.value = Array.from(warningSet);

            if (allMissingHeaders.value.length > 0 || allMissingFooters.value.length > 0 || allUnrecognizedSheets.value.length > 0) {
                statusMessage.value = "Unmapped fields or unrecognized sheets found. Please review.";
            } else if (proactiveWarnings.value.length > 0) {
                statusMessage.value = "Template analyzed with warnings.";
            } else {
                statusMessage.value = "Structure looks clean!";
            }

            // Suggest prefix from first filename
            filePrefix.value = selectedFiles.value[0].name.split('.')[0];
            currentStep.value = 2;

        } catch (e) {
            statusType.value = "error";
            statusMessage.value = e.message;
        } finally {
            isProcessing.value = false;
        }
    };

    /**
     * Generate template(s).
     * Single file: 1 API call. Dual files: 2 API calls into same bundle_dir_name.
     */
    const generateTemplate = async () => {
        if (!filePrefix.value) {
            alert("Please enter a prefix");
            return;
        }
        isProcessing.value = true;
        statusMessage.value = "Generating bundle configuration...";
        statusType.value = "info";
        generatedPrefixes.value = [];

        try {
            // Collect confirmed mappings
            const finalMappings = {};
            for (const [key, value] of Object.entries(userMappings)) {
                if (confirmedHeaders.value.includes(key)) {
                    finalMappings[key] = value;
                }
            }

            if (isDualMode.value) {
                // --- DUAL MODE: 2 files → KH + VN ---
                const suffixes = ['KH', 'VN'];
                const baseName = filePrefix.value;

                for (let i = 0; i < Math.min(fileTokens.value.length, 2); i++) {
                    const localeVal = suffixes[i];
                    const displayPrefix = `${baseName}_${localeVal}`;
                    statusMessage.value = `Generating ${displayPrefix}...`;

                    const res = await fetch('/api/template/generate', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({
                            customer_code: baseName,
                            locale: localeVal,
                            user_mappings: finalMappings,
                            temp_filename: fileTokens.value[i].filename,
                            bundle_dir_name: baseName,
                            confirmed_footers: confirmedFooters.value,
                            pricing_mode: pricingMode.value,
                            ignore_missing_description: ignoreMissingDescription.value
                        })
                    });
                    const data = await res.json();

                    if (!res.ok) {
                        throw new Error(data.error || `Generation failed for ${displayPrefix}`);
                    }

                    generatedPrefixes.value.push(displayPrefix);
                    bundlePath.value = data.bundle_path || '';
                }

                currentStep.value = 3;

            } else {
                // --- SINGLE MODE: 1 file ---
                const baseName = filePrefix.value;
                const localeVal = singleFileSuffix.value;
                const displayPrefix = `${baseName}_${localeVal}`;

                const res = await fetch('/api/template/generate', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        customer_code: baseName,
                        locale: localeVal,
                        user_mappings: finalMappings,
                        temp_filename: fileTokens.value[0].filename,
                        bundle_dir_name: baseName,
                        confirmed_footers: confirmedFooters.value,
                        pricing_mode: pricingMode.value,
                        ignore_missing_description: ignoreMissingDescription.value
                    })
                });
                const data = await res.json();

                if (!res.ok) {
                    throw new Error(data.error || "Generation failed");
                }

                generatedPrefixes.value.push(displayPrefix);
                bundlePath.value = data.bundle_path || '';
                currentStep.value = 3;
            }

        } catch (e) {
            statusType.value = "error";
            statusMessage.value = e.message;
        } finally {
            isProcessing.value = false;
        }
    };

    const resetFlow = () => {
        currentStep.value = 1;
        selectedFiles.value = [];
        rawFiles = [];
        singleFileSuffix.value = "KH";
        filePrefix.value = "";
        allMissingHeaders.value = [];
        statusMessage.value = "";
        bundlePath.value = "";
        generatedPrefixes.value = [];
        fileTokens.value = [];
        for (const prop of Object.getOwnPropertyNames(userMappings)) {
            delete userMappings[prop];
        }
        confirmedHeaders.value = [];
        confirmedFooters.value = [];
        allMissingFooters.value = [];
        allUnrecognizedSheets.value = [];
        proactiveWarnings.value = [];
        ignoreMissingDescription.value = false;
    };

    const forceAnalyze = async () => {
        ignoreMissingDescription.value = true;
        statusMessage.value = "";
        statusType.value = "info";
        await analyzeFiles();
    };

    return {
        currentStep,
        selectedFiles,
        singleFileSuffix,
        isDualMode,
        isProcessing,
        statusMessage,
        statusType,
        ignoreMissingDescription,
        forceAnalyze,
        handleFileUpload,
        analyzeFiles,
        generateTemplate,
        resetFlow,
        filePrefix,
        allMissingHeaders,
        allMissingFooters,
        allUnrecognizedSheets,
        userMappings,
        confirmedHeaders,
        confirmedFooters,
        toggleMapping,
        toggleFooter,
        systemOptions,
        bundlePath,
        generatedPrefixes,
        showMappings,
        globalMappings,
        mappingSearch,
        isSavingMappings,
        mappingStatusMessage,
        mappingStatusType,
        filteredMappings,
        updateMappingHeader,
        updateMappingColId,
        deleteMapping,
        saveMappings,
        activeMappingType,
        switchMappingType,
        newMappingKey,
        newMappingVal,
        newMappingType,
        addNewMapping,
        proactiveWarnings,
        pricingMode,
        fetchOptions,
        fetchMappings,
        fetchFooterMappings
    };
});
