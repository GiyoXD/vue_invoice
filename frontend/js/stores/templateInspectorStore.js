import { defineStore } from 'pinia';
import { ref, computed } from 'vue';

export const useTemplateInspectorStore = defineStore('templateInspector', () => {
    // --- State ---
    const templates = ref([]);
    const searchQuery = ref("");
    const selectedTemplateName = ref(null);
    const currentTemplate = ref(null);
    const currentSheetName = ref(null);

    // --- Getters ---
    const filteredTemplates = computed(() => {
        const q = (searchQuery.value || "").trim().toLowerCase();
        if (!q) return templates.value;
        return templates.value.filter(t => {
            const nameMatch = (t.name || "").toLowerCase().includes(q);
            const bundleMatch = (t.bundle_name || "").toLowerCase().includes(q);
            return nameMatch || bundleMatch;
        });
    });

    const templateNotes = computed(() => currentTemplate.value?.notes || "");
    const clientProfile = computed(() => currentTemplate.value?.client_profile || null);
    const templateLayout = computed(() => currentTemplate.value?.template_layout || {});
    const currentTemplateFingerprint = computed(() => currentTemplate.value?.fingerprint || null);

    // --- Actions ---
    const fetchTemplates = async () => {
        try {
            const res = await fetch('/api/templates');
            if (res.ok) {
                templates.value = await res.json();
            }
        } catch (e) {
            console.error("Failed to fetch templates", e);
        }
    };

    const loadTemplate = async (t, preserveState = false) => {
        selectedTemplateName.value = t.name;
        if (!preserveState) {
            currentTemplate.value = null; // Clear immediately to prevent showing old data
        }
        try {
            const url = `/api/template/view?customer_code=${encodeURIComponent(t.customer_code)}&locale=${encodeURIComponent(t.locale)}&_t=${Date.now()}`;
            const res = await fetch(url);
            if (res.ok) {
                currentTemplate.value = await res.json();
                const sheets = Object.keys(currentTemplate.value?.template_layout || {});

                if (!preserveState || !currentSheetName.value || !sheets.includes(currentSheetName.value)) {
                    if (sheets.length > 0) currentSheetName.value = sheets[0];
                    else currentSheetName.value = null;
                }
            }
        } catch (e) {
            console.error("Failed to load template", e);
        }
    };

    const deleteTemplate = async () => {
        if (!selectedTemplateName.value) return false;

        const t = templates.value.find(tmpl => tmpl.name === selectedTemplateName.value);
        if (!t) return false;

        if (!confirm(`WARNING: Are you sure you want to permanently delete the template for '${t.customer_code}' (${t.locale})?`)) {
            return false;
        }
        try {
            const url = `/api/template/${encodeURIComponent(t.customer_code)}?locale=${encodeURIComponent(t.locale)}`;
            const res = await fetch(url, {
                method: 'DELETE'
            });
            if (res.ok) {
                currentTemplate.value = null;
                selectedTemplateName.value = null;
                currentSheetName.value = null;
                await fetchTemplates();
                alert(`Template bundle deleted successfully.`);
                return true;
            } else {
                const data = await res.json();
                alert(`Failed to delete template: ${data.error || res.statusText}`);
            }
        } catch (e) {
            console.error("Error deleting template", e);
            alert('An error occurred while deleting the template.');
        }
        return false;
    };

    const saveNotes = async (notesText) => {
        if (!selectedTemplateName.value) return false;
        const t = templates.value.find(tmpl => tmpl.name === selectedTemplateName.value);
        if (!t) return false;
        try {
            const res = await fetch('/api/template/notes', {
                method: 'PATCH',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    customer_code: t.customer_code,
                    locale: t.locale,
                    notes: notesText
                })
            });

            if (res.ok) {
                if (currentTemplate.value) {
                    currentTemplate.value.notes = notesText;
                }
                return true;
            } else {
                const data = await res.json();
                alert(`Failed to save notes: ${data.error || 'Unknown error'}`);
            }
        } catch (e) {
            console.error("Error saving notes", e);
            alert("Failed to save notes. See console for details.");
        }
        return false;
    };

    const saveCellOverrides = async (cellAddress, standardValue, dafValue) => {
        if (!selectedTemplateName.value || !currentSheetName.value) {
            return { success: false, error: 'Missing template or sheet selection' };
        }
        const t = templates.value.find(tmpl => tmpl.name === selectedTemplateName.value);
        if (!t) return { success: false, error: 'Template not found' };

        try {
            const res = await fetch('/api/template/cell', {
                method: 'PATCH',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    customer_code: t.customer_code,
                    locale: t.locale,
                    sheet_name: currentSheetName.value,
                    cell_address: cellAddress,
                    overrides: {
                        standard: standardValue,
                        daf: dafValue
                    }
                })
            });
            const data = await res.json();
            if (res.ok) {
                await loadTemplate(t, true);
                return { success: true };
            } else {
                return { success: false, error: data.error || 'Save failed' };
            }
        } catch (e) {
            console.error("Error saving cell override", e);
            return { success: false, error: e.message };
        }
    };

    return {
        templates,
        searchQuery,
        selectedTemplateName,
        currentTemplate,
        currentSheetName,
        filteredTemplates,
        templateNotes,
        clientProfile,
        templateLayout,
        currentTemplateFingerprint,
        fetchTemplates,
        loadTemplate,
        deleteTemplate,
        saveNotes,
        saveCellOverrides
    };
});
