import { useTemplateExtractorStore } from '../../stores/templateExtractorStore.js';

export default {
    name: 'GlobalMappingsManager',
    template: `
        <div class="bg-slate-800/80 backdrop-blur-md border border-slate-700/50 shadow-2xl rounded-2xl p-8 mb-8 mt-8">
            <div class="flex justify-between items-center cursor-pointer select-none" @click="store.showMappings = !store.showMappings">
                <h2>Manage Global Mappings</h2>
                <span>{{ store.showMappings ? '▲ Collapse' : '▼ Expand' }}</span>
            </div>
            
            <div v-if="store.showMappings" class="mt-4">
                <p class="text-secondary mb-4">
                    View and edit the globally recognized mappings. These are used to automatically match headers and sheets in templates.
                </p>
                
                <div class="flex gap-4 mb-4">
                    <select :value="store.activeMappingType" @change="store.switchMappingType($event.target.value)" class="flex-none bg-slate-900 border border-slate-700 rounded-lg px-4 py-2 text-slate-100 focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 transition-all w-72 font-bold">
                        <option value="header_text_mappings">Header Mappings</option>
                        <option value="sheet_mappings">Sheet Mappings</option>
                        <option value="shipping_header_map">Shipping Header Map</option>
                        <option value="footer_label_mappings">Footer Labels (Total)</option>
                    </select>
                    <div class="flex-1 relative">
                        <input type="text" v-model="store.mappingSearch" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-4 py-2 text-slate-100 focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 transition-all" placeholder="Search..." />
                    </div>
                </div>

                <!-- Add New Mapping Row -->
                <div v-if="store.activeMappingType === 'sheet_mappings'" style="display: grid; grid-template-columns: 1fr 1fr auto; gap: 0.5rem; align-items: center;" class="mb-4 p-2 bg-emerald-500/5 border border-emerald-500/30 border-dashed rounded-md">
                    <input type="text" v-model="store.newMappingKey" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-4 py-2 text-slate-100 focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 transition-all" placeholder="Sheet Name (e.g. INV)" />
                    <select v-model="store.newMappingVal" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-4 py-2 text-slate-100 focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 transition-all">
                        <option value="" disabled selected>Select classification...</option>
                        <option value="aggregation">Aggregation (Invoice / Contract)</option>
                        <option value="processed_tables">Processed Tables (Packing List)</option>
                    </select>
                    <button class="px-6 py-2 bg-blue-500 hover:bg-blue-600 text-white font-medium rounded-lg shadow-sm transition-colors disabled:opacity-50 disabled:cursor-not-allowed min-w-[100px]" @click.prevent="store.addNewMapping" :disabled="!store.newMappingKey || !store.newMappingVal">Add</button>
                </div>
                <div v-else style="display: grid; grid-template-columns: 1fr 1fr auto; gap: 0.5rem; align-items: center;" class="mb-4 p-2 bg-emerald-500/5 border border-emerald-500/30 border-dashed rounded-md">
                    <input type="text" v-model="store.newMappingKey" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-4 py-2 text-slate-100 focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 transition-all" :placeholder="store.activeMappingType === 'shipping_header_map' ? 'New Keyword (e.g. PO)' : (store.activeMappingType === 'footer_label_mappings' ? 'New Footer Label (e.g. GRAND TOTAL)' : 'New Input Text (e.g. Qty(SF))')" />
                    
                    <input v-if="store.activeMappingType === 'footer_label_mappings'" type="text" v-model="store.newMappingVal" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-4 py-2 text-slate-100 focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 transition-all" placeholder="Auto-filled" :disabled="true" />
                    <select v-else v-model="store.newMappingVal" class="w-full bg-slate-900 border border-slate-700 rounded-lg px-4 py-2 text-slate-100 focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 transition-all">
                        <option value="" disabled selected>Select system field...</option>
                        <option v-for="opt in store.systemOptions" :value="opt.id" :key="opt.id">{{ opt.label }} ({{ opt.id }})</option>
                    </select>
                    
                    <button class="px-6 py-2 bg-blue-500 hover:bg-blue-600 text-white font-medium rounded-lg shadow-sm transition-colors disabled:opacity-50 disabled:cursor-not-allowed min-w-[100px]" @click.prevent="store.addNewMapping" :disabled="!store.newMappingKey || !store.newMappingVal">Add</button>
                </div>

                <div class="max-h-[400px] overflow-y-auto border border-white/10 rounded-md p-2">
                    <div class="mapping-grid grid gap-2">
                        <!-- Header Row -->
                        <div v-if="store.activeMappingType === 'sheet_mappings'" style="display: grid; grid-template-columns: 1fr 1fr auto; gap: 0.5rem; align-items: center;" class="font-bold p-2 border-b border-white/10">
                            <div>Sheet Name</div>
                            <div>Classification</div>
                            <div class="w-20 text-center">Action</div>
                        </div>
                        <div v-else style="display: grid; grid-template-columns: 1fr 1fr auto; gap: 0.5rem; align-items: center;" class="font-bold p-2 border-b border-white/10">
                            <div>{{ store.activeMappingType === 'shipping_header_map' ? 'Keyword' : (store.activeMappingType === 'footer_label_mappings' ? 'Footer Target Text' : 'Original Text (Excel)') }}</div>
                            <div>{{ store.activeMappingType === 'shipping_header_map' ? 'Mapped Target (System)' : (store.activeMappingType === 'footer_label_mappings' ? 'Type' : 'Mapped Target (System)') }}</div>
                            <div class="w-20 text-center">Action</div>
                        </div>
                        
                        <template v-if="store.activeMappingType === 'sheet_mappings'">
                            <div v-for="(val, sheetName) in store.filteredMappings" :key="sheetName" style="display: grid; grid-template-columns: 1fr 1fr auto; gap: 0.5rem; align-items: center;" class="bg-white/5 p-2 rounded">
                                <input type="text" :value="sheetName" @change="store.updateMappingHeader(sheetName, $event.target.value)" class="w-full h-9 bg-slate-900 border border-slate-700 rounded px-3 text-slate-100 focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 transition-all text-sm" />
                                
                                <select :value="val" @change="store.updateMappingColId(sheetName, $event.target.value)" class="w-full h-9 bg-slate-900 border border-slate-700 rounded px-3 text-slate-100 focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 transition-all text-sm">
                                    <option value="aggregation">Aggregation (Invoice / Contract)</option>
                                    <option value="processed_tables">Processed Tables (Packing List)</option>
                                </select>
                                
                                <button class="h-9 px-4 bg-red-500/80 hover:bg-red-500 text-white rounded cursor-pointer transition-colors w-full text-sm font-medium" @click="store.deleteMapping(sheetName)">Delete</button>
                            </div>
                        </template>
                        <template v-else>
                            <div v-for="(colId, headerText) in store.filteredMappings" :key="headerText" style="display: grid; grid-template-columns: 1fr 1fr auto; gap: 0.5rem; align-items: center;" class="bg-white/5 p-2 rounded">
                                <input type="text" :value="headerText" @change="store.updateMappingHeader(headerText, $event.target.value)" class="w-full h-9 bg-slate-900 border border-slate-700 rounded px-3 text-slate-100 focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 transition-all text-sm" />
                                
                                <input v-if="store.activeMappingType === 'footer_label_mappings'" type="text" :value="colId" @change="store.updateMappingColId(headerText, $event.target.value)" class="w-full h-9 bg-slate-900 border border-slate-700 rounded px-3 text-slate-100 focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 transition-all text-sm" :disabled="true" />
                                
                                <select v-else :value="colId" @change="store.updateMappingColId(headerText, $event.target.value)" class="w-full h-9 bg-slate-900 border border-slate-700 rounded px-3 text-slate-100 focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 transition-all text-sm">
                                    <option v-for="opt in store.systemOptions" :value="opt.id" :key="opt.id">
                                        {{ opt.label }} ({{ opt.id }})
                                    </option>
                                    <option v-if="!store.systemOptions.find(o => o.id === colId)" :value="colId">{{ colId }} (Unknown)</option>
                                </select>
                                
                                <button class="h-9 px-4 bg-red-500/80 hover:bg-red-500 text-white rounded cursor-pointer transition-colors w-full text-sm font-medium" @click="store.deleteMapping(headerText)">Delete</button>
                            </div>
                        </template>
                        <div v-if="Object.keys(store.filteredMappings).length === 0" class="p-4 text-center text-secondary">
                            No mappings found matching your search.
                        </div>
                    </div>
                </div>
                
                <div class="mt-4 flex justify-end">
                    <button class="px-6 py-3 bg-emerald-500 hover:bg-emerald-400 text-white font-bold rounded-full shadow-lg shadow-emerald-500/20 transition-all transform hover:-translate-y-0.5 disabled:opacity-50 disabled:cursor-not-allowed" @click="store.saveMappings" :disabled="store.isSavingMappings">
                        {{ store.isSavingMappings ? 'Saving...' : 'Save Mappings' }}
                    </button>
                </div>
                
                <div v-if="store.mappingStatusMessage" :class="['status-box', store.mappingStatusType]" class="mt-4">
                    {{ store.mappingStatusMessage }}
                </div>
            </div>
        </div>
    `,
    setup() {
        const store = useTemplateExtractorStore();
        return { store };
    }
};
