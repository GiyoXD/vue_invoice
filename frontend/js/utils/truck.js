export const TRUCKS = [
    { 
        name: '3T', 
        maxWeight: 3000, 
        maxCbm: 12, 
        maxPallets: 4, 
        maxPalletsWide: 2, 
        length: 4.2, 
        width: 1.8, 
        height: 1.8,
        displayName: '3T Truck 🚛', 
        color: 'text-indigo-400 bg-indigo-500/10 border-indigo-500/30' 
    },
    { 
        name: '5T', 
        maxWeight: 5000, 
        maxCbm: 20, 
        maxPallets: 6, 
        maxPalletsWide: 3, 
        length: 5.0, 
        width: 2.1, 
        height: 2.1,
        displayName: '5T Truck 🚛', 
        color: 'text-sky-400 bg-sky-500/10 border-sky-500/30' 
    },
    { 
        name: '8T', 
        maxWeight: 8000, 
        maxCbm: 32, 
        maxPallets: 10, 
        maxPalletsWide: 5, 
        length: 7.5, 
        width: 2.2, 
        height: 2.2,
        displayName: '8T Truck 🚛', 
        color: 'text-blue-400 bg-blue-500/10 border-blue-500/30' 
    },
    { 
        name: '20GP', 
        maxWeight: 20000, 
        maxCbm: 28, 
        maxPallets: 10, 
        maxPalletsWide: 4, 
        length: 5.9, 
        width: 2.35, 
        height: 2.39,
        displayName: '20GP Container 🚢', 
        color: 'text-emerald-400 bg-emerald-500/10 border-emerald-500/30' 
    },
    { 
        name: '40HC', 
        maxWeight: 25000, 
        maxCbm: 68, 
        maxPallets: 21, 
        maxPalletsWide: 9, 
        length: 12.03, 
        width: 2.35, 
        height: 2.69,
        displayName: '40HC Container 🚢', 
        color: 'text-teal-400 bg-teal-500/10 border-teal-500/30' 
    },
];

export function recommendTruck(weightKg, cbm = 0, pallets = 0, isWideCargo = false, options = {}) {
    const numericWeight = weightKg !== null && weightKg !== undefined && !isNaN(weightKg) ? Number(weightKg) : 0;
    const numericCbm = cbm !== null && cbm !== undefined && !isNaN(cbm) ? Number(cbm) : 0;
    const numericPallets = pallets !== null && pallets !== undefined && !isNaN(pallets) ? Number(pallets) : 0;

    const maxStackingLayers = options.maxStackingLayers !== undefined ? Number(options.maxStackingLayers) : 3;
    const palletLength = options.palletLength !== undefined ? Number(options.palletLength) : 1.2;
    const palletWidth = options.palletWidth !== undefined ? Number(options.palletWidth) : 1.0;
    const clearance = 0.08; // 8cm safety gap between pallets / walls

    for (const truck of TRUCKS) {
        // 1. Check weight constraint
        if (numericWeight > truck.maxWeight) continue;

        // 2. Check CBM constraint
        if (numericCbm > truck.maxCbm) continue;

        // 3. Perform 2D packing check with stacking limits
        // Check if we can load 2-wide: (palletWidth * 2) + clearance <= truck.width
        const canLoadTwoWide = !isWideCargo && ((palletWidth * 2) + (clearance * 2) <= truck.width);
        const rowWidthFactor = canLoadTwoWide ? 2 : 1;

        // How many floor spots are needed given the stacking layers limit?
        const floorSpotsNeeded = Math.ceil(numericPallets / maxStackingLayers);

        // How many rows along the length of the truck are needed?
        const rowsNeeded = Math.ceil(floorSpotsNeeded / rowWidthFactor);

        // Calculate the total length required
        const totalLengthRequired = rowsNeeded * (palletLength + clearance);

        // Check if it fits in the truck's length
        if (totalLengthRequired <= truck.length) {
            // Also keep standard check as a safety fallback for default inputs
            const allowedPallets = isWideCargo ? truck.maxPalletsWide : truck.maxPallets;
            if (maxStackingLayers === 3 && palletLength <= 1.2 && palletWidth <= 1.0 && numericPallets > allowedPallets) {
                continue;
            }
            
            return {
                ...truck,
                description: `${truck.displayName} (Max ${truck.maxWeight.toLocaleString()}kg, ${truck.maxCbm} CBM, fits ${numericPallets} Pallets in ${rowsNeeded} row(s) [Max Stack: ${maxStackingLayers}-high])`
            };
        }
    }

    // Over Limit fallback
    return {
        name: 'Over Limit',
        maxWeight: Infinity,
        maxCbm: Infinity,
        maxPallets: Infinity,
        displayName: 'Multiple Trucks / Over Limit 🚨',
        description: `Exceeds max single vehicle capacity (Weight/Length/Stack limit: ${maxStackingLayers}-high)`,
        color: 'text-rose-400 bg-rose-500/10 border-rose-500/30'
    };
}
