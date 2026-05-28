export const TRUCKS = [
    { name: '3T', maxWeight: 3000, maxCbm: 12, maxPallets: 4, maxPalletsWide: 2, displayName: '3T Truck 🚛', color: 'text-indigo-400 bg-indigo-500/10 border-indigo-500/30' },
    { name: '5T', maxWeight: 5000, maxCbm: 20, maxPallets: 6, maxPalletsWide: 3, displayName: '5T Truck 🚛', color: 'text-sky-400 bg-sky-500/10 border-sky-500/30' },
    { name: '8T', maxWeight: 8000, maxCbm: 32, maxPallets: 10, maxPalletsWide: 5, displayName: '8T Truck 🚛', color: 'text-blue-400 bg-blue-500/10 border-blue-500/30' },
    { name: '20GP', maxWeight: 20000, maxCbm: 28, maxPallets: 10, maxPalletsWide: 4, displayName: '20GP Container 🚢', color: 'text-emerald-400 bg-emerald-500/10 border-emerald-500/30' },
    { name: '40HC', maxWeight: 25000, maxCbm: 68, maxPallets: 21, maxPalletsWide: 9, displayName: '40HC Container 🚢', color: 'text-teal-400 bg-teal-500/10 border-teal-500/30' },
];

export function recommendTruck(weightKg, cbm = 0, pallets = 0, isWideCargo = false) {
    const numericWeight = weightKg !== null && weightKg !== undefined && !isNaN(weightKg) ? Number(weightKg) : 0;
    const numericCbm = cbm !== null && cbm !== undefined && !isNaN(cbm) ? Number(cbm) : 0;
    const numericPallets = pallets !== null && pallets !== undefined && !isNaN(pallets) ? Number(pallets) : 0;

    for (const truck of TRUCKS) {
        const allowedPallets = isWideCargo ? truck.maxPalletsWide : truck.maxPallets;
        if (
            numericWeight <= truck.maxWeight &&
            numericCbm <= truck.maxCbm &&
            numericPallets <= allowedPallets
        ) {
            return {
                ...truck,
                description: `${truck.displayName} (Max ${truck.maxWeight.toLocaleString()}kg, ${truck.maxCbm} CBM, ${allowedPallets} Pallets)`
            };
        }
    }

    // If limit exceeded
    return {
        name: 'Over Limit',
        maxWeight: Infinity,
        maxCbm: Infinity,
        maxPallets: Infinity,
        displayName: 'Multiple Trucks / Over Limit 🚨',
        description: `Exceeds max single vehicle capacity (25T / 68 CBM / ${isWideCargo ? '9 Wide' : '21 Standard'} Pallets)`,
        color: 'text-rose-400 bg-rose-500/10 border-rose-500/30'
    };
}
