export const TRUCKS = [
    { 
        name: '2T', 
        maxWeight: 2000, 
        maxCbm: 13, 
        maxPallets: 4, 
        maxPalletsWide: 2, 
        length: 4.2, 
        width: 1.75, 
        height: 1.8,
        displayName: '2T Truck 🚛', 
        color: 'text-indigo-400 bg-indigo-500/10 border-indigo-500/30' 
    },
    { 
        name: '5T', 
        maxWeight: 5000, 
        maxCbm: 25, 
        maxPallets: 8, 
        maxPalletsWide: 4, 
        length: 6.0, 
        width: 2.1, 
        height: 2.0,
        displayName: '5T Truck 🚛', 
        color: 'text-sky-400 bg-sky-500/10 border-sky-500/30' 
    },
    { 
        name: '8T', 
        maxWeight: 8000, 
        maxCbm: 38, 
        maxPallets: 10, 
        maxPalletsWide: 5, 
        length: 7.4, 
        width: 2.35, 
        height: 2.2,
        displayName: '8T Truck 🚛', 
        color: 'text-blue-400 bg-blue-500/10 border-blue-500/30' 
    },
    { 
        name: '10T', 
        maxWeight: 10000, 
        maxCbm: 43, 
        maxPallets: 12, 
        maxPalletsWide: 6, 
        length: 8.0, 
        width: 2.35, 
        height: 2.3,
        displayName: '10T Truck 🚛', 
        color: 'text-cyan-400 bg-cyan-500/10 border-cyan-500/30' 
    },
    { 
        name: '15T', 
        maxWeight: 15000, 
        maxCbm: 53, 
        maxPallets: 16, 
        maxPalletsWide: 8, 
        length: 9.7, 
        width: 2.35, 
        height: 2.35,
        displayName: '15T Truck 🚛', 
        color: 'text-amber-400 bg-amber-500/10 border-amber-500/30' 
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
    { 
        name: '45HC', 
        maxWeight: 28000, 
        maxCbm: 76, 
        maxPallets: 24, 
        maxPalletsWide: 10, 
        length: 13.55, 
        width: 2.35, 
        height: 2.69,
        displayName: '45HC Container 🚢', 
        color: 'text-purple-400 bg-purple-500/10 border-purple-500/30' 
    },
];

export function recommendTruck(weightKg, cbm = 0, pallets = 0, isWideCargo = false, options = {}) {
    const numericWeight = weightKg !== null && weightKg !== undefined && !isNaN(weightKg) ? Number(weightKg) : 0;
    const numericCbm = cbm !== null && cbm !== undefined && !isNaN(cbm) ? Number(cbm) : 0;
    const numericPallets = pallets !== null && pallets !== undefined && !isNaN(pallets) ? Number(pallets) : 0;

    const maxStackingLayers = Number(options.maxStackingLayers) > 0 ? Number(options.maxStackingLayers) : 2;
    const palletLength = options.palletLength !== undefined && options.palletLength !== null ? Number(options.palletLength) : null;
    const palletWidth = options.palletWidth !== undefined && options.palletWidth !== null ? Number(options.palletWidth) : null;
    const palletHeight = options.palletHeight !== undefined && options.palletHeight !== null && Number(options.palletHeight) > 0 ? Number(options.palletHeight) : null;
    const clearance = 0.08; // 8cm safety gap between pallets / walls
    const verticalClearance = 0.08;
    const totalStackHeight = palletHeight ? (palletHeight * maxStackingLayers) + verticalClearance : null;

    // Explicit error if dimensions cannot be determined
    if (!palletLength || !palletWidth || palletLength <= 0 || palletWidth <= 0) {
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

    // Explicit error if pallet count is 0
    if (numericPallets <= 0) {
        return {
            name: 'No Pallets',
            maxWeight: 0,
            maxCbm: 0,
            maxPallets: 0,
            displayName: 'No Pallets ⚠️',
            description: 'Cannot recommend truck: Pallet count is 0',
            color: 'text-amber-400 bg-amber-500/10 border-amber-500/30'
        };
    }

    const loadingOrientation = options.loadingOrientation || 'width';
    const dim1 = Number(palletLength);
    const dim2 = Number(palletWidth);
    const longDim = Math.max(dim1, dim2);
    const shortDim = Math.min(dim1, dim2);

    const sortedTrucks = [...TRUCKS].sort((a, b) => a.maxCbm - b.maxCbm);

    const fittingTrucks = [];
    let tightTruck = null;
    let minOverflow = Infinity;
    let lastRejected = null;

    for (const truck of sortedTrucks) {
        // 1. Weight Check
        if (numericWeight > truck.maxWeight) {
            lastRejected = { ...truck, reason: 'Weight limit' };
            continue;
        }

        // 2. Height Collision Check (Z-Axis)
        if (totalStackHeight && totalStackHeight > truck.height) {
            lastRejected = { ...truck, reason: 'Height limit' };
            continue;
        }

        // 3. Orientation & Width Simulation (X-Axis)
        let acrossWidth = 0;
        let alongLength = 0;

        if (loadingOrientation === 'width') {
            if (longDim + clearance <= truck.width) {
                acrossWidth = longDim;
                alongLength = shortDim;
            } else if (shortDim + clearance <= truck.width) {
                acrossWidth = shortDim;
                alongLength = longDim;
            } else {
                lastRejected = { ...truck, reason: 'Width limit' };
                continue;
            }
        } else {
            // Length-wise loading
            if (shortDim + clearance <= truck.width) {
                acrossWidth = shortDim;
                alongLength = longDim;
            } else {
                lastRejected = { ...truck, reason: 'Width limit' };
                continue;
            }
        }

        // 4. Lateral 2-Wide vs 1-Wide Simulation (X-Axis)
        // Automatically computed from physical pallet width:
        const canLoadTwoWide = ((acrossWidth * 2) + clearance <= truck.width);
        const rowWidthFactor = canLoadTwoWide ? 2 : 1;

        // 5. Floor Spots & Longitudinal Length Simulation (Y-Axis)
        const floorSpotsNeeded = Math.ceil(numericPallets / maxStackingLayers);
        const rowsNeeded = Math.ceil(floorSpotsNeeded / rowWidthFactor);
        const totalLengthRequired = rowsNeeded * (alongLength + clearance);

        // 6. Door-to-Front Length Check (Y-Axis)
        if (totalLengthRequired > truck.length) {
            const overflow = totalLengthRequired - truck.length;
            if (overflow < minOverflow) {
                minOverflow = overflow;
                tightTruck = { ...truck, reason: 'Length limit', overflow: Number(overflow.toFixed(2)) };
            }
            lastRejected = { ...truck, reason: 'Length limit' };
            continue;
        }

        // 7. Volume / CBM Check (at end for palletized cargo)
        if (numericCbm > truck.maxCbm) {
            lastRejected = { ...truck, reason: 'CBM limit' };
            continue;
        }

        // 8. Vehicle Passes All Physics Checks!
        fittingTrucks.push({
            truck,
            acrossWidth,
            alongLength,
            canLoadTwoWide,
            rowsNeeded,
            totalLengthRequired,
            totalStackHeight
        });
    }

    const avgPalletWeight = numericPallets > 0 ? numericWeight / numericPallets : 0;
    const maxPalletWeight = options.maxPalletWeight && Number(options.maxPalletWeight) > 0 ? Number(options.maxPalletWeight) : avgPalletWeight;
    const bottomPalletLoad = maxStackingLayers > 1 ? Math.round(maxPalletWeight * (maxStackingLayers - 1)) : 0;

    const safetyAlert = (maxStackingLayers >= 3 && (bottomPalletLoad >= 1000 || avgPalletWeight >= 500)) ? {
        type: 'crush_risk',
        loadKg: bottomPalletLoad,
        label: `⚠️ ~${(bottomPalletLoad / 1000).toFixed(1)}T Crush Risk`,
        message: `Heavy Stack Warning: Bottom pallet carries ~${bottomPalletLoad.toLocaleString()}kg. High risk of leather grain crushing, indentation marks, and vehicle rollover. 2x stack recommended.`
    } : null;

    // 9. Multi-Tier Resolution
    if (fittingTrucks.length === 0) {
        const tightCandidate = tightTruck || lastRejected;
        return {
            name: 'Over Limit',
            maxWeight: Infinity,
            maxCbm: Infinity,
            maxPallets: Infinity,
            displayName: 'Multiple Trucks / Over Limit 🚨',
            description: `Exceeds max single vehicle capacity (Weight/Length/Stack limit: ${maxStackingLayers}-high)`,
            color: 'text-rose-400 bg-rose-500/10 border-rose-500/30',
            safetyAlert,
            alternatives: {
                tight: tightCandidate ? {
                    name: tightCandidate.name,
                    displayName: tightCandidate.displayName,
                    color: tightCandidate.color,
                    reason: tightCandidate.reason + (tightCandidate.overflow ? ` (+${tightCandidate.overflow}m)` : '')
                } : null,
                optimal: null,
                buffer: null
            }
        };
    }

    const bestMatch = fittingTrucks[0];
    const bufferMatch = fittingTrucks.length > 1 ? fittingTrucks[1] : null;
    const bestTruck = bestMatch.truck;

    const weightPct = Math.min(100, Math.round((numericWeight / bestTruck.maxWeight) * 100));
    const cbmPct = bestTruck.maxCbm > 0 && numericCbm > 0 ? Math.min(100, Math.round((numericCbm / bestTruck.maxCbm) * 100)) : 0;
    const lengthPct = Math.min(100, Math.round((bestMatch.totalLengthRequired / bestTruck.length) * 100));
    const heightPct = bestMatch.totalStackHeight ? Math.min(100, Math.round((bestMatch.totalStackHeight / bestTruck.height) * 100)) : null;

    const layoutLabel = bestMatch.canLoadTwoWide ? '2-wide' : '1-wide';
    const orientLabel = bestMatch.acrossWidth === longDim ? 'Width-wise ↔' : 'Length-wise ↕';

    return {
        ...bestTruck,
        tier: 'optimal',
        safetyAlert,
        utilization: {
            weightPct,
            cbmPct,
            lengthPct,
            heightPct
        },
        alternatives: {
            tight: tightTruck ? {
                name: tightTruck.name,
                displayName: tightTruck.displayName,
                color: tightTruck.color,
                reason: tightTruck.reason + (tightTruck.overflow ? ` (+${tightTruck.overflow}m)` : '')
            } : null,
            optimal: {
                name: bestTruck.name,
                displayName: bestTruck.displayName,
                color: bestTruck.color
            },
            buffer: bufferMatch ? {
                name: bufferMatch.truck.name,
                displayName: bufferMatch.truck.displayName,
                color: bufferMatch.truck.color
            } : null
        },
        description: `${bestTruck.displayName} — Max ${bestTruck.maxWeight.toLocaleString()}kg, ${bestTruck.maxCbm} CBM | Fits ${numericPallets} Pallet(s) in ${bestMatch.rowsNeeded} row(s) [${layoutLabel}, ${orientLabel}, Max Stack: ${maxStackingLayers}x]`
    };
}
