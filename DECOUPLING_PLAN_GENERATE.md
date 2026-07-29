# Decoupling Plan: `api/routers/generate.py`

This document details the coupling issues inside the invoice generation endpoint (`/api/generate`) and provides a concrete implementation plan for refactoring it into a clean, modular structure.

---

## 1. Coupling Issues Identified

1. **Inline Pricing Modifications**: Handlers for applying manual price adjustments and injecting net-weight pricing are implemented inline inside the FastAPI route function.
2. **In-Memory Re-aggregation**: Standard, custom, DAF compounding, and manifest aggregations are recalculated inside the router after data adjustments are made.
3. **File System Operations**: Reading raw JSON files, validating integrity, and writing files atomically using tempfiles is done directly inside the route logic.
4. **Variant Resolution & packaging**: The routing code is coupled with resolving variant templates (Standard, Custom, DAF, KH, VN), calling the orchestrator, and encoding binary files to base64.

---

## 2. Proposed Architecture

Instead of a single "God function", the generation routing will be split into a dedicated sub-package `api/routers/generate/` with clear boundaries:

```
api/routers/generate/
│
├── __init__.py           # Exposes the router to api/main.py
│
├── router.py             # FastAPI routing only (endpoints, request/response models)
│
├── pricing.py            # Handles price adjustment and net weight price injection logic
│
├── aggregator.py         # Handles running standard, custom, DAF, and manifest aggregations
│
└── generator.py          # Handles template variant resolution, execution, and ZIP packaging
```

---

## 3. Implementation Steps for Future Refactoring

### Step 1: Implement the Pricing Layer (`pricing.py`)
Extract the raw data modification logic. This takes the raw JSON dictionary and applies pricing edits:
```python
# api/routers/generate/pricing.py
from typing import List, Dict, Any, Optional

def apply_invoice_pricing_overrides(
    full_data: Dict[str, Any],
    price_adjustments: Optional[List[List[Any]]] = None,
    global_unit_price: Optional[float] = None
) -> Dict[str, Any]:
    # 1. Update invoice_info metadata
    # 2. Apply price adjustments via apply_aggregation_adjustment
    # 3. Inject net-weight pricing via inject_net_weight_pricing if global_unit_price is provided
    # 4. Return the modified dict
```

### Step 2: Implement the Aggregator Layer (`aggregator.py`)
Extract the re-aggregation logic that recalculates all tables and footers:
```python
# api/routers/generate/aggregator.py
from typing import Dict, Any
from core.data_parser.data_processor import (
    aggregate_standard_by_po_item_price,
    aggregate_custom_by_po_item,
    format_aggregation_as_list,
    aggregate_per_po_with_pallets,
    calculate_footer_totals,
    perform_DAF_compounding
)

def recalculate_invoice_aggregations(full_data: Dict[str, Any]) -> Dict[str, Any]:
    # 1. Extract multi_table data and flatten rows
    # 2. Recalculate footers and update grand total footer data
    # 3. Run standard and custom aggregations
    # 4. Run DAF compounding (now accepts merged rows directly)
    # 5. Run manifest/pallet aggregation per PO
    # 6. Update and return full_data
```

### Step 3: Implement the Generator/Packager Layer (`generator.py`)
Extract the variant execution and packaging logic:
```python
# api/routers/generate/generator.py
from pathlib import Path
from typing import List, Dict, Any

def execute_generation_variants(
    json_path: Path,
    targets: List[str],
    options: dict
) -> List[Dict[str, Any]]:
    # 1. Resolve variant template suffixes
    # 2. Invoke core.orchestrator.generate_invoice
    # 3. Zip files and package them into base64 payload response format
```

### Step 4: Simplify the Router (`router.py`)
Clean up the router endpoint to act strictly as an orchestrator/coordinator:
```python
# api/routers/generate/router.py
from fastapi import APIRouter
from .pricing import apply_invoice_pricing_overrides
from .aggregator import recalculate_invoice_aggregations
from .generator import execute_generation_variants

@router.post("/generate")
def generate_invoice(request: GenerateRequest):
    # 1. Load JSON from disk
    # 2. Apply pricing updates
    # 3. Recalculate aggregates
    # 4. Save JSON atomically to disk
    # 5. Run variant generation and package response
```
