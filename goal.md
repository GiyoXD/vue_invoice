# 🎯 Giyo Invoice — Improvement Goals

> Generated from code review on 2026-04-24. Prioritized by impact.

---

## 🔴 High Priority (Fix Soon)

### 1. Replace bare `except:` with `except Exception:` + add logging
- **Files**: `api/routers/templates.py` (lines 65, 76, 92, 110, 185)
- **Why**: Bare `except:` catches `KeyboardInterrupt` and `SystemExit`, making the app un-killable during errors. Silent swallowing hid real errors.
- **Fix**: Narrowed to `except Exception:` + added `logger.exception()` at all 5 sites
- **Status**: [x] Done (2026-04-24)

### 2. Remove orphaned code in `data_processor.py`
- **File**: `core/data_parser/data_processor.py` line 667
- **What**: Floating `logging.info(...)` statement at module level (outside any function). Left behind when `validate_weight_integrity` was extracted to `validation.py`
- **Effort**: 5 min
- **Status**: [x] Done (2026-04-24)

### 3. Deduplicate CSS in `style.css`
- **File**: `frontend/css/style.css`
- **What**: `.nav-bar`, `.nav-btn`, `.data-table`, `.btn-small`, `.history-list`, `.history-item` are duplicated (~200 lines of copy-paste)
- **Effort**: 30 min
- **Status**: [x] Done (2026-04-24)

### 4. Remove legacy flag parsing in orchestrator
- **File**: `core/orchestrator.py` lines 63-86
- **What**: Orchestrator converts string flags (`"--DAF"`, `"--custom"`) to bools. Now that we use API (not CLI), pass typed arguments directly from the router.
- **Effort**: 1 hr
- **Status**: [x] Done (2026-04-24)

---

## 🟡 Medium Priority (Improve Quality)

### 5. Move imports to top of file
- **Files**: `core/orchestrator.py` (traceback), `core/data_parser/main.py` (math), `core/data_parser/data_processor.py` (PipelineMonitor at line 289)
- **Rule**: Always import at file top. Only import inside a function if there's a circular dependency.
- **Effort**: 30 min
- **Status**: [ ] Not started

### 6. Break up god files
- `core/data_parser/main.py` (892 lines) → Extract DAF compounding, footer calc, pallet formatting into separate modules
- `core/data_parser/data_processor.py` (1267 lines) → Split aggregation functions, distribution logic, and pricing injection into focused modules
- **Effort**: 3 hrs
- **Status**: [ ] Not started

### 7. Standardize error handling across routers
- Currently: `generate.py` returns full tracebacks, `templates.py` returns one-liners, `upload.py` mixes both
- Goal: Consistent error response shape `{ error: str, step?: str, traceback?: str, details?: [] }`
- **Effort**: 2 hrs
- **Status**: [ ] Not started

### 8. Add return type hints to router helpers
- Focus on: `templates.py` helper functions, `orchestrator.py` methods
- **Effort**: 1 hr
- **Status**: [ ] Not started

---

## 🟢 Low Priority (Polish)

### 9. Make DAF parameters config-driven
- Move `DAF_CHUNK_SIZE`, `>7 PO threshold` from module constants to per-client config
- **Effort**: 1 hr
- **Status**: [ ] Not started

### 10. Consider Vite + SFC for frontend
- Current: Vue via CDN with inline template strings (36KB JS files with HTML as strings)
- Benefit: Syntax highlighting, hot reload, component splitting
- **Effort**: 8 hrs (migration)
- **Status**: [ ] Not started

---

## 🧪 Testing — Getting Started Guide

> "i dont know how to" — Here's the practical path:

### Step 1: Test what has burned you before
Look at your `note.md` bugs. Each bug you fixed = one test you should write.

### Step 2: Start with pure functions (no file I/O)
These are the easiest to test because they take data in and return data out:

```python
# tests/test_data_processor.py
import pytest
from decimal import Decimal
from core.data_parser.data_processor import (
    process_cbm_column,
    distribute_values,
    normalize_pallet_count,
)

def test_cbm_calculation_from_formula():
    """CBM string '1*2*3' should calculate to 6.0000"""
    rows = [{"col_cbm": "1*2*3"}]
    result = process_cbm_column(rows)
    assert result[0]["col_cbm"] == Decimal("6.0000")

def test_cbm_already_numeric():
    """Numeric CBM values should pass through unchanged"""
    rows = [{"col_cbm": 12.5}]
    result = process_cbm_column(rows)
    assert result[0]["col_cbm"] == Decimal("12.5000")

def test_pallet_normalization_xy_format():
    """'1-39' format should normalize to 1 (new pallet) or 0 (continuation)"""
    rows = [
        {"col_pallet_count": "1-39"},
        {"col_pallet_count": "1-39"},  # same pallet → 0
        {"col_pallet_count": "2-39"},  # new pallet → 1
    ]
    result = normalize_pallet_count(rows)
    assert result[0]["col_pallet_count"] == 1
    assert result[1]["col_pallet_count"] == 0
    assert result[2]["col_pallet_count"] == 1

def test_distribute_values_proportional():
    """100 distributed over [50, 30, 20] sqft should give [50, 30, 20] amount"""
    rows = [
        {"col_qty_sf": Decimal("50"), "col_amount": Decimal("100")},
        {"col_qty_sf": Decimal("30"), "col_amount": None},
        {"col_qty_sf": Decimal("20"), "col_amount": None},
    ]
    result = distribute_values(rows, ["col_amount"], "col_qty_sf")
    assert result[0]["col_amount"] == Decimal("50.0000")
    assert result[1]["col_amount"] == Decimal("30.0000")
    assert result[2]["col_amount"] == Decimal("20.0000")
```

### Step 3: Run tests
```bash
# Install pytest (one time)
pip install pytest

# Run all tests
pytest tests/ -v

# Run a specific test
pytest tests/test_data_processor.py::test_cbm_calculation_from_formula -v
```

### Step 4: Add tests for each new bug fix
Every time you fix a bug, write a test that would have caught it. This is the fastest way to build real coverage.

---

## 📝 Notes from original `note.md`
- Track invoice gen exception bugs
- Grand total bug
- Refactor file organization
- Add safety measure for CBM and net weight placement
