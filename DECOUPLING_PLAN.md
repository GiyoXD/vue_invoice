# Decoupling Plan: `api/routers/blueprint.py`

This document details the coupling issues inside `api/routers/blueprint.py` and provides a concrete implementation plan for refactoring it in the future.

---

## 1. Coupling Issues Identified

1. **Inline Database Management**: DB sessions are manually initialized and closed using inline `try/finally` blocks rather than using FastAPI's standard dependency injection.
2. **Transformations in Controllers**: Request transformations (e.g. converting a flat dictionary into structured formats for `shipping_header_map`) are done directly inside the HTTP route handlers.
3. **Module Rebuild Side-Effects**: Dynamically re-initializing global configurations and rebuilding index caches is done directly inside the route controllers, causing side effects that are hard to isolate and test.

---

## 2. Proposed Architecture

```
HTTP Request ──> Controller (Router) ──> MappingService ──> Repository ──> Database
                                                │
                                                └──> Dynamic State Reload Helper
```

- **HTTP Router (`api/routers/blueprint.py`)**: Responsible only for handling request formats, validating routing inputs, and returning HTTP responses.
- **Mapping Service (`core/services/mapping_service.py`)**: Handles mapping transformations and state reloads.
- **Repository (`core/database/repositories.py`)**: Handles DB queries and updates.

---

## 3. Implementation Steps for Future Refactoring

### Step 1: Support DB Session Dependency Injection
Create or update FastAPI database session setup so route functions can inject `db` dynamically:
```python
# In api/dependencies.py or core/database/db_manager.py
from core.database.db_manager import SessionLocal

def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()
```

### Step 2: Implement the `MappingService`
Create a new service at `core/services/mapping_service.py` to encapsulate business logic:
```python
# core/services/mapping_service.py
from sqlalchemy.orm import Session
from core.database.repositories import BlueprintRepository

class MappingService:
    def __init__(self, db: Session):
        self.db = db
        self.repo = BlueprintRepository(db)

    def get_mappings(self, mapping_type: str) -> dict:
        # 1. Fetch from config via repository
        # 2. Format / flatten if necessary
        ...

    def update_mappings(self, mapping_type: str, mappings: dict) -> None:
        # 1. Unflatten or structure raw mappings
        # 2. Save config via repository
        # 3. Trigger dynamic cache reloads
        ...
```

### Step 3: Extract Dynamic Reload Logic
Extract dynamic configuration reloads into a reusable utility function:
```python
def dynamic_reload_mappings(data: dict) -> None:
    from core.data_parser.config import load_and_update_mappings
    from core.data_parser.sheet_parser import _build_alias_lookup
    import core.data_parser.sheet_parser as _sp_module
    from core.blueprint_generator.schema import BlueprintSchema

    load_and_update_mappings()
    _sp_module._ALIAS_REVERSE_LOOKUP = _build_alias_lookup()
    BlueprintSchema.load_dynamic_columns(data)
```

### Step 4: Simplify Route Handlers
Update the route endpoints in `api/routers/blueprint.py` to use dependencies:
```python
from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.orm import Session
from core.database.db_manager import get_db
from core.services.mapping_service import MappingService

@router.post("/mappings")
async def update_mappings(
    request: MappingsUpdateRequest, 
    db: Session = Depends(get_db)
):
    try:
        service = MappingService(db)
        service.update_mappings(request.mapping_type, request.mappings)
        return {"status": "success"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
```
