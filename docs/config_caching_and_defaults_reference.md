# Reference Guide: Config Store Caching & Global Defaults

> [!IMPORTANT]
> **Why changes to `master_config.json` might not show up immediately in the Web UI**
> 
> Python caches master configuration defaults in RAM using `@lru_cache(maxsize=1)` inside `ConfigStore._get_master_defaults()`.

---

## 1. How Master Defaults Are Cached

- **Location**: [store.py](file:///c:/Users/JPZ031127/Desktop/project/vue_invoice_project/core/invoice_generator/config/store.py#L72-L86)
- **Function**: `ConfigStore._get_master_defaults()`
- **Mechanism**: Decorator `@lru_cache(maxsize=1)` reads [master_config.json](file:///c:/Users/JPZ031127/Desktop/project/vue_invoice_project/database/blueprints/mapper/master_config.json) once and holds the dictionary in process RAM.

### What happens when you edit `master_config.json` on disk:
While the server process (`.\start_dev.ps1`) is running, Uvicorn will **NOT** automatically bust in-memory Python `@lru_cache` functions unless the Python process is restarted.

---

## 2. How to Clear / Reload the Cache

If you edit `master_config.json` or update global static defaults (`static_payload`), choose one of the following methods to apply changes:

### Option A: Restart Dev Server (Recommended)
In your terminal, press `Ctrl + C` to stop the dev server, then re-run:
```powershell
.\start_dev.ps1
```

### Option B: Programmatically Clear Cache
In Python or test scripts, invoke:
```python
from core.invoice_generator.config.store import ConfigStore

# Programmatically purge the master defaults cache
ConfigStore._get_master_defaults.cache_clear()
```

---

## 3. Database Blueprint Legacy Records

- Blueprints saved in the SQLite database (`blueprints` table) store their own copy of `config_json`.
- When generating invoices from the UI, `ConfigStore` automatically falls back to `master_config.json` defaults for any missing `static_payload` or fallback rules.
- If you change template structures or static payloads extensively, re-clicking **"Generate Blueprint"** on the UI Template Management page updates the database row to the new format.
