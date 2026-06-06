import logging
from core.utils.cache import mapping_cache
from core.database.session import SessionLocal
from core.database.models.global_map import (
    GlobalMapSheet,
    GlobalMapHeaderTextMapping,
    GlobalMapColumn,
    GlobalMapColumnKeyword,
    GlobalMapFooterLabelKeyword,
    GlobalMapFallbackStrategy,
)

logger = logging.getLogger(__name__)

def get_global_mapping_config(db=None) -> dict:
    cached = mapping_cache.get()
    if cached is not None:
        return cached

    close_session = False
    if db is None:
        db = SessionLocal()
        close_session = True
        
    try:
        # Load from relational tables and reconstruct dictionary
        config = {}
        
        # 1 & 2. Load sheets and classifications from unified table
        sheets = db.query(GlobalMapSheet).all()
        agg_set = set()
        proc_set = set()
        config["sheet_name_mappings"] = {"mappings": {}}
        
        for s in sheets:
            sheet_lower = s.sheet_name.lower().strip()
            if s.processing_type == "aggregation":
                agg_set.add(sheet_lower)
            elif s.processing_type == "processed_tables":
                proc_set.add(sheet_lower)
                        
        config["aggregation_sheets"] = sorted(list(agg_set))
        config["processed_tables_sheets"] = sorted(list(proc_set))
        
        # 3. Load header text mappings
        header_maps = db.query(GlobalMapHeaderTextMapping).all()
        config["header_text_mappings"] = {
            "mappings": {hm.raw_text: hm.canonical_col_id for hm in header_maps}
        }
        
        # 4. Load shipping header map (columns + keywords)
        columns = db.query(GlobalMapColumn).all()
        config["shipping_header_map"] = {}
        for col in columns:
            config["shipping_header_map"][col.col_id] = {
                "keywords": [kw.keyword for kw in col.keywords],
                "format": col.excel_format
            }
            
        # 5. Load footer label mappings
        footers = db.query(GlobalMapFooterLabelKeyword).all()
        config["footer_label_mappings"] = {
            "keywords": [f.keyword for f in footers]
        }
        
        # 6. Load fallback strategies
        strategies = db.query(GlobalMapFallbackStrategy).all()
        config["fallback_strategies"] = {}
        for s in strategies:
            val = s.strategy_value
            if val.lower() == "true": val = True
            elif val.lower() == "false": val = False
            elif val.replace('.', '', 1).isdigit(): val = float(val) if '.' in val else int(val)
            config["fallback_strategies"][s.strategy_key] = val
            
        mapping_cache.set(config)
        return config

    except Exception as e:
        logger.exception("Database query failed in get_global_mapping_config")
        raise e
    finally:
        if close_session:
            db.close()

def save_global_mapping_config(data: dict, db=None) -> None:
    close_session = False
    if db is None:
        db = SessionLocal()
        close_session = True
        
    try:
        # 1. Clear existing relational mapping data to overwrite completely
        # Delete children first to satisfy foreign key constraints
        db.query(GlobalMapSheet).delete()
        db.query(GlobalMapHeaderTextMapping).delete()
        db.query(GlobalMapColumnKeyword).delete()
        db.query(GlobalMapColumn).delete()
        db.query(GlobalMapFooterLabelKeyword).delete()
        db.query(GlobalMapFallbackStrategy).delete()
        
        # 2 & 3. Insert sheet classifications into unified table
        agg_sheets = data.get("aggregation_sheets", [])
        proc_sheets = data.get("processed_tables_sheets", [])
        
        seen_sheets = set()
        
        for sheet in agg_sheets:
            sheet_clean = sheet.strip()
            if not sheet_clean:
                continue
            sheet_lower = sheet_clean.lower()
            if sheet_lower not in seen_sheets:
                seen_sheets.add(sheet_lower)
                db.add(GlobalMapSheet(
                    sheet_name=sheet_clean,
                    processing_type="aggregation"
                ))
                
        for sheet in proc_sheets:
            sheet_clean = sheet.strip()
            if not sheet_clean:
                continue
            sheet_lower = sheet_clean.lower()
            if sheet_lower not in seen_sheets:
                seen_sheets.add(sheet_lower)
                db.add(GlobalMapSheet(
                    sheet_name=sheet_clean,
                    processing_type="processed_tables"
                ))
            
        # 4. Insert shipping columns and column keywords
        shipping_map = data.get("shipping_header_map", {})
        for col_id, props in shipping_map.items():
            if not isinstance(props, dict):
                continue
            fmt = props.get("format", "@")
            col = GlobalMapColumn(col_id=col_id, excel_format=fmt)
            db.add(col)
            
            # Keywords
            kws = props.get("keywords", [])
            seen_kws = set()
            for kw in kws:
                kw_clean = kw.strip()
                if not kw_clean:
                    continue
                kw_lower = kw_clean.lower()
                if kw_lower not in seen_kws:
                    seen_kws.add(kw_lower)
                    db.add(GlobalMapColumnKeyword(col_id=col_id, keyword=kw_clean))
                
        # 5. Insert header text mappings (overrides)
        # We only write overrides pointing to column IDs that exist in the columns list
        header_mappings = data.get("header_text_mappings", {}).get("mappings", {})
        for raw, canonical in header_mappings.items():
            if canonical in shipping_map:
                db.add(GlobalMapHeaderTextMapping(raw_text=raw, canonical_col_id=canonical))
                
        # 6. Insert footer label mappings
        footer_kws = data.get("footer_label_mappings", {}).get("keywords", [])
        seen_footers = set()
        for kw in footer_kws:
            kw_clean = kw.strip()
            if not kw_clean:
                continue
            kw_lower = kw_clean.lower()
            if kw_lower not in seen_footers:
                seen_footers.add(kw_lower)
                db.add(GlobalMapFooterLabelKeyword(keyword=kw_clean))
            
        # 7. Insert fallback strategies
        strategies = data.get("fallback_strategies", {})
        for key, val in strategies.items():
            db.add(GlobalMapFallbackStrategy(strategy_key=key, strategy_value=str(val)))
            
        db.commit()
        mapping_cache.invalidate()
    except Exception as e:
        db.rollback()
        raise e
    finally:
        if close_session:
            db.close()
