import logging
from typing import Dict, Any, List
from sqlalchemy.orm import Session
from core.database.db_manager import (
    get_global_mapping_config, save_global_mapping_config, GlobalMapSheet,
    GlobalMapFallbackStrategy, GlobalMapHeaderTextMapping
)

logger = logging.getLogger(__name__)

from core.utils.cache import mapping_cache

class MappingService:
    """
    Service layer handles business logic for global mapping operations,
    formatting mapping records, database interactions, and cache/state reloading.
    """
    
    def __init__(self, db: Session):
        self.db = db

    def get_mappings(self, mapping_type: str = "header_text_mappings") -> dict:
        """
        Retrieves global mapping definitions from the database and flattens/formats
        them according to what the frontend expects.
        """
        data = get_global_mapping_config(self.db)

        if mapping_type == "shipping_header_map":
            # Return flat mapping of {keyword: col_id} for UI
            from core.database.db_manager import GlobalMapColumnKeyword
            keywords = self.db.query(GlobalMapColumnKeyword).all()
            return {kw.keyword: kw.col_id for kw in keywords}
            
        elif mapping_type == "footer_label_mappings":
            keywords = data.get("footer_label_mappings", {}).get("keywords", [])
            return {kw: "Footer Keyword" for kw in keywords}
            
        elif mapping_type == "sheet_classifications":
            flat = {}
            for s in data.get("aggregation_sheets", []):
                flat[s] = "aggregation"
            for s in data.get("processed_tables_sheets", []):
                flat[s] = "processed_tables"
            return flat
            
        elif mapping_type == "sheet_mappings":
            sheets = self.db.query(GlobalMapSheet).all()
            res_dict = {}
            for s in sheets:
                res_dict[s.sheet_name] = s.processing_type
            return res_dict

        elif mapping_type == "fallback_strategies":
            strategies = self.db.query(GlobalMapFallbackStrategy).all()
            return {s.strategy_key: s.strategy_value for s in strategies}

        elif mapping_type == "header_text_mappings":
            overrides = self.db.query(GlobalMapHeaderTextMapping).all()
            strategies = self.db.query(GlobalMapFallbackStrategy).all()
            fallback_dict = {s.strategy_key: s.strategy_value for s in strategies}
            res = {}
            for override in overrides:
                fb = fallback_dict.get(override.raw_text, fallback_dict.get(override.canonical_col_id, ""))
                res[override.raw_text] = {
                    "col_id": override.canonical_col_id,
                    "fallback": fb if fb is not None else ""
                }
            return res

        return data.get(mapping_type, {}).get("mappings", {})

    def update_mappings(self, mapping_type: str, mappings: Dict[str, Any]) -> None:
        """
        Parses UI-provided configurations, structures them back into database models,
        commits updates to relational tables using targeted delta updates, and triggers
        an in-memory hot reload.
        """
        from core.database.db_manager import (
            GlobalMapColumn, GlobalMapColumnKeyword, GlobalMapHeaderTextMapping,
            GlobalMapSheet, GlobalMapFooterLabelKeyword, GlobalMapFallbackStrategy
        )

        try:
            if mapping_type == "shipping_header_map":
                # Extract target (col_id, keyword) pairs supporting flat {kw: col_id} and nested {col_id: {'keywords': [...]}}
                target_pairs = set()
                for k, v in mappings.items():
                    if isinstance(v, dict):
                        col_id = str(k).strip()
                        kws = v.get("keywords", [])
                        if isinstance(kws, (list, tuple, set)):
                            for kw in kws:
                                kw_str = str(kw).strip() if kw is not None else ""
                                if kw_str and col_id:
                                    target_pairs.add((col_id, kw_str))
                        elif isinstance(kws, str):
                            kw_str = kws.strip()
                            if kw_str and col_id:
                                target_pairs.add((col_id, kw_str))
                    elif isinstance(v, (list, tuple, set)):
                        col_id = str(k).strip()
                        for kw in v:
                            kw_str = str(kw).strip() if kw is not None else ""
                            if kw_str and col_id:
                                target_pairs.add((col_id, kw_str))
                    elif isinstance(v, str):
                        col_id = v.strip()
                        kw_str = str(k).strip()
                        if kw_str and col_id:
                            target_pairs.add((col_id, kw_str))

                # Ensure parent columns exist for all unique col_ids in target pairs
                unique_cols = {col_id for col_id, _ in target_pairs}
                for col_id in unique_cols:
                    col_exists = self.db.query(GlobalMapColumn).filter_by(col_id=col_id).first()
                    if not col_exists:
                        self.db.add(GlobalMapColumn(col_id=col_id, excel_format="@"))
                self.db.flush()

                # Query existing GlobalMapColumnKeyword and build map {(kw.col_id, kw.keyword): kw}
                existing_kws = self.db.query(GlobalMapColumnKeyword).all()
                existing_map = {(kw.col_id, kw.keyword): kw for kw in existing_kws}

                # Delete any existing row whose (col_id, keyword) is not in target pairs
                for pair, kw_obj in existing_map.items():
                    if pair not in target_pairs:
                        self.db.delete(kw_obj)
                self.db.flush()

                # Add GlobalMapColumnKeyword(col_id=col_id, keyword=kw_text) for any pair in target pairs not in existing
                for col_id, kw_text in target_pairs:
                    if (col_id, kw_text) not in existing_map:
                        self.db.add(GlobalMapColumnKeyword(col_id=col_id, keyword=kw_text))
                
            elif mapping_type == "footer_label_mappings":
                existing_footers = self.db.query(GlobalMapFooterLabelKeyword).all()
                existing_map = {f.keyword: f for f in existing_footers}
                existing_set = set(existing_map.keys())
                new_set = {str(k).strip() for k in mappings.keys() if str(k).strip() and k != "keywords"}
                
                # Deletes
                for kw in (existing_set - new_set):
                    self.db.delete(existing_map[kw])
                
                # Adds
                for kw in (new_set - existing_set):
                    self.db.add(GlobalMapFooterLabelKeyword(keyword=kw))
                
            elif mapping_type in ("sheet_classifications", "sheet_mappings"):
                clean_mappings = {str(k).strip(): str(v).strip() for k, v in mappings.items() if str(k).strip()}
                existing_sheets = self.db.query(GlobalMapSheet).all()
                existing_map = {s.sheet_name: s for s in existing_sheets}
                
                # Deletes
                for sheet_name, sheet_obj in existing_map.items():
                    if sheet_name not in clean_mappings:
                        self.db.delete(sheet_obj)
                
                # Adds and Updates
                for sheet_name, proc_type in clean_mappings.items():
                    if sheet_name not in existing_map:
                        self.db.add(GlobalMapSheet(sheet_name=sheet_name, processing_type=proc_type))
                    elif existing_map[sheet_name].processing_type != proc_type:
                        existing_map[sheet_name].processing_type = proc_type
                        
            elif mapping_type == "header_text_mappings":
                clean_mappings = {}
                fallback_updates = {}
                for k, v in mappings.items():
                    raw_text = str(k).strip() if k is not None else ""
                    if not raw_text:
                        continue
                    if isinstance(v, dict):
                        col_id = str(v.get("col_id", "")).strip()
                        if not col_id or col_id.startswith("col_unknown") or col_id == "col_unknown":
                            continue
                        if not col_id.startswith("col_"):
                            col_id = f"col_{col_id}"
                        col_id = col_id.replace(" ", "_").lower()
                        clean_mappings[raw_text] = col_id
                        if "fallback" in v:
                            fb = v.get("fallback")
                            fallback_updates[raw_text] = str(fb).strip() if fb is not None else ""
                    else:
                        col_id = str(v).strip() if v is not None else ""
                        if not col_id or col_id.startswith("col_unknown") or col_id == "col_unknown":
                            continue
                        if not col_id.startswith("col_"):
                            col_id = f"col_{col_id}"
                        col_id = col_id.replace(" ", "_").lower()
                        clean_mappings[raw_text] = col_id

                existing_overrides = self.db.query(GlobalMapHeaderTextMapping).all()
                existing_map = {o.raw_text: o for o in existing_overrides}
                
                # Ensure parent columns exist for any custom/new col_id
                unique_cols = {c for c in clean_mappings.values() if c and not c.startswith("col_unknown")}
                for col_id in unique_cols:
                    col_exists = self.db.query(GlobalMapColumn).filter_by(col_id=col_id).first()
                    if not col_exists:
                        self.db.add(GlobalMapColumn(col_id=col_id, excel_format="@"))
                self.db.flush()
                
                # Deletes
                for raw_text, override_obj in existing_map.items():
                    if raw_text not in clean_mappings:
                        self.db.delete(override_obj)
                
                # Adds and Updates
                for raw_text, col_id in clean_mappings.items():
                    if raw_text not in existing_map:
                        self.db.add(GlobalMapHeaderTextMapping(raw_text=raw_text, canonical_col_id=col_id))
                    elif existing_map[raw_text].canonical_col_id != col_id:
                        existing_map[raw_text].canonical_col_id = col_id
                self.db.flush()

                # Manage Fallback Strategies for raw_text keys
                existing_strategies = self.db.query(GlobalMapFallbackStrategy).all()
                existing_strat_map = {s.strategy_key: s for s in existing_strategies}

                # Delete fallback strategy if mapped header was deleted
                for raw_text in existing_map.keys():
                    if raw_text not in clean_mappings:
                        strat_key = raw_text[:128]
                        if strat_key in existing_strat_map:
                            self.db.delete(existing_strat_map[strat_key])

                # Update or delete fallback strategy for specified fallbacks
                for raw_text, fb_val in fallback_updates.items():
                    strat_key = raw_text[:128]
                    if fb_val:
                        if strat_key not in existing_strat_map:
                            self.db.add(GlobalMapFallbackStrategy(strategy_key=strat_key, strategy_value=fb_val))
                        elif existing_strat_map[strat_key].strategy_value != fb_val:
                            existing_strat_map[strat_key].strategy_value = fb_val
                    else:
                        if strat_key in existing_strat_map:
                            self.db.delete(existing_strat_map[strat_key])
                        
            elif mapping_type == "fallback_strategies":
                clean_mappings = {str(k).strip(): str(v).strip() for k, v in mappings.items() if str(k).strip()}
                existing_strategies = self.db.query(GlobalMapFallbackStrategy).all()
                existing_map = {s.strategy_key: s for s in existing_strategies}
                
                # Deletes
                for strat_key, strat_obj in existing_map.items():
                    if strat_key not in clean_mappings:
                        self.db.delete(strat_obj)
                
                # Adds and Updates
                for strat_key, strat_val in clean_mappings.items():
                    if strat_key not in existing_map:
                        self.db.add(GlobalMapFallbackStrategy(strategy_key=strat_key, strategy_value=strat_val))
                    elif existing_map[strat_key].strategy_value != strat_val:
                        existing_map[strat_key].strategy_value = strat_val
            
            self.db.commit()
            mapping_cache.invalidate()
        except Exception as e:
            self.db.rollback()
            logger.error(f"Failed to update mappings in database: {e}")
            raise e

        # Reload mappings in-memory immediately so changes take effect
        self.reload_dynamic_state()

    def reload_dynamic_state(self, data: dict = None) -> None:
        """Reload parsed alias indexing and schema headers dynamically without server restart."""
        try:
            if data is None:
                data = get_global_mapping_config(self.db)
            
            from core.data_parser.config import load_and_update_mappings
            load_and_update_mappings()
            
            from core.data_parser.sheet_parser import _build_alias_lookup
            import core.data_parser.sheet_parser as _sp_module
            _sp_module._ALIAS_REVERSE_LOOKUP = _build_alias_lookup()
            
            from core.blueprint_generator.schema import BlueprintSchema
            BlueprintSchema.load_dynamic_columns(data)
        except Exception as e:
            logger.warning(f"Could not automatically reload mappings: {e}")
