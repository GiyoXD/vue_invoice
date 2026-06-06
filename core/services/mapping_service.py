import logging
from typing import Dict, Any, List
from sqlalchemy.orm import Session
from core.database.db_manager import (
    get_global_mapping_config, save_global_mapping_config, GlobalMapSheet
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

        return data.get(mapping_type, {}).get("mappings", {})

    def update_mappings(self, mapping_type: str, mappings: Dict[str, Any]) -> None:
        """
        Parses UI-provided configurations, structures them back into database models,
        commits updates to relational tables using targeted delta updates, and triggers
        an in-memory hot reload.
        """
        from core.database.db_manager import (
            GlobalMapColumn, GlobalMapColumnKeyword, GlobalMapHeaderTextMapping,
            GlobalMapSheet, GlobalMapFooterLabelKeyword
        )

        try:
            if mapping_type == "shipping_header_map":
                # Flat UI format: {keyword: col_id}
                existing_kws = self.db.query(GlobalMapColumnKeyword).all()
                existing_map = {kw.keyword: kw for kw in existing_kws}
                
                # Ensure parent columns exist
                unique_cols = set(mappings.values())
                for col_id in unique_cols:
                    col_exists = self.db.query(GlobalMapColumn).filter_by(col_id=col_id).first()
                    if not col_exists:
                        self.db.add(GlobalMapColumn(col_id=col_id, excel_format="@"))
                
                # Deletes: keyword exists in DB but not in mappings
                for kw_text, kw_obj in existing_map.items():
                    if kw_text not in mappings:
                        self.db.delete(kw_obj)
                
                # Adds and Updates
                for kw_text, col_id in mappings.items():
                    if kw_text not in existing_map:
                        self.db.add(GlobalMapColumnKeyword(col_id=col_id, keyword=kw_text))
                    elif existing_map[kw_text].col_id != col_id:
                        existing_map[kw_text].col_id = col_id
                
            elif mapping_type == "footer_label_mappings":
                existing_footers = self.db.query(GlobalMapFooterLabelKeyword).all()
                existing_set = {f.keyword for f in existing_footers}
                new_set = {k.strip() for k in mappings.keys() if k.strip()}
                
                # Deletes
                for kw in (existing_set - new_set):
                    f_item = self.db.query(GlobalMapFooterLabelKeyword).filter_by(keyword=kw).first()
                    if f_item:
                        self.db.delete(f_item)
                
                # Adds
                for kw in (new_set - existing_set):
                    self.db.add(GlobalMapFooterLabelKeyword(keyword=kw))
                
            elif mapping_type in ("sheet_classifications", "sheet_mappings"):
                existing_sheets = self.db.query(GlobalMapSheet).all()
                existing_map = {s.sheet_name: s for s in existing_sheets}
                
                # Deletes
                for sheet_name, sheet_obj in existing_map.items():
                    if sheet_name not in mappings:
                        self.db.delete(sheet_obj)
                
                # Adds and Updates
                for sheet_name, proc_type in mappings.items():
                    clean_name = sheet_name.strip()
                    if not clean_name:
                        continue
                    if clean_name not in existing_map:
                        self.db.add(GlobalMapSheet(sheet_name=clean_name, processing_type=proc_type))
                    elif existing_map[clean_name].processing_type != proc_type:
                        existing_map[clean_name].processing_type = proc_type
                        
            elif mapping_type == "header_text_mappings":
                existing_overrides = self.db.query(GlobalMapHeaderTextMapping).all()
                existing_map = {o.raw_text: o for o in existing_overrides}
                
                # Ensure parent columns exist
                unique_cols = set(mappings.values())
                for col_id in unique_cols:
                    col_exists = self.db.query(GlobalMapColumn).filter_by(col_id=col_id).first()
                    if not col_exists:
                        self.db.add(GlobalMapColumn(col_id=col_id, excel_format="@"))
                
                # Deletes
                for raw_text, override_obj in existing_map.items():
                    if raw_text not in mappings:
                        self.db.delete(override_obj)
                
                # Adds and Updates
                for raw_text, col_id in mappings.items():
                    clean_text = raw_text.strip()
                    if not clean_text:
                        continue
                    if clean_text not in existing_map:
                        self.db.add(GlobalMapHeaderTextMapping(raw_text=clean_text, canonical_col_id=col_id))
                    elif existing_map[clean_text].canonical_col_id != col_id:
                        existing_map[clean_text].canonical_col_id = col_id
            
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
