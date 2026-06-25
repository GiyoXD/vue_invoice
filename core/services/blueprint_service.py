import json
import logging
import copy
from typing import Dict, List, Optional, Any
from sqlalchemy.orm import Session

from core.orchestrator import Orchestrator
from core.services.mapping_service import MappingService
from core.database.repositories import BlueprintRepository
from core.utils.storage import TempFileStorage

logger = logging.getLogger(__name__)

class BlueprintService:
    """
    Service layer handles business logic for blueprint template uploading, 
    scanning, parsing, manual mapping configurations, bundle generation,
    and repository database operations.
    """

    def __init__(self, db: Session):
        self.db = db
        self.orchestrator = Orchestrator()
        self.mapping_service = MappingService(db)
        self.repository = BlueprintRepository(db)

    def scan_and_analyze_template(self, filename: str, file_stream) -> Dict[str, Any]:
        """
        Saves the uploaded Excel template, scans it using Orchestrator.analyze_template,
        and extracts any unrecognized column headers or footers.
        """
        file_path = TempFileStorage.get_temp_path(filename, prefix="scan")
        file_token = file_path.name
        
        TempFileStorage.save_file(file_stream, file_path)
            
        try:
            analysis_json_str = self.orchestrator.analyze_template(file_path)
            analysis = json.loads(analysis_json_str)
            
            unknown_headers = []
            unconfirmed_footers = []
            
            for sheet in analysis.get("sheets", []):
                for header in sheet.get("header_positions", []):
                    col_id = header.get("col_id")
                    if col_id and col_id.startswith("col_unknown"):
                        unknown_headers.append(header.get("keyword"))
                
                uf = sheet.get("unconfirmed_footer")
                if uf:
                    unconfirmed_footers.append(uf)
            
            # Deduplicate items
            unknown_headers = list(set(unknown_headers))
            unconfirmed_footers = list(set(unconfirmed_footers))
            
            status = "needs_mapping" if (unknown_headers or unconfirmed_footers) else "clean"
            
            return {
                "status": status,
                "file_token": file_token,
                "unknown_headers": unknown_headers,
                "unconfirmed_footers": unconfirmed_footers,
                "warnings": analysis.get("warnings", []),
                "preview_analysis": analysis
            }
        except Exception as e:
            # Clean up file if parsing fails
            TempFileStorage.delete_file(file_path)
            raise e

    def analyze_template_legacy(self, filename: str, file_stream, ignore_missing_description: bool = False) -> Dict[str, Any]:
        """
        Saves the template Excel and parses missing headers and footers to match the legacy
        endpoints in templates router.
        """
        file_path = TempFileStorage.get_temp_path(filename)
        safe_filename = file_path.name
        
        TempFileStorage.save_file(file_stream, file_path)
            
        try:
            analysis_json_str = self.orchestrator.analyze_template(
                file_path, legacy_format=True, ignore_missing_description=ignore_missing_description
            )
            analysis = json.loads(analysis_json_str)
            
            missing_headers, missing_footers = self._extract_legacy_missing_headers_and_footers(analysis)
            
            return {
                "missing_headers": missing_headers,
                "missing_footers": missing_footers,
                "warnings": analysis.get("warnings", []),
                "temp_filename": safe_filename,
                "suggested_prefix": safe_filename.split('.')[0]
            }
        except Exception as e:
            # Clean up intermediate file only on analysis failure
            TempFileStorage.delete_file(file_path)
            raise e

    def generate_blueprint(
        self,
        file_token: str,
        customer_code: str,
        locale: str,
        mappings: Dict[str, str],
        footer_mappings: List[str],
        pricing_mode: str = "standard",
        ignore_missing_description: bool = False,
        bundle_dir_name: str = ""
    ) -> Dict[str, Any]:
        """
        Unifies mapping updates, configurations generation, and saves the final template output 
        to database. Cleans up temp upload and cache folders.
        """
        file_path = TempFileStorage.get_temp_path(file_token)
        
        if not file_path.exists():
            raise FileNotFoundError("Temporary template file expired or invalid. Please re-scan.")
            
        try:
            # 1. Update/Merge Column and Footer mappings dynamically in global config in a single batch
            if mappings or footer_mappings:
                from core.database.db_manager import get_global_mapping_config, save_global_mapping_config
                data = get_global_mapping_config(self.db)
                updated = False
                
                if mappings:
                    filtered_mappings = {k: v for k, v in mappings.items() if v and v != "col_unknown"}
                    if filtered_mappings:
                        if "header_text_mappings" not in data:
                            data["header_text_mappings"] = {"mappings": {}}
                        if "mappings" not in data["header_text_mappings"]:
                            data["header_text_mappings"]["mappings"] = {}
                        data["header_text_mappings"]["mappings"].update(filtered_mappings)
                        updated = True

                if footer_mappings:
                    if "footer_label_mappings" not in data:
                        data["footer_label_mappings"] = {"keywords": []}
                    existing_footers = data["footer_label_mappings"].get("keywords", [])
                    for fm in footer_mappings:
                        if fm not in existing_footers:
                            existing_footers.append(fm)
                    data["footer_label_mappings"]["keywords"] = existing_footers
                    updated = True

                if updated:
                    save_global_mapping_config(data, self.db)
                    self.mapping_service.reload_dynamic_state(data)

            # 3. Retrieve any existing layout JSON configuration to preserve user custom changes
            existing_template_json = None
            existing = self.repository.get_blueprint(customer_code, locale)
            if existing:
                existing_template_json = existing.template_json
                
            # 4. Generate Blueprint config and spreadsheet bundle via Orchestrator
            config_data, template_json_data, template_xlsx_bytes = self.orchestrator.generate_blueprint_bundle(
                template_path=file_path,
                custom_prefix=customer_code,
                runtime_mappings=mappings,
                bundle_dir_name=bundle_dir_name or None,
                pricing_mode=pricing_mode,
                ignore_missing_description=ignore_missing_description,
                in_memory=True,
                existing_template_json=existing_template_json
            )
            
            # 5. Save generated configuration and binaries to repositories db
            self.repository.save_blueprint(
                customer_code=customer_code,
                locale=locale,
                config_data=config_data,
                template_json_data=template_json_data,
                xlsx_bytes=template_xlsx_bytes,
                filename=f"{customer_code}_{locale}.xlsx"
            )
            
            # 6. Clear runtime cache folder for this specific customer
            TempFileStorage.clear_runtime_cache(customer_code, locale)
            
            return {
                "status": "success",
                "message": f"Blueprint generated and saved to database for {customer_code}"
            }
            
        finally:
            # Clean up the uploaded Excel file now that the blueprint is generated/failed
            TempFileStorage.delete_file(file_path)

    @staticmethod
    def get_mapping_options() -> List[Dict[str, str]]:
        """
        Return list of valid system columns for mapping.
        """
        from core.blueprint_generator.schema import BlueprintSchema
        
        options = []
        # Sort by ID
        sorted_cols = sorted(BlueprintSchema.COLUMNS.values(), key=lambda c: c.id)
        
        for col in sorted_cols:
            options.append({
                "id": col.id,
                "label": BlueprintService._format_label(col.id),
                "description": f"Internal ID: {col.id}"
            })
            
        return options

    @staticmethod
    def _format_label(col_id: str) -> str:
        """col_qty_pcs -> Qty Pcs"""
        return col_id.replace("col_", "").replace("_", " ").title()

    # --- Internal Helpers ---

    def _get_header_suggestions(self, header_text: str) -> str:
        header_lower = (header_text or "").lower()
        suggestions = {
            "col_po": ['p.o', 'po'], 
            "col_item": ['item', 'no.'], 
            "col_desc": ['description', 'desc'],
            "col_qty_sf": ['quantity', 'qty'], 
            "col_unit_price": ['unit', 'price'], 
            "col_amount": ['amount', 'total', 'value'],
            "col_net": ['n.w', 'net'], 
            "col_gross": ['g.w', 'gross'], 
            "col_cbm": ['cbm'],
            "col_pallet": ['pallet'], 
            "col_remarks": ['remarks', 'notes'], 
            "col_static": ['mark', 'note']
        }
        for col_id, keywords in suggestions.items():
            if any(word in header_lower for word in keywords):
                return col_id
        return "col_unknown"

    def _extract_legacy_missing_headers_and_footers(self, analysis_data: dict) -> tuple:
        missing_headers = []
        missing_footers = []
        
        for sheet in analysis_data.get('sheets', []):
            for header_pos in sheet.get('header_positions', []):
                col_id = header_pos.get('col_id', '')
                header_text = header_pos.get('keyword', '')
                if col_id.startswith("col_unknown"):
                    missing_headers.append({
                        "text": header_text, 
                        "suggestion": self._get_header_suggestions(header_text)
                    })
                    
            uf = sheet.get("unconfirmed_footer")
            if uf:
                missing_footers.append(uf)
                
        return missing_headers, list(set(missing_footers))

    def list_blueprints(self) -> List[Dict[str, Any]]:
        import datetime
        rows = self.repository.get_all_blueprints()
        templates = []
        for row in rows:
            try:
                data = row.template_json
                source = data.get("fingerprint", {}).get("source_file", "Unknown")
                templates.append({
                    "name": f"{row.customer_code}_{row.locale}",
                    "customer_code": row.customer_code,
                    "locale": row.locale,
                    "bundle_name": row.customer_code,
                    "modified": row.updated_at.isoformat() if row.updated_at else datetime.datetime.now().isoformat(),
                    "source_file": source
                })
            except Exception:
                logger.exception(f"Failed to parse DB blueprint {row.customer_code}_{row.locale}")
        return templates

    def view_blueprint(self, customer_code: str, locale: str = "KH") -> Optional[Dict[str, Any]]:
        row = self.repository.get_blueprint(customer_code, locale)
        if not row:
            return None
            
        data = copy.deepcopy(row.template_json)
        config_data = copy.deepcopy(row.config_json)
        info = config_data.get("table_info", {})
        if info:
            if "table_info" not in data:
                data["table_info"] = {}
            data["table_info"].update(info)

        # Inject column mappings so the frontend knows exactly which column is mapped to which system ID
        from core.blueprint_generator.schema import BlueprintSchema
        layout_bundle = config_data.get("layout_bundle", {})
        column_mappings = {}
        for sheet_name, sheet_conf in layout_bundle.items():
            if not isinstance(sheet_conf, dict):
                continue
            cols = sheet_conf.get("structure", {}).get("columns", [])
            sheet_cols = []
            for c in cols:
                if not isinstance(c, dict):
                    continue
                col_id = c.get("id")
                schema_col = BlueprintSchema.COLUMNS.get(col_id)
                width = c.get("width") or (schema_col.width if schema_col else 15.0)
                sheet_cols.append({
                    "id": col_id,
                    "header": c.get("header"),
                    "width": width,
                    "format": c.get("format", "@")
                })
            column_mappings[sheet_name] = sheet_cols
        data["column_mappings"] = column_mappings

        # Inject parsed client profile so the Template Inspector can display it
        from core.invoice_generator.extractors.template_client_profile_parser import TemplateClientProfileParser
        from core.blueprint_generator.internal.scanner.models import TemplateLayout
        from openpyxl.utils import get_column_letter

        invoice_sheet = data.get("template_layout", {}).get("Invoice", {})
        if invoice_sheet:
            layout = TemplateLayout.from_dict(invoice_sheet)
            header_content = {}
            for row_obj in layout.header_rows:
                r_idx = row_obj.relative_index + 1
                for cell in row_obj.cells:
                    if cell.value is not None:
                        coord = f"{get_column_letter(cell.col_index)}{r_idx}"
                        header_content[coord] = cell.value

            if header_content:
                parser = TemplateClientProfileParser(header_content)
                data["client_profile"] = {
                    "fullname": parser.get_client_fullname(),
                    "address":  parser.get_client_address(),
                    "contact":  parser.get_client_contact(),
                    "shipping": parser.get_shipping_method(),
                }
        return data

    def update_blueprint_cell(self, customer_code: str, locale: str, sheet_name: str, cell_address: str, overrides: Dict[str, str]) -> bool:
        from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
        from core.blueprint_generator.internal.scanner.models import TemplateLayout, UnitRow, UnitCell
        row = self.repository.get_blueprint(customer_code, locale)
        if not row:
            return False
            
        data = copy.deepcopy(row.template_json)
        sheet = data.get("template_layout", {}).get(sheet_name)
        if not sheet:
            return False

        layout = TemplateLayout.from_dict(sheet)
        if not layout:
            return False

        try:
            col_letter, row_val = coordinate_from_string(cell_address)
            col_idx = column_index_from_string(col_letter)
        except Exception:
            return False

        max_header_r = max((r.relative_index + 1 for r in layout.header_rows), default=0)
        is_f = row_val > max_header_r

        if is_f:
            rel = row_val - max_header_r - 1
            # Find or create UnitRow in footer_rows
            row_item = next((r for r in layout.footer_rows if r.relative_index == rel), None)
            if not row_item:
                row_item = UnitRow(relative_index=rel, cells=[])
                layout.footer_rows.append(row_item)
                layout.footer_rows.sort(key=lambda r: r.relative_index)
            
            # Find or create UnitCell
            cell = next((c for c in row_item.cells if c.col_index == col_idx), None)
            if not cell:
                cell = UnitCell(col_index=col_idx)
                row_item.cells.append(cell)
                row_item.cells.sort(key=lambda c: c.col_index)
        else:
            rel = row_val - 1
            # Find or create UnitRow in header_rows
            row_item = next((r for r in layout.header_rows if r.relative_index == rel), None)
            if not row_item:
                row_item = UnitRow(relative_index=rel, cells=[])
                layout.header_rows.append(row_item)
                layout.header_rows.sort(key=lambda r: r.relative_index)
            
            # Find or create UnitCell
            cell = next((c for c in row_item.cells if c.col_index == col_idx), None)
            if not cell:
                cell = UnitCell(col_index=col_idx)
                row_item.cells.append(cell)
                row_item.cells.sort(key=lambda c: c.col_index)

        # Apply overrides
        val = cell.value
        curr_map = val if isinstance(val, dict) else {"default": str(val) if val is not None else ""}
        for m, v in overrides.items():
            if v is None or (isinstance(v, str) and not v.strip()):
                if m in curr_map:
                    del curr_map[m]
            else:
                curr_map[m] = v

        if len(curr_map) == 1 and "default" in curr_map:
            cell.value = curr_map["default"]
        elif not curr_map:
            cell.value = ""
        else:
            cell.value = curr_map

        # Save back to database representation
        data["template_layout"][sheet_name] = layout.to_dict()
        row.template_json = data
        self.db.commit()
        
        # Clear cache file so it regenerates from updated DB
        TempFileStorage.clear_runtime_cache(customer_code, locale)
        return True

    def update_blueprint_notes(self, customer_code: str, locale: str, notes: str) -> bool:
        row = self.repository.get_blueprint(customer_code, locale)
        if not row:
            return False
            
        data = copy.deepcopy(row.template_json)
        data["notes"] = notes
        row.template_json = data
        self.db.commit()
        
        # Clear cache
        TempFileStorage.clear_runtime_cache(customer_code, locale)
        return True

    def delete_blueprint(self, customer_code: str, locale: str = "KH") -> bool:
        deleted = self.repository.delete_blueprint(customer_code, locale)
        if not deleted:
            return False
            
        # Clear cache
        TempFileStorage.clear_runtime_cache(customer_code, locale)
        return True
