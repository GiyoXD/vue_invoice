import logging
from typing import Dict, Any, List

logger = logging.getLogger(__name__)

class ConfigValidator:
    """
    Validates a generated Invoice Configuration against the 'Ideal Master Config' structure.
    """

    
    REQUIRED_TOP_LEVEL_KEYS = ["_meta", "processing", "styling_bundle", "layout_bundle", "defaults"]
    REQUIRED_META_KEYS = ["config_version", "customer", "created_at"]
    
    def validate(self, config: Dict[str, Any]) -> List[Dict[str, str]]:
        """
        Validates the configuration structure against the Master Config Schema.
        Returns a list of error dictionaries containing 'issue', 'detail', and 'fix'.
        """
        errors = []

        # 1. Top Level Structure
        for key in self.REQUIRED_TOP_LEVEL_KEYS:
            if key not in config:
                errors.append({
                    "issue": f"Missing Top-Level Section: '{key}'",
                    "detail": f"The '{key}' section is a fundamental part of the Master Config structure.",
                    "fix": f"Ensure the generator creates a '{key}' dictionary at the root of the JSON."
                })
        
        # 2. Meta Data
        meta = config.get("_meta", {})
        for key in self.REQUIRED_META_KEYS:
            if key not in meta:
                 errors.append({
                     "issue": f"Missing Meta Field: '_meta.{key}'",
                     "detail": "Metadata is crucial for version control and file identification.",
                     "fix": f"Add '{key}' to the '_meta' section."
                 })

        # 3. Processing
        processing = config.get("processing", {})
        if "sheets" not in processing:
             errors.append({
                 "issue": "Missing Field: 'processing.sheets'",
                 "detail": "List of sheets to process (e.g. Invoice, Packing list).",
                 "fix": "Add 'sheets' list to 'processing'."
             })
        
        # 4. Styling Bundle (Deep Check)
        if "styling_bundle" in config:
            sb = config.get("styling_bundle", {})
            if "defaults" not in sb:
                errors.append({
                    "issue": "Missing Section: 'styling_bundle.defaults'",
                    "detail": "Global default styles are required.",
                    "fix": "Add 'defaults' to 'styling_bundle'."
                })
            
            # Check for per-sheet styling
            for sheet in processing.get("sheets", []):
                if sheet not in sb:
                     errors.append({
                         "issue": f"Missing Styling for Sheet: '{sheet}'",
                         "detail": f"Every sheet listed in 'processing' needs a matching entry in 'styling_bundle'.",
                         "fix": f"Add '{sheet}' to 'styling_bundle' with its column/row styles."
                     })

        # 5. Layout Bundle (Deep Check)
        if "layout_bundle" in config:
            lb = config.get("layout_bundle", {})
            
            # Check for per-sheet layout
            for sheet in processing.get("sheets", []):
                if sheet not in lb:
                     errors.append({
                         "issue": f"Missing Layout for Sheet: '{sheet}'",
                         "detail": f"Every sheet listed in 'processing' needs a matching entry in 'layout_bundle'.",
                         "fix": f"Add '{sheet}' to 'layout_bundle' with its sections."
                     })
                else:
                    # Deep check sheet layout structure
                    sl = lb[sheet]
                    if "_sections" not in sl:
                        errors.append({
                            "issue": f"Missing Sections Definition in '{sheet}' Layout",
                            "detail": "Layout needs explicit '_sections' list defining order.",
                            "fix": "Add '_sections': ['structure', 'data_flow', ...] to layout."
                        })

        return errors

    def _validate_sheet_config(self, sheet_conf: Dict[str, Any], sheet_name: str, errors: List[Dict[str, str]]):
        """Helper to validate individual sheet configurations (Deprecated/Unused for now but kept for ref)."""
        pass
            
