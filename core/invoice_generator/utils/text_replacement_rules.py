"""
Centralized text replacement rules for invoice generation.
This module provides shared replacement rule builders to avoid duplication.
"""

from typing import List, Dict, Any, Optional


def build_replacement_rules(args: Optional[Any] = None) -> List[Dict[str, Any]]:
    """
    Build text replacement rules for template state.
    
    Args:
        args: Arguments object with DAF flag (optional)
    
    Returns:
        List of replacement rule dicts with keys:
        - find: Text to find
        - replace: Replacement text (for hardcoded rules)
        - data_path: Path to data (for data-driven rules)
        - is_date: Whether to format as date
        - match_mode: 'exact' or 'substring'
    """
    rules = []
    
    # Standard placeholder rules (data-driven)
    rules.extend([
        {
            "find": "JFINV",
            "data_path": ["invoice_info", "inv_no"],
            "fallback_path": ["processed_tables_data", "1", "col_inv_no", 0],
            "match_mode": "exact"
        },
        {
            "find": "JFTIME",
            "data_path": ["invoice_info", "inv_date"],
            "fallback_path": ["processed_tables_data", "1", "col_inv_date", 0],
            "is_date": True,
            "match_mode": "exact"
        },
        {
            "find": "JFREF",
            "data_path": ["invoice_info", "inv_ref"],
            "fallback_path": ["processed_tables_data", "1", "col_inv_ref", 0],
            "match_mode": "exact"
        },
        {
            "find": "[[CUSTOMER_NAME]]",
            "data_path": ["customer_info", "name"],
            "match_mode": "exact"
        },
        {
            "find": "[[CUSTOMER_ADDRESS]]",
            "data_path": ["customer_info", "address"],
            "match_mode": "exact"
        }
    ])
    
    # DAF-specific rules (hardcoded replacements)
    if args and hasattr(args, 'DAF') and args.DAF:
        rules.extend([
            {"find": "BINH PHUOC", "replace": "BAVET", "match_mode": "exact"},
            {"find": "BAVET, SVAY RIENG", "replace": "BAVET", "match_mode": "exact"},
            {"find": "BAVET,SVAY RIENG", "replace": "BAVET", "match_mode": "exact"},
            {"find": "BAVET, SVAYRIENG", "replace": "BAVET", "match_mode": "exact"},
            {"find": "BINH DUONG", "replace": "BAVET", "match_mode": "exact"},
            {"find": "FCA  BAVET,SVAYRIENG", "replace": "DAF BAVET", "match_mode": "exact"},
            {"find": "FCA: BAVET,SVAYRIENG", "replace": "DAF: BAVET", "match_mode": "exact"},
            {"find": "DAF  BAVET,SVAYRIENG", "replace": "DAF BAVET", "match_mode": "exact"},
            {"find": "DAF: BAVET,SVAYRIENG", "replace": "DAF: BAVET", "match_mode": "exact"},
            {"find": "SVAY RIENG", "replace": "BAVET", "match_mode": "exact"},
            {"find": "PORT KLANG", "replace": "BAVET", "match_mode": "exact"},
            {"find": "HCM", "replace": "BAVET", "match_mode": "exact"},
            {"find": "DAP", "replace": "DAF", "match_mode": "substring"},
            {"find": "FCA", "replace": "DAF", "match_mode": "substring"},
            {"find": "CIF", "replace": "DAF", "match_mode": "substring"},
        ])
    
    return rules
