from typing import Any, Dict, Optional
from .fallback_resolver import apply_fallback
from .formula_resolver import resolve_mode_formula, parse_formula_def
from ..models.context import MappingContext

class RuleEngine:
    """Stateless evaluator for column mapping rules."""

    @staticmethod
    def evaluate_column_rule(
        target_id: str,
        mapping_rule: Dict[str, Any],
        row_dict: Dict[str, Any],
        context: MappingContext
    ) -> None:
        """
        Applies mapping rules (fallback, static values, or formula directives) to row_dict in-place.
        """
        if not isinstance(mapping_rule, dict):
            return

        # 1. Static value override
        if "static_value" in mapping_rule:
            row_dict[target_id] = mapping_rule["static_value"]
            return

        # 2. Fallback resolution
        apply_fallback(
            row_dict=row_dict,
            target_id=target_id,
            mapping_rule=mapping_rule,
            DAF_mode=context.DAF_mode,
            custom_mode=context.custom_mode
        )

        # 3. Formula rule resolution
        formula = resolve_mode_formula(mapping_rule, DAF_mode=context.DAF_mode, custom_mode=context.custom_mode)
        if formula is not None:
            parsed = parse_formula_def(formula)
            if parsed:
                row_dict[f"{target_id}_formula"] = parsed
