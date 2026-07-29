from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Union
from .transforms import extract_table_data


@dataclass
class TableBinding:
    """
    Unified container binding pre-extracted data, mapping rules, sheet layout, and generic options for a single table.
    Enables single-input mapping ("Unify First, Map Second").
    """
    data: Any = None
    mapping_rules: Dict[str, Any] = field(default_factory=dict)
    sheet_layout: Optional[Any] = None
    static_payload: Dict[str, Any] = field(default_factory=dict)
    footer_data: Dict[str, Any] = field(default_factory=dict)
    data_source_type: str = "aggregation"
    options: Dict[str, Any] = field(default_factory=dict)

    @classmethod
    def from_bundles(
        cls,
        data_config: Dict[str, Any],
        context_config: Dict[str, Any],
        layout_config: Optional[Dict[str, Any]] = None
    ) -> "TableBinding":
        """
        Factory method to resolve and unify data, rules, layout, and options from configuration bundles.
        """
        sheet_layout = None
        if layout_config:
            from core.invoice_generator.models.config.layout import SheetLayoutModel
            sheet_layout = SheetLayoutModel.model_validate(layout_config.get('sheet_config', {}) or layout_config)

        data_source_type = data_config.get('data_source_type', 'aggregation')
        raw_data = data_config.get('data_source')
        table_data = extract_table_data(raw_data, data_source_type)

        raw_static = (
            (layout_config and layout_config.get('static_payload')) or
            data_config.get('static_payload') or
            {}
        )

        options = {
            'DAF_mode': context_config.get('DAF_mode', False),
            'custom_mode': context_config.get('custom_mode', False),
            'pricing_net_weight': context_config.get('pricing_net_weight', False),
            **context_config.get('options', {})
        }

        return cls(
            data=table_data,
            mapping_rules=data_config.get('mapping_rules', {}) or {},
            sheet_layout=sheet_layout,
            static_payload=raw_static or {},
            footer_data=data_config.get('footer_data', {}) or {},
            data_source_type=data_source_type,
            options=options
        )

    def resolve(self) -> "ResolvedTableData":
        """
        Executes resolution pipeline by delegating to TableDataMapper.
        """
        from .mapper import TableDataMapper
        return TableDataMapper(binding=self).resolve()

