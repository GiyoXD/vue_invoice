from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Union

@dataclass
class MappingContext:
    """Immutable context holding execution flags, configuration, and data sources for mappers."""
    data_source_type: str = "aggregation"
    data_source: Union[Dict[str, Any], List[Any], None] = None
    mapping_rules: Dict[str, Any] = field(default_factory=dict)
    sheet_layout: Optional[Any] = None
    DAF_mode: bool = False
    custom_mode: bool = False
    static_content: Dict[str, Any] = field(default_factory=dict)
    pricing_net_weight: bool = False
    footer_data: Dict[str, Any] = field(default_factory=dict)

    @classmethod
    def from_bundles(
        cls,
        data_config: Dict[str, Any],
        context_config: Dict[str, Any],
        layout_config: Optional[Dict[str, Any]] = None
    ) -> "MappingContext":
        """Factory method to construct MappingContext from config bundles."""
        sheet_layout = None
        if layout_config:
            from core.invoice_generator.models.config.layout import SheetLayoutModel
            sheet_layout = SheetLayoutModel.model_validate(layout_config.get('sheet_config', {}) or layout_config)

        return cls(
            data_source_type=data_config.get('data_source_type', 'aggregation'),
            data_source=data_config.get('data_source'),
            mapping_rules=data_config.get('mapping_rules', {}),
            sheet_layout=sheet_layout,
            DAF_mode=context_config.get('DAF_mode', False),
            custom_mode=context_config.get('custom_mode', False),
            static_content=data_config.get('static_content', {}),
            pricing_net_weight=context_config.get('pricing_net_weight', False),
            footer_data=data_config.get('footer_data', {})
        )
