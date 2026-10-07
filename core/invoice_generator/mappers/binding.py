from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional
from .models import InvoiceMode
from .source_normalizer import normalize_source_to_rows


@dataclass
class TableBinding:
    """
    Unified container binding normalized data, mapping rules, sheet layout, and mode for a single table.
    Acts as the boundary gate for mappers ("Unify First, Map Second").
    """
    data: List[Dict[str, Any]] = field(default_factory=list)
    mapping_rules: Dict[str, Any] = field(default_factory=dict)
    sheet_layout: Optional[Any] = None
    static_payload: Dict[str, Any] = field(default_factory=dict)
    footer_data: Dict[str, Any] = field(default_factory=dict)
    mode: InvoiceMode = field(default_factory=InvoiceMode)

    @classmethod
    def from_bundles(
        cls,
        data_config: Optional[Dict[str, Any]] = None,
        context_config: Optional[Dict[str, Any]] = None,
        layout_config: Optional[Dict[str, Any]] = None
    ) -> "TableBinding":
        """The official gate: ingests raw bundles, normalizes rows, and resolves mode."""
        data_config = data_config or {}
        context_config = context_config or {}
        sheet_layout = None
        if layout_config:
            from core.invoice_generator.models.config.layout import SheetLayoutModel
            sheet_layout = SheetLayoutModel.model_validate(layout_config.get('sheet_config', {}) or layout_config)

        raw_static = (
            (layout_config and layout_config.get('static_payload')) or
            data_config.get('static_payload') or
            {}
        )

        return cls(
            data=normalize_source_to_rows(data_config.get('data_source')),
            mapping_rules=data_config.get('mapping_rules') or {},
            sheet_layout=sheet_layout,
            static_payload=dict(raw_static),
            footer_data=data_config.get('footer_data') or {},
            mode=InvoiceMode.from_context(context_config)
        )

    def resolve(self) -> "ResolvedTableData":
        """Executes resolution pipeline by delegating to TableDataMapper."""
        from .mapper import TableDataMapper
        return TableDataMapper(binding=self).resolve()
