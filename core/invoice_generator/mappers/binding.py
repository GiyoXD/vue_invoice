from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Union
from .transforms import extract_table_data


@dataclass
class TableBinding:
    """
    Unified container binding pre-extracted data, mapping rules, sheet layout, and flags for a single table.
    Enables single-input mapping ("Unify First, Map Second").
    """
    data: Any = None
    mapping_rules: Dict[str, Any] = field(default_factory=dict)
    sheet_layout: Optional[Any] = None
    DAF_mode: bool = False
    custom_mode: bool = False
    static_content: Dict[str, Any] = field(default_factory=dict)
    pricing_net_weight: bool = False
    footer_data: Dict[str, Any] = field(default_factory=dict)
    data_source_type: str = "aggregation"

    @classmethod
    def from_bundles(
        cls,
        data_config: Dict[str, Any],
        context_config: Dict[str, Any],
        layout_config: Optional[Dict[str, Any]] = None
    ) -> "TableBinding":
        """
        Factory method to resolve and unify data, rules, and layout from configuration bundles.
        """
        sheet_layout = None
        if layout_config:
            from core.invoice_generator.models.config.layout import SheetLayoutModel
            sheet_layout = SheetLayoutModel.model_validate(layout_config.get('sheet_config', {}) or layout_config)

        data_source_type = data_config.get('data_source_type', 'aggregation')
        raw_data = data_config.get('data_source')
        table_data = extract_table_data(raw_data, data_source_type)

        return cls(
            data=table_data,
            mapping_rules=data_config.get('mapping_rules', {}) or {},
            sheet_layout=sheet_layout,
            DAF_mode=context_config.get('DAF_mode', False),
            custom_mode=context_config.get('custom_mode', False),
            static_content=data_config.get('static_content', {}) or {},
            pricing_net_weight=context_config.get('pricing_net_weight', False),
            footer_data=data_config.get('footer_data', {}) or {},
            data_source_type=data_source_type
        )

    def resolve(self) -> "ResolvedTableData":
        """
        Executes all mapping, fallbacks, static merges, pallet counts, and footer extractions.
        Outputs ONLY clean ResolvedTableData (data_rows and footer) with no flags required downstream.
        """
        from .models import ResolvedTableData, ResolvedTableFooter
        from .rules import parse_mapping_rules
        from .transforms import (
            prepare_data_rows,
            merge_static_content,
            format_pallet_counts,
            extract_summaries
        )

        column_id_map = {}
        column_map = {}
        parent_column_ids = []

        if self.sheet_layout:
            bundled_columns, column_map, column_id_map, _ = (
                self.sheet_layout.structure.resolve_mappings(
                    DAF_mode=self.DAF_mode,
                    custom_mode=self.custom_mode
                )
            )
            parent_column_ids = [col.id for col in bundled_columns if col.children]

        idx_to_header_map = {v: k for k, v in column_map.items()}

        parsed = parse_mapping_rules(
            mapping_rules=self.mapping_rules,
            column_id_map=column_id_map,
            idx_to_header_map=idx_to_header_map
        )

        desc_col_id = 'col_desc' if 'col_desc' in column_id_map else None
        desc_col_idx = column_id_map.get(desc_col_id, -1) if desc_col_id else -1

        data_rows, num_data_rows = prepare_data_rows(
            data_source_type=self.data_source_type,
            data_source=self.data,
            dynamic_mapping_rules=parsed['dynamic_mapping_rules'],
            column_id_map=column_id_map,
            idx_to_header_map=idx_to_header_map,
            desc_col_idx=desc_col_idx,
            num_static_labels=parsed['num_static_labels'],
            static_value_map=parsed['static_value_map'],
            DAF_mode=self.DAF_mode,
            custom_mode=self.custom_mode,
            parent_column_ids=parent_column_ids,
            pricing_net_weight=self.pricing_net_weight
        )

        merge_static_content(
            data_rows=data_rows,
            static_content=self.static_content,
            dynamic_mapping_rules=parsed['dynamic_mapping_rules'],
            DAF_mode=self.DAF_mode,
            custom_mode=self.custom_mode
        )

        pallet_col_id = 'col_pallet_count' if 'col_pallet_count' in column_id_map else None
        if pallet_col_id:
            format_pallet_counts(
                data_rows=data_rows,
                num_data_rows=num_data_rows,
                pallet_col_id=pallet_col_id
            )

        leather_summary, weight_summary, pallet_summary_total = extract_summaries(
            data_source=self.data,
            footer_data=self.footer_data
        )

        resolved_footer = ResolvedTableFooter(
            leather_summary=leather_summary,
            weight_summary=weight_summary,
            pallet_summary_total=pallet_summary_total
        )

        return ResolvedTableData(
            data_rows=data_rows,
            num_data_rows=num_data_rows,
            footer=resolved_footer
        )

