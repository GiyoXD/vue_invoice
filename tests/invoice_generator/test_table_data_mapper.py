import pytest
from typing import Dict, Any

from core.invoice_generator.mappers import (
    TableDataMapper,
    ResolvedTableData,
    TableBinding,
    MappingContext,
    InvoiceMode,
    resolve_static_placeholders,
    populate_static_content,
    extract_summaries,
    format_pallet_counts
)


def test_helpers_populate_static_content():
    """Verify resolve_static_placeholders and populate_static_content sequence."""
    data_rows = [{"col_item": "dynamic1"}, {"col_item": "dynamic2"}]
    static_payload = {"col_static": ["Static1", "{col_desc_fallback}"]}

    resolved_static = resolve_static_placeholders(
        static_payload=static_payload,
        desc_fallback_str="Standard Desc"
    )
    assert resolved_static["col_static"][1] == "Standard Desc"

    populate_static_content(
        data_rows=data_rows,
        static_payload=resolved_static
    )

    # Values should be merged into col_static
    assert data_rows[0]["col_static"] == "Static1"
    assert data_rows[1]["col_static"] == "Standard Desc"

    # Test extending rows if we have more static values than data rows
    data_rows_short = [{"col_item": "dynamic1"}]
    static_payload_long = {"col_static": ["Static1", "Static2"]}
    populate_static_content(
        data_rows=data_rows_short,
        static_payload=static_payload_long
    )
    assert len(data_rows_short) == 2
    assert data_rows_short[0]["col_static"] == "Static1"
    assert data_rows_short[1]["col_static"] == "Static2"


def test_helpers_extract_summaries():
    """Verify that extract_summaries handles various source types and footer data formats."""
    footer_data = {
        "leather_summary": "l_sum",
        "weight_summary": "w_sum",
        "grand_total": {"col_pallet_count": 4}
    }
    l, w, p = extract_summaries(footer_data=footer_data)
    assert l == "l_sum"
    assert w == "w_sum"
    assert p == 4

    # Pre-calculated in footer_data grand_total
    footer_data = {
        "grand_total": {
            "col_pallet_count": 10
        }
    }
    _, _, p = extract_summaries(None, footer_data)
    assert p == 10

    # Multi-table table_totals list
    footer_data_multi = {
        "table_totals": [
            {"col_pallet_count": 3},
            {"col_pallet_count": 5}
        ]
    }
    _, _, p0 = extract_summaries(None, footer_data_multi)
    assert p0 == 3


def test_helpers_format_pallet_counts():
    """Verify that format_pallet_counts carries values forward for vertical merging."""
    data_rows = [
        {"col_item": "Item 1", "col_pallet_count": "1-5"},
        {"col_item": "Item 2", "col_pallet_count": 0},
        {"col_item": "Item 3", "col_pallet_count": "2-5"}
    ]
    format_pallet_counts(
        data_rows=data_rows,
        num_data_rows=3,
        pallet_col_id="col_pallet_count"
    )

    assert data_rows[0]["col_pallet_count"] == "1-5"
    assert data_rows[1]["col_pallet_count"] == "1-5"
    assert data_rows[2]["col_pallet_count"] == "2-5"


def test_mapping_context():
    """Verify MappingContext initialization and fields."""
    from core.invoice_generator.mappers import MappingContext

    context = MappingContext(
        data_source=[{"col_a": "val1"}],
        invoice_format="daf",
    )
    assert context.DAF_mode is True
    assert context.custom_mode is False
    assert context.data_source == [{"col_a": "val1"}]


def test_table_binding_single_input_mapping():
    """Verify single-input TableBinding construction and mapper resolution."""
    from core.invoice_generator.mappers import TableBinding, TableDataMapper

    data_config = {
        "data_source": [{"col_item": "leather_bag"}],
        "mapping_rules": {"col_item": {"column": "col_item"}}
    }
    context_config = {"DAF_mode": False}

    binding = TableBinding.from_bundles(data_config, context_config)
    assert binding.data == [{"col_item": "leather_bag"}]

    mapper = TableDataMapper(binding=binding)
    assert mapper.binding is binding
    assert mapper.data_source == binding.data


def test_populate_static_content_empty_data_rows():
    """Verify resolve_static_placeholders + populate_static_content when data_rows is initially empty."""
    data_rows = []
    static_payload = {"col_static": ["VENDOR#:", "Des: {col_desc_fallback}", "MADE IN CAMBODIA"]}

    resolved_static = resolve_static_placeholders(
        static_payload=static_payload,
        desc_fallback_str="Standard Cow Leather"
    )

    populate_static_content(
        data_rows=data_rows,
        static_payload=resolved_static
    )

    assert len(data_rows) == 3
    assert data_rows[0]["col_static"] == "VENDOR#:"
    assert data_rows[1]["col_static"] == "Des: Standard Cow Leather"
    assert data_rows[2]["col_static"] == "MADE IN CAMBODIA"


def test_price_adjustment_flag_behavior():
    """Verify price adjustment modifies amounts and mapping works without _price_adjust."""
    from core.invoice_generator.mappers.transforms import prepare_data_rows
    from core.invoice_generator.utils.aggregation_modifier import _apply_to_row_list

    rows = [{"col_amount": "100.0", "col_qty_sf": "10", "col_unit_price": "10.0"}]
    _apply_to_row_list(rows, 10.0)

    assert float(rows[0]["col_amount"]) == 110.0

    mapping_rules = {
        "col_amount": {"column": "col_amount"},
        "col_unit_price": {"column": "col_unit_price"},
        "col_qty_sf": {"column": "col_qty_sf"}
    }
    data_rows, _ = prepare_data_rows(
        data_source=rows,
        dynamic_mapping_rules=mapping_rules,
    )
    assert len(data_rows) == 1
    assert data_rows[0]["col_amount"] == "110.0"


def test_table_data_mapper_properties():
    """Verify TableDataMapper property delegation."""
    # Property delegation
    mapper_std = TableDataMapper(TableBinding(mode=InvoiceMode("standard")))
    assert mapper_std.mode.invoice_format == "standard"
    assert mapper_std.mode.DAF_mode is False
    assert mapper_std.mode.custom_mode is False
    assert mapper_std.mode.item_merged_mode is False

    mapper_daf = TableDataMapper(TableBinding(mode=InvoiceMode("daf")))
    assert mapper_daf.mode.invoice_format == "daf"
    assert mapper_daf.mode.DAF_mode is True
    assert mapper_daf.mode.custom_mode is False
    assert mapper_daf.mode.item_merged_mode is False

    mapper_custom = TableDataMapper(TableBinding(mode=InvoiceMode("custom")))
    assert mapper_custom.mode.invoice_format == "custom"
    assert mapper_custom.mode.custom_mode is True
    assert mapper_custom.mode.DAF_mode is False
    assert mapper_custom.mode.item_merged_mode is False

    mapper_merged = TableDataMapper(TableBinding(mode=InvoiceMode("item_merged")))
    assert mapper_merged.mode.invoice_format == "item_merged"
    assert mapper_merged.mode.item_merged_mode is True


def test_bundle_factories_format_alias():
    """Verify bundle factories handle format='daf' alias."""
    from core.invoice_generator.mappers import TableBinding, MappingContext

    data_config = {"data_source": []}

    # format="daf" alias works across factories
    mapper_daf = TableDataMapper.create_from_bundles(data_config, {"format": "daf"})
    assert mapper_daf.mode.invoice_format == "daf"
    assert mapper_daf.mode.DAF_mode is True

    binding_daf = TableBinding.from_bundles(data_config, {"format": "daf"})
    assert binding_daf.mode.invoice_format == "daf"
    assert binding_daf.mode.DAF_mode is True

    context_daf = MappingContext.from_bundles(data_config, {"format": "daf"})
    assert context_daf.invoice_format == "daf"
    assert context_daf.DAF_mode is True


def test_binding_resolution_pipeline():
    """Verify TableBinding resolution pipeline and mode properties."""
    from core.invoice_generator.mappers import TableBinding

    data_config = {
        "data_source": [{"col_item": "leather_bag"}],
        "mapping_rules": {"col_item": {"column": "col_item"}}
    }
    context_config = {"custom_mode": False}

    binding = TableBinding.from_bundles(data_config, context_config)
    assert binding.mode.custom_mode is False
    assert binding.mode.invoice_format == "standard"

    resolved = binding.resolve()
    assert isinstance(resolved, ResolvedTableData)
    assert resolved.num_data_rows == 1
    assert resolved.data_rows[0]["col_item"] == "leather_bag"


def test_table_footer_mapper_from_binding():
    """Verify TableFooterMapper initialized with a TableBinding resolves correctly."""
    from core.invoice_generator.mappers import TableBinding, TableFooterMapper, ResolvedTableFooter

    data_config = {
        "data_source": [{"col_item": "leather_bag"}],
        "footer_data": {
            "grand_total": {"col_pallet_count": 5}
        }
    }
    context_config = {"invoice_format": "standard"}

    binding = TableBinding.from_bundles(data_config, context_config)
    footer_mapper = TableFooterMapper(binding=binding)
    resolved_footer = footer_mapper.resolve()

    assert isinstance(resolved_footer, ResolvedTableFooter)
    assert resolved_footer.pallet_summary_total == 5
    assert footer_mapper.mode is binding.mode
    assert footer_mapper.context.invoice_format == "standard"
    assert footer_mapper.context.pricing_basis == "sqft"


def test_table_footer_mapper_create_from_bundles():
    """Verify TableFooterMapper.create_from_bundles creates mapper with self.mode and resolves properly."""
    from core.invoice_generator.mappers import TableFooterMapper, ResolvedTableFooter

    data_config = {
        "data_source": [{"col_item": "leather_bag"}],
        "footer_data": {
            "grand_total": {"col_pallet_count": 3}
        }
    }
    context_config = {"invoice_format": "standard", "pricing_basis": "sqft"}

    footer_mapper = TableFooterMapper.create_from_bundles(data_config, context_config)
    assert footer_mapper.mode is not None
    assert footer_mapper.mode.invoice_format == "standard"
    assert footer_mapper.mode.pricing_basis == "sqft"

    resolved_footer = footer_mapper.resolve()
    assert isinstance(resolved_footer, ResolvedTableFooter)
    assert resolved_footer.pallet_summary_total == 3



def test_table_data_mapper_mode_properties():
    """Verify TableDataMapper.mode always reflects the resolved mode."""
    mapper_default = TableDataMapper()
    assert mapper_default.mode.invoice_format == "standard"

    mapper_daf = TableDataMapper(TableBinding(mode=InvoiceMode("daf")))
    assert mapper_daf.mode.invoice_format == "daf"

    mapper_custom = TableDataMapper(TableBinding(mode=InvoiceMode("custom")))
    assert mapper_custom.mode.invoice_format == "custom"

    mapper_merged = TableDataMapper(TableBinding(mode=InvoiceMode("item_merged")))
    assert mapper_merged.mode.invoice_format == "item_merged"

    data_config = {"data_source": [{"col_po": "PO-1"}]}
    binding = TableBinding.from_bundles(data_config, {"format": "daf"})
    mapper_from_binding = TableDataMapper(binding=binding)
    assert mapper_from_binding.mode.invoice_format == "daf"

    context = MappingContext.from_bundles(data_config, {"format": "custom"})
    mapper_from_context = TableDataMapper(TableBinding(data=context.data_source, mode=InvoiceMode(invoice_format=context.invoice_format, pricing_basis=context.pricing_basis)))
    assert mapper_from_context.mode.invoice_format == "custom"
