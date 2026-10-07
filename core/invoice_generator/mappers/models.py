from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple, Union
from pydantic import BaseModel, Field


class ResolvedTableFooter(BaseModel):
    grand_total: Optional[Dict[str, Any]] = None
    leather_summary: Optional[Any] = None
    weight_summary: Optional[Dict[str, Any]] = None
    pallet_summary_total: Optional[int] = None


class ResolvedTableData(BaseModel):
    data_rows: List[Dict[str, Any]] = Field(default_factory=list)
    num_data_rows: int = 0
    footer: Optional[ResolvedTableFooter] = None


@dataclass
class MappingContext:
    """Immutable context holding execution flags, configuration, and data sources for mappers."""
    data_source: Union[Dict[str, Any], List[Any], None] = None
    mapping_rules: Dict[str, Any] = field(default_factory=dict)
    sheet_layout: Optional[Any] = None
    invoice_format: str = "standard"
    pricing_basis: str = "sqft"
    static_payload: Dict[str, Any] = field(default_factory=dict)
    footer_data: Dict[str, Any] = field(default_factory=dict)

    @property
    def DAF_mode(self) -> bool:
        return self.invoice_format == "daf"

    @property
    def custom_mode(self) -> bool:
        return self.invoice_format == "custom"

    @property
    def item_merged_mode(self) -> bool:
        return self.invoice_format == "item_merged"

    @property
    def pricing_net_weight(self) -> bool:
        return self.pricing_basis == "net_weight"

    @classmethod
    def from_bundles(
        cls,
        data_config: Dict[str, Any],
        context_config: Dict[str, Any],
        layout_config: Optional[Dict[str, Any]] = None
    ) -> "MappingContext":
        """Constructs MappingContext from config bundles via TableBinding gate."""
        from .binding import TableBinding
        b = TableBinding.from_bundles(data_config, context_config, layout_config)
        return cls(
            data_source=b.data,
            mapping_rules=b.mapping_rules,
            sheet_layout=b.sheet_layout,
            static_payload=b.static_payload,
            footer_data=b.footer_data,
            invoice_format=b.mode.invoice_format,
            pricing_basis=b.mode.pricing_basis,
        )


VALID_INVOICE_FORMATS = frozenset({"standard", "custom", "daf", "item_merged"})
VALID_PRICING_BASES = frozenset({"sqft", "net_weight"})


@dataclass
class InvoiceMode:
    """Execution mode and formatting configuration for invoice generation."""
    invoice_format: str = "standard"
    pricing_basis: str = "sqft"

    def __init__(
        self,
        invoice_format: str = "standard",
        pricing_basis: str = "sqft",
    ):
        fmt = str(invoice_format or "standard").strip().lower()
        if fmt not in VALID_INVOICE_FORMATS:
            raise ValueError(f"Invalid invoice_format '{invoice_format}'. Must be one of {sorted(VALID_INVOICE_FORMATS)}")

        basis = str(pricing_basis or "sqft").strip().lower()
        if basis not in VALID_PRICING_BASES:
            raise ValueError(f"Invalid pricing_basis '{pricing_basis}'. Must be one of {sorted(VALID_PRICING_BASES)}")

        self.invoice_format = fmt
        self.pricing_basis = basis

    @property
    def DAF_mode(self) -> bool:
        return self.invoice_format == "daf"

    @property
    def custom_mode(self) -> bool:
        return self.invoice_format == "custom"

    @property
    def item_merged_mode(self) -> bool:
        return self.invoice_format == "item_merged"

    @property
    def pricing_net_weight(self) -> bool:
        return self.pricing_basis == "net_weight"

    @classmethod
    def from_context(cls, context: Optional[Dict[str, Any]] = None) -> "InvoiceMode":
        """Extracts InvoiceMode from runtime context bundle (args and metadata)."""
        if not context:
            return cls()
        args = context.get('args')
        fmt = context.get('invoice_format') or context.get('format')
        if not fmt and args:
            if getattr(args, 'item_merged', False): fmt = 'item_merged'
            elif getattr(args, 'custom', False): fmt = 'custom'
            elif getattr(args, 'DAF', False): fmt = 'daf'

        invoice_data = context.get('invoice_data') or {}
        metadata = invoice_data.get('metadata') or context.get('metadata') or {}
        basis = context.get('pricing_basis') or metadata.get('pricing_basis')
        if not basis:
            pricing_net_weight = context.get('pricing_net_weight') or metadata.get('pricing_net_weight')
            basis = "net_weight" if pricing_net_weight else "sqft"

        return cls(invoice_format=fmt or "standard", pricing_basis=basis or "sqft")

    @classmethod
    def from_dict(cls, data: Optional[Union["InvoiceMode", Dict[str, Any]]] = None) -> "InvoiceMode":
        if isinstance(data, cls):
            return data
        if data is None:
            return cls()
        if not isinstance(data, dict):
            raise ValueError(f"Expected dict or InvoiceMode, got {type(data).__name__}")
        if not data:
            return cls()
        fmt = data.get("invoice_format") or data.get("format")
        if not fmt:
            if data.get("DAF_mode") or data.get("daf") or data.get("DAF"):
                fmt = "daf"
            elif data.get("custom_mode") or data.get("custom"):
                fmt = "custom"
            elif data.get("item_merged_mode") or data.get("item_merged"):
                fmt = "item_merged"
            else:
                fmt = "standard"
        basis = data.get("pricing_basis") or ("net_weight" if data.get("pricing_net_weight") else "sqft")
        return cls(invoice_format=fmt, pricing_basis=basis)



__all__ = [
    "InvoiceMode",
    "MappingContext",
    "ResolvedTableData",
    "ResolvedTableFooter",
    "VALID_INVOICE_FORMATS",
    "VALID_PRICING_BASES",
]

