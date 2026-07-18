import logging
from typing import Dict, Any, Optional
from core.system_config import ConfigurationError

logger = logging.getLogger(__name__)

class ContextResolver:
    """Resolves dynamic, runtime context variables (weights, summaries, pallet calculations)."""
    
    def __init__(self, sheet_name: str, args: Any, invoice_data: Optional[Dict[str, Any]], 
                 pallets: int, config_loader: Any, context_overrides: Dict[str, Any]):
        self.sheet_name = sheet_name
        self.args = args
        self.invoice_data = invoice_data
        self.pallets = pallets
        self.config_loader = config_loader
        self.context_overrides = context_overrides

    def get_context_bundle(self, table_key: Optional[str] = None, **additional_context) -> Dict[str, Any]:
        base_context = {
            'sheet_name': self.sheet_name,
            'args': self.args,
            'invoice_data': self.invoice_data,
            'pallets': self.pallets,
            'all_sheet_configs': self.config_loader.get_raw_config().get('layout_bundle', {}),
        }
        
        # Aggregate pre-calculated summaries from footer_data.grand_total
        footer_data = (self.invoice_data or {}).get('footer_data') or {}
        grand_total = footer_data.get('grand_total') or {}
        
        if grand_total:
            try:
                total_net = float(grand_total.get('col_net') or 0)
                total_gross = float(grand_total.get('col_gross') or 0)
            except (ValueError, TypeError) as e:
                raise ConfigurationError(
                    f"CRITICAL: Net or Gross weight in grand_total has invalid numeric format for sheet '{self.sheet_name}'. "
                    f"col_net='{grand_total.get('col_net')}', col_gross='{grand_total.get('col_gross')}'."
                ) from e
                
            try:
                total_pallets = int(grand_total.get('col_pallet_count') or 0)
            except (ValueError, TypeError) as e:
                raise ConfigurationError(
                    f"CRITICAL: Pallet count in grand_total has invalid integer format for sheet '{self.sheet_name}'. "
                    f"col_pallet_count='{grand_total.get('col_pallet_count')}'."
                ) from e
            
            if not grand_total.get('col_pallet_count'):
                logger.warning("⚠ No col_pallet_count in footer_data.grand_total. Pallet count will be 0.")

            summaries = {
                'total_net_weight': total_net,
                'total_gross_weight': total_gross,
                'total_pallets': total_pallets
            }
            base_context.update(summaries)
            logger.debug(f"Added global summaries to context: {summaries}")
        
        # Merge overrides and additional context
        base_context.update(self.context_overrides)
        base_context.update(additional_context)
        
        return base_context

    def get_footer_data(
        self,
        footer_row_start_idx: int,
        data_start_row: int,
        data_end_row: int,
        pallet_count: Optional[int] = None,
        leather_summary: Optional[Dict] = None,
        weight_summary: Optional[Dict] = None
    ):
        from ...models.footer import FooterData
        
        final_pallets = pallet_count if pallet_count is not None else self.pallets
        final_weight_summary = {'net': 0.0, 'gross': 0.0}
        
        if weight_summary:
            final_weight_summary.update(weight_summary)
            
        if final_weight_summary['net'] == 0 and final_weight_summary['gross'] == 0:
            context = self.get_context_bundle()
            if 'total_net_weight' in context:
                final_weight_summary['net'] = context['total_net_weight']
            if 'total_gross_weight' in context:
                final_weight_summary['gross'] = context['total_gross_weight']
                
        return FooterData(
            footer_row_start_idx=footer_row_start_idx,
            data_start_row=data_start_row,
            data_end_row=data_end_row,
            total_pallets=final_pallets,
            leather_summary=leather_summary,
            weight_summary=final_weight_summary
        )
