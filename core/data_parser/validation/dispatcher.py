import logging
from typing import List, Dict, Any, Optional
from .presence import validate_table_data_presence
from .weight import validate_weight_integrity
from .cbm import validate_cbm_pcs_proportion
from .pallet import verify_pallet_integrity


def validate_data(
    data_rows: List[Dict[str, Any]], 
    table_id_str: str, 
    column_mapping: Dict[str, str], 
    monitor: Optional[Any] = None,
    phase: str = 'presence',
    ignore_tare_warning: bool = False,
    ignore_cbm_warning: bool = False
):
    """
    Unified entry point for all table-level data validation.
    
    Args:
        data_rows: The list of row dictionaries to validate.
        table_id_str: Human readable ID (e.g. "Table 1") for error messages.
        column_mapping: The canonical-to-Excel-letter mapping.
        monitor: Optional PipelineMonitor for logging errors.
        phase: 'presence' (check for at least one value per required col) 
               or 'integrity' (check weight integrity and tare matching)
               or 'cbm_proportion' (check CBM and PCS proportions).
    """
    if phase == 'presence':
        validate_table_data_presence(data_rows, table_id_str, column_mapping, monitor=monitor)
    elif phase == 'integrity':
        validate_weight_integrity(
            data_rows, 
            table_id_str=table_id_str, 
            monitor=monitor, 
            ignore_tare_warning=ignore_tare_warning
        )
        verify_pallet_integrity(
            data_rows,
            table_id_str=table_id_str,
            monitor=monitor
        )
    elif phase == 'cbm_proportion':
        validate_cbm_pcs_proportion(
            data_rows, 
            table_id_str, 
            column_mapping, 
            monitor=monitor, 
            ignore_cbm_warning=ignore_cbm_warning
        )
    else:
        raise ValueError(f"Unknown validation phase requested: {phase}. Supported: 'presence', 'integrity', 'cbm_proportion'.")
