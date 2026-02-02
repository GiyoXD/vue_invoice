
import logging
import openpyxl
from typing import Dict, Any

logger = logging.getLogger(__name__)

class DeepSheetBuilder:
    """
    Builder responsible for creating and populating the hidden 'DeepSheet' 
    which stores metadata for stable referencing in Excel formulas.
    """

    @staticmethod
    def build(workbook: openpyxl.Workbook, invoice_data: Dict[str, Any]):
        """
        Injects a hidden 'DeepSheet' containing metadata.
        
        Structure:
        - Row 1: labels (col_invoice_no, col_ref_no, col_date)
        - Row 2: values (from invoice_data inputs or extracted tables)
        """
        SHEET_NAME = "DeepSheet"
        
        # Create or get sheet
        if SHEET_NAME in workbook.sheetnames:
            ws = workbook[SHEET_NAME]
        else:
            ws = workbook.create_sheet(SHEET_NAME)
        
        # Set to very hidden
        ws.sheet_state = 'veryHidden'
        
        # Headers (User requested labels)
        headers = ["col_invoice_no", "col_ref_no", "col_date"]
        
        # Metadata extraction logic (Try multiple sources)
        inv_no = ""
        ref_no = ""
        inv_date = ""
        
        # source 1: invoice_info (Created by some parsers or UI overrides)
        if 'invoice_info' in invoice_data:
            info = invoice_data['invoice_info']
            inv_no = info.get('col_inv_no', "") or info.get('inv_no', "")
            ref_no = info.get('col_inv_ref', "") or info.get('inv_ref', "")
            inv_date = info.get('col_inv_date', "") or info.get('inv_date', "")
        
        # source 2: processed_tables_multi['1'] (Standard parser output)
        # If values are still empty, try to find them here
        if not (inv_no and ref_no and inv_date):
            tables = invoice_data.get('processed_tables_multi', {})
            # Look at table '1' specifically as it's usually the main invoice table
            table_1 = tables.get('1', {})
            
            # Helper to get first non-empty value from a column list
            def get_first_val(key):
                vals = table_1.get(key, [])
                if isinstance(vals, list):
                    for v in vals:
                        if v: return str(v)
                return ""

            if not inv_no: inv_no = get_first_val('col_inv_no')
            if not ref_no: ref_no = get_first_val('col_inv_ref')
            if not inv_date: inv_date = get_first_val('col_inv_date')

        # Write Headers (Row 1)
        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_idx, value=header)

        # Write Values (Row 2)
        ws.cell(row=2, column=1, value=inv_no)
        ws.cell(row=2, column=2, value=ref_no)
        ws.cell(row=2, column=3, value=inv_date)
        
        logger.info(f"Injected veryHidden '{SHEET_NAME}' with metadata: Inv={inv_no}, Ref={ref_no}, Date={inv_date}")
