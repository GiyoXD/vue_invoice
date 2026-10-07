import io
import csv
import logging
import datetime
import math
from typing import Optional, List, Union, Any
from sqlalchemy import Numeric, Float
from sqlalchemy.orm import Session
from core.database.models import ProcessedData, InvoiceItem
from core.utils.clock import ICT, now as ict_now

logger = logging.getLogger(__name__)


def _to_float(val: Any) -> float:
    """Convert value to float, handling currency symbols, commas, and parentheses for negatives."""
    if val is None or isinstance(val, bool):
        return 0.0
    if isinstance(val, (int, float)):
        return 0.0 if math.isnan(val) or math.isinf(val) else float(val)
    try:
        clean = str(val).strip()
        is_negative = False
        if clean.startswith("(") and clean.endswith(")"):
            is_negative = True
            clean = clean[1:-1].strip()
        clean = clean.replace(",", "").replace("$", "").strip()
        if clean.startswith("-"):
            is_negative = True
            clean = clean[1:].strip()
        if not clean:
            return 0.0
        res = float(clean)
        if math.isnan(res) or math.isinf(res):
            return 0.0
        return -res if is_negative else res
    except (ValueError, TypeError):
        return 0.0


def _to_str(val: Any) -> str:
    """Convert value to stripped string, returning empty string for None, NaN, null."""
    if val is None:
        return ""
    if isinstance(val, float) and (math.isnan(val) or math.isinf(val)):
        return ""
    s = str(val).strip()
    if s.lower() in ("none", "nan", "null"):
        return ""
    return s


def _to_datetime(val: Any, default_now: bool = False) -> Optional[datetime.datetime]:
    """Convert input to timezone-aware (ICT) datetime, or None / ict_now()."""
    if val is None:
        return ict_now() if default_now else None
    if isinstance(val, datetime.datetime):
        return val.astimezone(ICT) if val.tzinfo is not None else val.replace(tzinfo=ICT)
    if isinstance(val, datetime.date):
        return datetime.datetime.combine(val, datetime.time.min, tzinfo=ICT)
    try:
        s = str(val).strip()
        if not s or s.lower() in ("none", "nan", "null"):
            return ict_now() if default_now else None
        dt = datetime.datetime.fromisoformat(s)
        return dt.astimezone(ICT) if dt.tzinfo is not None else dt.replace(tzinfo=ICT)
    except Exception:
        return ict_now() if default_now else None


def _parse_filter_date(
    val: Optional[Union[str, datetime.date, datetime.datetime]],
    is_end_date: bool = False,
    param_name: str = "date",
) -> Optional[datetime.datetime]:
    """
    Parse a filter date parameter into a timezone-aware (ICT) datetime.
    For end_date with date-only input, adds 1 day for inclusive filtering (< end_dt).
    Raises ValueError on invalid date format.
    """
    if val is None:
        return None

    if isinstance(val, str):
        s = val.strip()
        if not s:
            return None
        try:
            dt = datetime.datetime.fromisoformat(s)
        except (ValueError, TypeError) as e:
            raise ValueError(f"Invalid {param_name} format '{val}': {e}")
        if is_end_date and len(s) <= 10:
            dt += datetime.timedelta(days=1)
        return dt.astimezone(ICT) if dt.tzinfo is not None else dt.replace(tzinfo=ICT)
    elif isinstance(val, datetime.datetime):
        return val.astimezone(ICT) if val.tzinfo is not None else val.replace(tzinfo=ICT)
    elif isinstance(val, datetime.date):
        dt = datetime.datetime.combine(val, datetime.time.min, tzinfo=ICT)
        if is_end_date:
            dt += datetime.timedelta(days=1)
        return dt
    return None


class InvoiceRepository:
    """
    Encapsulates all database query and storage operations for ProcessedData and InvoiceItems.
    """

    def __init__(self, db: Session):
        self.db = db

    def get_by_filename(self, filename: str) -> Optional[ProcessedData]:
        """Fetch a single ProcessedData record by filename."""
        clean_filename = str(filename).strip() if filename is not None else ""
        if not clean_filename:
            return None
        return self.db.query(ProcessedData).filter(ProcessedData.filename == clean_filename).first()

    def list_recent(self, limit: int = 10) -> List[ProcessedData]:
        """Fetch recent ProcessedData records ordered by timestamp descending, id descending."""
        return (
            self.db.query(ProcessedData)
            .order_by(ProcessedData.timestamp.desc(), ProcessedData.id.desc())
            .limit(limit)
            .all()
        )

    def delete_invoice(self, filename: str) -> bool:
        """
        Delete InvoiceItem records where invoice_id == filename
        and ProcessedData record where filename == filename.
        Use transaction with commit and rollback on error.
        Return True if existed and deleted, False otherwise.
        """
        clean_filename = str(filename).strip() if filename is not None else ""
        if not clean_filename:
            return False
        try:
            processed = self.get_by_filename(clean_filename)
            items_deleted = (
                self.db.query(InvoiceItem)
                .filter(InvoiceItem.invoice_id == clean_filename)
                .delete(synchronize_session=False)
            )
            existed = (processed is not None) or (items_deleted > 0)
            if processed:
                self.db.delete(processed)
            self.db.commit()
            return existed
        except Exception as e:
            self.db.rollback()
            logger.error(f"Failed to delete invoice {clean_filename}: {e}")
            raise

    def save_processed_data(
        self,
        filename: str,
        item_count: int,
        total_sqft: float,
        total_net: float,
        total_amount: float,
        total_pallets: float,
        data_payload: dict,
        status: Optional[str] = None,
        commit: bool = True,
    ) -> ProcessedData:
        """
        If ProcessedData exists with filename, update its fields and timestamp to ict_now().
        Else insert new ProcessedData. Commit and refresh (or flush if commit=False). Return record.
        """
        clean_filename = str(filename).strip() if filename is not None else ""
        if not clean_filename:
            raise ValueError("Invalid filename")

        try:
            item_count = int(_to_float(item_count))
            total_sqft = _to_float(total_sqft)
            total_net = _to_float(total_net)
            total_amount = _to_float(total_amount)
            total_pallets = _to_float(total_pallets)

            record = self.get_by_filename(clean_filename)
            now_dt = ict_now()
            if record:
                record.item_count = item_count
                record.total_sqft = total_sqft
                record.total_net = total_net
                record.total_amount = total_amount
                record.total_pallets = total_pallets
                record.data_payload = data_payload
                if status is not None:
                    record.status = status
                record.timestamp = now_dt
            else:
                record = ProcessedData(
                    filename=clean_filename,
                    item_count=item_count,
                    total_sqft=total_sqft,
                    total_net=total_net,
                    total_amount=total_amount,
                    total_pallets=total_pallets,
                    data_payload=data_payload,
                    status=status or "Accepted",
                    timestamp=now_dt,
                )
                self.db.add(record)
            if commit:
                self.db.commit()
                self.db.refresh(record)
            else:
                self.db.flush()
            return record
        except Exception as e:
            self.db.rollback()
            logger.error(f"Failed to save processed data for {clean_filename}: {e}")
            raise

    def _row_to_invoice_item(
        self, invoice_id: str, row: Union[dict, list, tuple], is_adjustment: int = 0
    ) -> InvoiceItem:
        """
        Map a dictionary row or list/tuple adjustment directly to an InvoiceItem instance.
        """
        if isinstance(row, (list, tuple)) and len(row) >= 2:
            row = {
                "col_desc": _to_str(row[0]),
                "col_amount": _to_float(row[1]),
            }
            is_adj = 1
        elif not isinstance(row, dict):
            row = {}
            is_adj = 0
        else:
            def _parse_adj(v) -> int:
                if isinstance(v, str):
                    return 1 if v.strip().lower() in ("1", "true", "yes") else 0
                return 1 if bool(v) else 0

            is_adj = _parse_adj(is_adjustment)
            if not is_adj and "is_adjustment" in row:
                is_adj = _parse_adj(row.get("is_adjustment"))

        desc_raw = row.get("col_desc")
        if (desc_raw is None or (isinstance(desc_raw, str) and not desc_raw.strip())) and is_adj:
            desc_raw = row.get("description")

        amount_raw = row.get("col_amount")
        if (amount_raw is None or (isinstance(amount_raw, str) and not amount_raw.strip())) and is_adj:
            amount_raw = row.get("amount")

        cbm_raw = row.get("col_cbm_raw")
        if cbm_raw is None or (isinstance(cbm_raw, str) and not cbm_raw.strip()):
            cbm_raw = row.get("col_cbm")

        return InvoiceItem(
            invoice_id=invoice_id,
            is_adjustment=is_adj,
            timestamp=_to_datetime(row.get("timestamp"), default_now=True),
            col_dc=_to_str(row.get("col_dc")),
            col_po=_to_str(row.get("col_po")),
            col_production_order_no=_to_str(row.get("col_production_order_no")),
            col_production_date=_to_str(row.get("col_production_date")),
            col_line_no=_to_str(row.get("col_line_no")),
            col_direction=_to_str(row.get("col_direction")),
            col_item=_to_str(row.get("col_item")),
            col_reference_code=_to_str(row.get("col_reference_code")),
            col_desc=_to_str(desc_raw),
            col_level=_to_str(row.get("col_level")),
            col_grade=_to_str(row.get("col_grade")),
            col_qty_pcs=_to_float(row.get("col_qty_pcs")),
            col_qty_sf=_to_float(row.get("col_qty_sf")),
            col_pallet_count=_to_float(row.get("col_pallet_count")),
            col_pallet_no=_to_str(row.get("col_pallet_no")),
            col_net=_to_float(row.get("col_net")),
            col_gross=_to_float(row.get("col_gross")),
            col_cbm_raw=_to_str(cbm_raw),
            col_hs_code=_to_str(row.get("col_hs_code")),
            col_unit_price=_to_float(row.get("col_unit_price")),
            col_amount=_to_float(amount_raw),
        )

    def save_invoice_items(
        self,
        filename: str,
        items: List[dict],
        delete_existing: bool = True,
        commit: bool = True,
    ) -> int:
        """
        If delete_existing, delete InvoiceItem where invoice_id == filename.
        Add each item dict as an InvoiceItem record.
        Add all, commit (or flush if commit=False), return count of items added.
        """
        clean_filename = str(filename).strip() if filename is not None else ""
        if not clean_filename:
            raise ValueError("Invalid filename")

        if not isinstance(items, list):
            items = []

        try:
            if delete_existing:
                self.db.query(InvoiceItem).filter(InvoiceItem.invoice_id == clean_filename).delete(
                    synchronize_session="fetch"
                )

            items_to_add = [
                self._row_to_invoice_item(
                    clean_filename,
                    row,
                    is_adjustment=row.get("is_adjustment", 0) if isinstance(row, dict) else 0,
                )
                for row in items
                if isinstance(row, dict)
            ]

            if items_to_add:
                self.db.add_all(items_to_add)
            if commit:
                self.db.commit()
            else:
                self.db.flush()
            return len(items_to_add)
        except Exception as e:
            self.db.rollback()
            logger.error(f"Failed to save invoice items for {clean_filename}: {e}")
            raise

    def accept_processed_invoice(self, filename: str, data: dict) -> ProcessedData:
        """
        Validate, map, and accept a processed invoice payload into ProcessedData and InvoiceItems.
        """
        clean_filename = str(filename).strip() if filename is not None else ""
        if not clean_filename:
            raise ValueError("Invalid filename")

        footer_data = data.get("footer_data") if isinstance(data, dict) else None
        grand_total = footer_data.get("grand_total") if isinstance(footer_data, dict) else None
        if not grand_total or not isinstance(grand_total, dict):
            raise ValueError(f"Missing footer_data or grand_total in invoice data for '{clean_filename}'")

        try:
            total_pallets = _to_float(grand_total.get("col_pallet_count", 0))
            total_sqft = _to_float(grand_total.get("col_qty_sf", 0))
            total_net = _to_float(grand_total.get("col_net", 0))
            gt_amount = grand_total.get("col_amount")
            if gt_amount is None or (isinstance(gt_amount, str) and not gt_amount.strip()):
                gt_amount = grand_total.get("amount", 0)
            total_amount = _to_float(gt_amount)

            self.db.query(InvoiceItem).filter(InvoiceItem.invoice_id == clean_filename).delete(
                synchronize_session="fetch"
            )

            items_to_add: List[InvoiceItem] = []
            for table in data.get("multi_table") or []:
                if isinstance(table, list):
                    for row in table:
                        if isinstance(row, dict):
                            items_to_add.append(self._row_to_invoice_item(clean_filename, row, is_adjustment=0))

            item_count = len(items_to_add)
            if item_count == 0 and "item_count" in grand_total:
                item_count = int(_to_float(grand_total.get("item_count", 0)))

            for adj in data.get("price_adjustment") or []:
                if isinstance(adj, (dict, list, tuple)):
                    items_to_add.append(self._row_to_invoice_item(clean_filename, adj, is_adjustment=1))

            if items_to_add:
                self.db.add_all(items_to_add)

            record = self.save_processed_data(
                filename=clean_filename,
                item_count=item_count,
                total_sqft=total_sqft,
                total_net=total_net,
                total_amount=total_amount,
                total_pallets=total_pallets,
                data_payload=data,
                status="Accepted",
                commit=False,
            )

            self.db.commit()
            self.db.refresh(record)
            return record
        except Exception as e:
            self.db.rollback()
            logger.error(f"Failed to accept processed invoice for {clean_filename}: {e}")
            raise

    def query_items(
        self,
        start_date: Optional[Union[str, datetime.date, datetime.datetime]] = None,
        end_date: Optional[Union[str, datetime.date, datetime.datetime]] = None,
    ) -> List[InvoiceItem]:
        """
        Query InvoiceItem with optional start_date and end_date filtering.
        Order by timestamp.desc(), invoice_id.
        Raises ValueError if invalid date strings are provided.
        """
        query = self.db.query(InvoiceItem)
        start_dt = _parse_filter_date(start_date, is_end_date=False, param_name="start_date")
        if start_dt is not None:
            query = query.filter(InvoiceItem.timestamp >= start_dt)

        end_dt = _parse_filter_date(end_date, is_end_date=True, param_name="end_date")
        if end_dt is not None:
            query = query.filter(InvoiceItem.timestamp < end_dt)

        return query.order_by(InvoiceItem.timestamp.desc(), InvoiceItem.invoice_id).all()

    def generate_registry_csv(
        self,
        start_date: Optional[Union[str, datetime.date, datetime.datetime]] = None,
        end_date: Optional[Union[str, datetime.date, datetime.datetime]] = None,
    ) -> str:
        """
        Generate UTF-8 BOM encoded CSV content of InvoiceItems within date range.
        Dynamically introspects InvoiceItem columns and appends a numeric summary row.
        """
        results = self.query_items(start_date=start_date, end_date=end_date)
        columns = list(InvoiceItem.__table__.columns)

        header_titles = {
            "id": "ID",
            "invoice_id": "Invoice ID",
            "timestamp": "Timestamp",
            "col_dc": "DC",
            "col_po": "PO",
            "col_production_order_no": "Production Order No",
            "col_production_date": "Production Date",
            "col_line_no": "Line No",
            "col_direction": "Direction",
            "col_item": "Item",
            "col_reference_code": "Reference Code",
            "col_desc": "Description",
            "col_level": "Level",
            "col_grade": "Grade",
            "col_qty_pcs": "Qty (PCS)",
            "col_qty_sf": "Qty (SF)",
            "col_pallet_count": "Pallet Count",
            "col_pallet_no": "Pallet No",
            "col_net": "Net Weight",
            "col_gross": "Gross Weight",
            "col_cbm_raw": "CBM Raw",
            "col_hs_code": "HS Code",
            "col_unit_price": "Unit Price",
            "col_amount": "Amount",
            "is_adjustment": "Is Adjustment",
        }

        headers = [
            header_titles.get(col.name, col.name.removeprefix("col_").replace("_", " ").title())
            for col in columns
        ]

        output = io.StringIO()
        writer = csv.writer(output)
        writer.writerow(headers)

        summable_cols = {
            "col_qty_pcs",
            "col_qty_sf",
            "col_pallet_count",
            "col_net",
            "col_gross",
            "col_amount",
        }
        numeric_totals = {col_name: 0.0 for col_name in summable_cols}

        for row in results:
            row_vals = []
            for col in columns:
                val = getattr(row, col.name)
                if col.name == "invoice_id":
                    row_vals.append((str(val) if val is not None else "").replace(".json", ""))
                elif col.name == "is_adjustment":
                    row_vals.append("Yes" if val else "No")
                elif isinstance(val, (datetime.datetime, datetime.date)):
                    row_vals.append(val.isoformat())
                elif isinstance(col.type, (Float, Numeric)):
                    num_val = _to_float(val)
                    if col.name in summable_cols:
                        numeric_totals[col.name] += num_val
                    row_vals.append(val if val is not None else "")
                elif val is None:
                    row_vals.append("")
                else:
                    row_vals.append(str(val))
            writer.writerow(row_vals)

        summary_row = [""] * len(columns)
        for idx, col in enumerate(columns):
            if col.name == "invoice_id":
                summary_row[idx] = "TOTAL"
            elif col.name in summable_cols:
                summary_row[idx] = round(numeric_totals[col.name], 2)

        writer.writerow([])
        writer.writerow(summary_row)

        output.seek(0)
        return "\ufeff" + output.getvalue()
