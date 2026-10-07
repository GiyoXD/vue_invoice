import json
import logging
from pathlib import Path
from typing import Optional, Union, List
import datetime
from sqlalchemy.orm import Session
from core.database.models import ProcessedData
from core.database.repositories.invoice_repository import InvoiceRepository
from core.system_config import sys_config

logger = logging.getLogger(__name__)


class ShipmentService:
    """
    Coordinates shipment business logic including pending JSON file loading,
    acceptance, rejection, history management, and registry operations.
    """

    def __init__(self, db: Optional[Session] = None):
        self.db = db
        self.repository = InvoiceRepository(db) if db is not None else None

    def _validate_filename(self, filename: str) -> str:
        if not filename or ".." in filename or "/" in filename or "\\" in filename:
            raise ValueError(f"Invalid filename: {filename}")
        return filename.strip()

    def get_history(self) -> List[dict]:
        """Encapsulates the file scanning logic from run_log and temp_uploads_dir / 'processed'."""
        history = []

        run_log_dir = sys_config.run_log_dir
        if run_log_dir.exists():
            for f in run_log_dir.glob("*_metadata.json"):
                try:
                    with open(f, "r", encoding="utf-8") as json_file:
                        data = json.load(json_file)
                        history.append({
                            "filename": f.name,
                            "type": "run_log",
                            "timestamp": data.get("timestamp"),
                            "output_file": data.get("output_file"),
                            "status": data.get("status"),
                            "item_count": data.get("database_export", {}).get("summary", {}).get("item_count", 0),
                            "total_sqft": data.get("database_export", {}).get("summary", {}).get("total_sqft", 0),
                        })
                except Exception as e:
                    logger.debug(f"Failed to read run_log metadata {f.name}: {e}")
                    continue

        processed_dir = sys_config.temp_uploads_dir / "processed"
        if processed_dir.exists():
            for f in processed_dir.glob("*.json"):
                try:
                    if any(h["filename"] == f.name for h in history):
                        continue
                    stats = f.stat()
                    with open(f, "r", encoding="utf-8") as json_file:
                        data = json.load(json_file)
                        item_count = 0
                        if "multi_table" in data and data["multi_table"]:
                            item_count = sum(len(table) for table in data["multi_table"] if isinstance(table, list))

                        history.append({
                            "filename": f.name,
                            "type": "processed",
                            "timestamp": datetime.datetime.fromtimestamp(stats.st_mtime).isoformat(),
                            "output_file": f.name,
                            "status": "Ready",
                            "item_count": item_count,
                            "total_sqft": 0,
                        })
                except Exception as e:
                    logger.debug(f"Failed to read processed json {f.name}: {e}")
                    continue

        history.sort(key=lambda x: x["timestamp"] or "", reverse=True)
        return history

    def get_history_item(self, filename: str, source: str = "run_log") -> dict:
        """
        Validates filename using _validate_filename.
        If source == 'processed': reads from sys_config.temp_uploads_dir / 'processed' / filename.
        If source == 'accepted': fetches record from repository; if not record, raises LookupError('Database record not found'); returns record.data_payload.
        Else: reads from sys_config.run_log_dir / filename (fallback to sys_config.temp_uploads_dir / 'processed' / filename).
        Raises FileNotFoundError if file doesn't exist.
        Returns loaded json data.
        """
        clean_filename = self._validate_filename(filename)

        if source == "processed":
            file_path = sys_config.temp_uploads_dir / "processed" / clean_filename
        elif source == "accepted":
            if self.repository is None:
                raise RuntimeError("Database session required for accepted history items")
            record = self.repository.get_by_filename(clean_filename)
            if not record:
                raise LookupError("Database record not found")
            return record.data_payload
        else:
            file_path = sys_config.run_log_dir / clean_filename
            if not file_path.exists():
                fallback_path = sys_config.temp_uploads_dir / "processed" / clean_filename
                if fallback_path.exists():
                    file_path = fallback_path

        if not file_path.exists():
            raise FileNotFoundError(f"File not found: {clean_filename}")

        with open(file_path, "r", encoding="utf-8") as f:
            return json.load(f)

    def check_exists(self, filename: str) -> bool:
        """Check if an invoice has already been accepted in the database."""
        clean_filename = self._validate_filename(filename)
        if self.repository is None:
            raise RuntimeError("Database session required to check existence")
        record = self.repository.get_by_filename(clean_filename)
        return record is not None

    def accept(self, filename: str) -> ProcessedData:
        """
        Loads pending JSON from temp_uploads_dir / 'processed' / filename.
        Raises FileNotFoundError if file does not exist.
        Parses JSON (letting JSONDecodeError or OSError bubble).
        Calls InvoiceRepository.accept_processed_invoice(filename, data).
        Unlinks the pending JSON file on success.
        Returns the created/updated ProcessedData.
        """
        clean_filename = self._validate_filename(filename)
        if self.repository is None:
            raise RuntimeError("Database session required to accept invoice")
        pending_path = sys_config.temp_uploads_dir / "processed" / clean_filename
        if not pending_path.exists():
            raise FileNotFoundError(f"Pending file '{clean_filename}' not found")

        with open(pending_path, "r", encoding="utf-8") as f:
            data = json.load(f)

        record = self.repository.accept_processed_invoice(clean_filename, data)

        try:
            pending_path.unlink()
        except OSError as oe:
            logger.warning(f"Failed to unlink pending file {pending_path} after successful accept: {oe}")

        logger.info(f"Accepted shipment '{clean_filename}' ({record.item_count} items)")
        return record

    def reject(self, filename: str) -> bool:
        """
        Raises FileNotFoundError if pending file does not exist.
        Calls InvoiceRepository.delete_invoice(filename).
        Unlinks pending file.
        Returns True.
        """
        clean_filename = self._validate_filename(filename)
        if self.repository is None:
            raise RuntimeError("Database session required to reject invoice")
        pending_path = sys_config.temp_uploads_dir / "processed" / clean_filename
        if not pending_path.exists():
            raise FileNotFoundError(f"Pending file '{clean_filename}' not found")

        self.repository.delete_invoice(clean_filename)
        try:
            pending_path.unlink()
        except OSError as oe:
            logger.warning(f"Failed to unlink pending file {pending_path} after successful reject: {oe}")
        logger.info(f"Rejected shipment '{clean_filename}'")
        return True

    def list_registry(self, limit: int = 10) -> List[dict]:
        """Fetch recent records from repository and format as dictionary records."""
        if self.repository is None:
            raise RuntimeError("Database session required for listing registry")
        results = self.repository.list_recent(limit=limit)
        return [
            {
                "id": r.id,
                "filename": r.filename,
                "timestamp": r.timestamp.isoformat() if r.timestamp else None,
                "item_count": r.item_count,
                "total_sqft": r.total_sqft,
                "total_net": getattr(r, "total_net", 0.0),
                "total_pallets": getattr(r, "total_pallets", 0.0),
                "total_amount": r.total_amount,
            }
            for r in results
        ]

    def export_registry_csv(
        self,
        start_date: Optional[Union[str, datetime.date, datetime.datetime]] = None,
        end_date: Optional[Union[str, datetime.date, datetime.datetime]] = None,
    ) -> str:
        """Export registry CSV content for invoices within date range."""
        if self.repository is None:
            raise RuntimeError("Database session required to export registry CSV")
        csv_data = self.repository.generate_registry_csv(start_date=start_date, end_date=end_date)
        logger.info("Exported shipment registry CSV")
        return csv_data
