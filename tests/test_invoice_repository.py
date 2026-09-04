import io
import csv
import datetime
import pytest
from core.database.db_manager import InvoiceRepository, ProcessedData, InvoiceItem
from core.database.repositories.invoice_repository import _to_float, _to_str
from core.utils.clock import ICT, now as ict_now


def test_save_processed_data_new(db):
    repo = InvoiceRepository(db)
    filename = "test_invoice_new.json"
    payload = {"invoice_number": "INV-001", "items": [{"id": 1}]}

    record = repo.save_processed_data(
        filename=filename,
        item_count=10,
        total_sqft=2500.5,
        total_net=1200.0,
        total_amount=5432.1,
        total_pallets=4.0,
        data_payload=payload,
    )

    assert record.id is not None
    assert record.filename == filename
    assert record.item_count == 10
    assert record.total_sqft == 2500.5
    assert record.total_net == 1200.0
    assert record.total_amount == 5432.1
    assert record.total_pallets == 4.0
    assert record.data_payload == payload
    assert record.status == "Accepted"
    assert record.timestamp is not None

    # Fetch from repository
    fetched = repo.get_by_filename(filename)
    assert fetched is not None
    assert fetched.id == record.id
    assert fetched.filename == filename


def test_save_processed_data_update(db):
    repo = InvoiceRepository(db)
    filename = "test_invoice_update.json"
    initial_payload = {"version": 1}

    initial_record = repo.save_processed_data(
        filename=filename,
        item_count=5,
        total_sqft=1000.0,
        total_net=500.0,
        total_amount=1500.0,
        total_pallets=2.0,
        data_payload=initial_payload,
    )
    initial_id = initial_record.id
    initial_timestamp = initial_record.timestamp

    # Update the record
    updated_payload = {"version": 2, "updated": True}
    updated_record = repo.save_processed_data(
        filename=filename,
        item_count=8,
        total_sqft=1600.0,
        total_net=800.0,
        total_amount=2400.0,
        total_pallets=3.0,
        data_payload=updated_payload,
        status="Reviewed",
    )

    assert updated_record.id == initial_id
    assert updated_record.item_count == 8
    assert updated_record.total_sqft == 1600.0
    assert updated_record.total_net == 800.0
    assert updated_record.total_amount == 2400.0
    assert updated_record.total_pallets == 3.0
    assert updated_record.data_payload == updated_payload
    assert updated_record.status == "Reviewed"
    assert updated_record.timestamp >= initial_timestamp


def test_get_by_filename(db):
    repo = InvoiceRepository(db)
    assert repo.get_by_filename("non_existent_file.json") is None

    filename = "lookup_target.json"
    repo.save_processed_data(
        filename=filename,
        item_count=1,
        total_sqft=10.0,
        total_net=5.0,
        total_amount=100.0,
        total_pallets=1.0,
        data_payload={"foo": "bar"},
    )

    found = repo.get_by_filename(filename)
    assert found is not None
    assert found.filename == filename


def test_list_recent(db):
    repo = InvoiceRepository(db)

    # Clean existing ProcessedData if any
    db.query(ProcessedData).delete()
    db.commit()

    # Create 5 records with distinct timestamps
    base_time = datetime.datetime(2026, 9, 1, 10, 0, 0, tzinfo=ICT)
    for i in range(5):
        record = ProcessedData(
            filename=f"recent_inv_{i}.json",
            item_count=i + 1,
            total_sqft=100.0 * (i + 1),
            total_net=50.0 * (i + 1),
            total_amount=200.0 * (i + 1),
            total_pallets=1.0 * (i + 1),
            data_payload={},
            timestamp=base_time + datetime.timedelta(hours=i),
        )
        db.add(record)
    db.commit()

    # Test limit=3
    recent_3 = repo.list_recent(limit=3)
    assert len(recent_3) == 3
    assert recent_3[0].filename == "recent_inv_4.json"
    assert recent_3[1].filename == "recent_inv_3.json"
    assert recent_3[2].filename == "recent_inv_2.json"

    # Test default limit=10
    recent_all = repo.list_recent()
    assert len(recent_all) == 5
    assert recent_all[0].filename == "recent_inv_4.json"
    assert recent_all[-1].filename == "recent_inv_0.json"


def test_save_invoice_items_and_delete_existing(db):
    repo = InvoiceRepository(db)
    filename = "invoice_items_test.json"

    items_batch_1 = [
        {
            "col_dc": "DC1",
            "col_po": "PO123",
            "col_production_order_no": "PROD999",
            "col_production_date": "2026-09-01",
            "col_line_no": "01",
            "col_direction": "NORTH",
            "col_item": "ITEM-001",
            "col_reference_code": "REF-A",
            "col_desc": "First item description",
            "col_level": "Level 1",
            "col_grade": "Grade A",
            "col_qty_pcs": "10",
            "col_qty_sf": "25.5",
            "col_pallet_count": "1.0",
            "col_pallet_no": "PAL-1",
            "col_net": "50.5",
            "col_gross": "55.2",
            "col_cbm_raw": "0.12",
            "col_hs_code": "1234.56",
            "col_unit_price": "$15.00",
            "col_amount": "$150.00",
            "is_adjustment": 0,
        },
        {
            "col_item": "ITEM-002",
            "col_desc": "Second item",
            "col_qty_pcs": 20,
            "col_qty_sf": 50.0,
            "col_pallet_count": 2.0,
            "col_net": 100.0,
            "col_gross": 110.0,
            "col_cbm": "0.24",  # test fallback to col_cbm
            "col_unit_price": 20.0,
            "col_amount": 400.0,
            "is_adjustment": "1",
        },
    ]

    count = repo.save_invoice_items(filename, items_batch_1, delete_existing=True)
    assert count == 2

    stored = db.query(InvoiceItem).filter(InvoiceItem.invoice_id == filename).all()
    assert len(stored) == 2

    item1 = next(it for it in stored if it.col_item == "ITEM-001")
    assert item1.col_dc == "DC1"
    assert item1.col_po == "PO123"
    assert item1.col_qty_pcs == 10.0
    assert item1.col_qty_sf == 25.5
    assert item1.col_unit_price == 15.0
    assert item1.col_amount == 150.0
    assert item1.col_cbm_raw == "0.12"
    assert item1.is_adjustment == 0

    item2 = next(it for it in stored if it.col_item == "ITEM-002")
    assert item2.col_cbm_raw == "0.24"
    assert item2.is_adjustment == 1

    # Overwrite with delete_existing=True
    items_batch_2 = [
        {
            "col_item": "ITEM-003",
            "col_desc": "Replacement item",
            "col_amount": 500.0,
        }
    ]
    count_replace = repo.save_invoice_items(filename, items_batch_2, delete_existing=True)
    assert count_replace == 1

    stored_after_replace = db.query(InvoiceItem).filter(InvoiceItem.invoice_id == filename).all()
    assert len(stored_after_replace) == 1
    assert stored_after_replace[0].col_item == "ITEM-003"

    # Append with delete_existing=False
    items_batch_3 = [
        {
            "col_item": "ITEM-004",
            "col_desc": "Appended item",
            "col_amount": 600.0,
        }
    ]
    count_append = repo.save_invoice_items(filename, items_batch_3, delete_existing=False)
    assert count_append == 1

    stored_after_append = db.query(InvoiceItem).filter(InvoiceItem.invoice_id == filename).all()
    assert len(stored_after_append) == 2
    item_names = {it.col_item for it in stored_after_append}
    assert item_names == {"ITEM-003", "ITEM-004"}


def test_query_items_with_date_filtering(db):
    repo = InvoiceRepository(db)

    # Clean existing items
    db.query(InvoiceItem).delete()
    db.commit()

    # Seed items across different timestamps
    records_to_seed = [
        {
            "col_item": "AUG-01",
            "timestamp": "2026-08-01T10:00:00",
        },
        {
            "col_item": "AUG-15",
            "timestamp": "2026-08-15T12:00:00",
        },
        {
            "col_item": "SEP-01",
            "timestamp": "2026-09-01T08:00:00",
        },
        {
            "col_item": "SEP-02",
            "timestamp": "2026-09-02T15:30:00",
        },
        {
            "col_item": "SEP-05",
            "timestamp": "2026-09-05T09:00:00",
        },
    ]
    repo.save_invoice_items("query_test_invoice.json", records_to_seed, delete_existing=True)

    # 1. No filters: returns all 5 ordered timestamp desc
    all_items = repo.query_items()
    assert len(all_items) == 5
    assert [it.col_item for it in all_items] == ["SEP-05", "SEP-02", "SEP-01", "AUG-15", "AUG-01"]

    # 2. start_date filter (YYYY-MM-DD)
    sep_items = repo.query_items(start_date="2026-09-01")
    assert len(sep_items) == 3
    assert [it.col_item for it in sep_items] == ["SEP-05", "SEP-02", "SEP-01"]

    # 3. end_date filter (YYYY-MM-DD -> inclusive of that day via +1 day < end_dt)
    aug_items = repo.query_items(end_date="2026-08-31")
    assert len(aug_items) == 2
    assert [it.col_item for it in aug_items] == ["AUG-15", "AUG-01"]

    # 4. Both start_date and end_date
    mid_items = repo.query_items(start_date="2026-08-15", end_date="2026-09-02")
    assert len(mid_items) == 3
    assert [it.col_item for it in mid_items] == ["SEP-02", "SEP-01", "AUG-15"]

    # 5. Exact ISO datetime end_date (len > 10)
    cutoff_items = repo.query_items(end_date="2026-09-02T12:00:00")
    assert len(cutoff_items) == 3
    assert [it.col_item for it in cutoff_items] == ["SEP-01", "AUG-15", "AUG-01"]


def test_delete_invoice(db):
    repo = InvoiceRepository(db)

    del_file = "to_delete.json"
    keep_file = "to_keep.json"

    # Create ProcessedData and InvoiceItems for to_delete
    repo.save_processed_data(
        filename=del_file,
        item_count=2,
        total_sqft=100.0,
        total_net=50.0,
        total_amount=200.0,
        total_pallets=1.0,
        data_payload={},
    )
    repo.save_invoice_items(
        del_file,
        [{"col_item": "DEL-1"}, {"col_item": "DEL-2"}],
        delete_existing=True,
    )

    # Create ProcessedData and InvoiceItems for to_keep
    repo.save_processed_data(
        filename=keep_file,
        item_count=1,
        total_sqft=50.0,
        total_net=25.0,
        total_amount=100.0,
        total_pallets=1.0,
        data_payload={},
    )
    repo.save_invoice_items(
        keep_file,
        [{"col_item": "KEEP-1"}],
        delete_existing=True,
    )

    # Verify both exist before deletion
    assert repo.get_by_filename(del_file) is not None
    assert len(db.query(InvoiceItem).filter(InvoiceItem.invoice_id == del_file).all()) == 2
    assert repo.get_by_filename(keep_file) is not None
    assert len(db.query(InvoiceItem).filter(InvoiceItem.invoice_id == keep_file).all()) == 1

    # Execute cascade deletion
    deleted = repo.delete_invoice(del_file)
    assert deleted is True

    # Verify del_file is completely removed
    assert repo.get_by_filename(del_file) is None
    assert len(db.query(InvoiceItem).filter(InvoiceItem.invoice_id == del_file).all()) == 0

    # Verify keep_file is completely preserved
    assert repo.get_by_filename(keep_file) is not None
    assert len(db.query(InvoiceItem).filter(InvoiceItem.invoice_id == keep_file).all()) == 1

    # Deleting again should return False
    assert repo.delete_invoice(del_file) is False
    assert repo.delete_invoice("non_existent.json") is False


def test_transaction_rollback_on_error(db, monkeypatch):
    repo = InvoiceRepository(db)

    # Mock commit to simulate failure
    def mock_commit():
        raise RuntimeError("Simulated DB commit error")

    monkeypatch.setattr(db, "commit", mock_commit)

    with pytest.raises(RuntimeError, match="Simulated DB commit error"):
        repo.save_processed_data(
            filename="error_test.json",
            item_count=1,
            total_sqft=10.0,
            total_net=5.0,
            total_amount=10.0,
            total_pallets=1.0,
            data_payload={},
        )

    with pytest.raises(RuntimeError, match="Simulated DB commit error"):
        repo.delete_invoice("error_test.json")

    with pytest.raises(RuntimeError, match="Simulated DB commit error"):
        repo.save_invoice_items("error_test.json", [{"col_item": "ERR-1"}])


def test_to_float():
    assert _to_float(None) == 0.0
    assert _to_float(True) == 0.0
    assert _to_float(False) == 0.0
    assert _to_float(float("nan")) == 0.0
    assert _to_float(float("inf")) == 0.0
    assert _to_float(float("-inf")) == 0.0
    assert _to_float("nan") == 0.0
    assert _to_float("inf") == 0.0
    assert _to_float("-inf") == 0.0
    assert _to_float("abc") == 0.0
    assert _to_float("") == 0.0
    assert _to_float(10) == 10.0
    assert _to_float(-10) == -10.0
    assert _to_float("10.50") == 10.50
    assert _to_float("-10.50") == -10.50
    assert _to_float("- $10.50") == -10.50
    assert _to_float("$ -10.50") == -10.50
    assert _to_float("-$10.50") == -10.50
    assert _to_float("$-10.50") == -10.50
    assert _to_float("(100.00)") == -100.00
    assert _to_float(" ( 100.00 ) ") == -100.00
    assert _to_float("( $100.00 )") == -100.00
    assert _to_float("($100.00)") == -100.00
    assert _to_float("$1,234.56") == 1234.56


def test_to_str():
    assert _to_str(None) == ""
    assert _to_str(float("nan")) == ""
    assert _to_str(float("inf")) == ""
    assert _to_str(float("-inf")) == ""
    assert _to_str("hello") == "hello"
    assert _to_str("  hello  ") == "hello"
    assert _to_str(123) == "123"
    assert _to_str("none") == ""
    assert _to_str("None") == ""
    assert _to_str("NONE") == ""
    assert _to_str("nan") == ""
    assert _to_str("NaN") == ""
    assert _to_str("null") == ""
    assert _to_str("Null") == ""
    assert _to_str("NULL") == ""
    assert _to_str("  none  ") == ""
    assert _to_str("  nan  ") == ""
    assert _to_str("  null  ") == ""


def test_col_cbm_raw_fallback_when_none(db):
    repo = InvoiceRepository(db)
    filename = "test_cbm_fallback.json"
    items = [
        {"col_item": "ITEM-1", "col_cbm_raw": None, "col_cbm": "0.75"},
        {"col_item": "ITEM-2", "col_cbm_raw": "", "col_cbm": "0.85"},
        {"col_item": "ITEM-3", "col_cbm_raw": "0.95", "col_cbm": "0.10"},
        {"col_item": "ITEM-4", "col_cbm_raw": None, "col_cbm": None},
    ]
    repo.save_invoice_items(filename, items, delete_existing=True)

    stored = db.query(InvoiceItem).filter(InvoiceItem.invoice_id == filename).all()
    item_map = {it.col_item: it.col_cbm_raw for it in stored}

    assert item_map["ITEM-1"] == "0.75"
    assert item_map["ITEM-2"] == "0.85"
    assert item_map["ITEM-3"] == "0.95"
    assert item_map["ITEM-4"] == ""


def test_delete_invoice_empty_or_none(db):
    repo = InvoiceRepository(db)
    assert repo.delete_invoice("") is False
    assert repo.delete_invoice(None) is False
    assert repo.delete_invoice("   ") is False


def test_multi_step_atomic_transaction(db):
    repo = InvoiceRepository(db)
    filename = "atomic_test.json"

    # Multi-step atomic transaction with commit=False followed by single db.commit()
    record = repo.save_processed_data(
        filename=filename,
        item_count=2,
        total_sqft=100.0,
        total_net=50.0,
        total_amount=200.0,
        total_pallets=1.0,
        data_payload={"phase": "staged"},
        commit=False,
    )
    assert record.id is not None  # flush assigns id

    count = repo.save_invoice_items(
        filename=filename,
        items=[{"col_item": "A1"}, {"col_item": "B1"}],
        commit=False,
    )
    assert count == 2

    # Now commit atomically
    db.commit()

    saved_proc = repo.get_by_filename(filename)
    assert saved_proc is not None
    assert saved_proc.item_count == 2

    items = db.query(InvoiceItem).filter(InvoiceItem.invoice_id == filename).all()
    assert len(items) == 2


def test_list_recent_tie_breaker(db):
    repo = InvoiceRepository(db)
    db.query(ProcessedData).delete()
    db.commit()

    fixed_time = datetime.datetime(2026, 9, 1, 12, 0, 0, tzinfo=ICT)
    rec1 = ProcessedData(
        filename="tie_1.json",
        item_count=1,
        total_sqft=1.0,
        total_net=1.0,
        total_amount=1.0,
        total_pallets=1.0,
        data_payload={},
        timestamp=fixed_time,
    )
    rec2 = ProcessedData(
        filename="tie_2.json",
        item_count=2,
        total_sqft=2.0,
        total_net=2.0,
        total_amount=2.0,
        total_pallets=2.0,
        data_payload={},
        timestamp=fixed_time,
    )
    db.add(rec1)
    db.add(rec2)
    db.commit()

    recent = repo.list_recent()
    assert len(recent) == 2
    assert recent[0].id > recent[1].id
    assert recent[0].filename == "tie_2.json"


def test_row_to_invoice_item(db):
    repo = InvoiceRepository(db)
    invoice_id = "test_row_conv.json"

    # 1. Dynamic reflection and type conversion
    row_data = {
        "col_item": "ITEM-100",
        "col_desc": "Standard description",
        "col_po": "PO-999",
        "col_qty_pcs": "50",
        "col_qty_sf": "125.75",
        "col_pallet_count": "2.0",
        "col_net": "$500.25",
        "col_gross": "550.00",
        "col_unit_price": "$10.50",
        "col_amount": "$525.00",
        "unrecognized_field_1": "should_be_ignored",
        "extra_nested_data": {"some": "data"},
    }
    item = repo._row_to_invoice_item(invoice_id, row_data, is_adjustment=0)

    assert isinstance(item, InvoiceItem)
    assert item.invoice_id == invoice_id
    assert item.col_item == "ITEM-100"
    assert item.col_desc == "Standard description"
    assert item.col_po == "PO-999"
    assert item.col_qty_pcs == 50.0
    assert item.col_qty_sf == 125.75
    assert item.col_pallet_count == 2.0
    assert item.col_net == 500.25
    assert item.col_gross == 550.0
    assert item.col_unit_price == 10.5
    assert item.col_amount == 525.0
    assert item.is_adjustment == 0
    assert not hasattr(item, "unrecognized_field_1") or getattr(item, "unrecognized_field_1", None) is None

    # 2. col_cbm_raw fallback to col_cbm
    row_cbm_fallback = {
        "col_item": "ITEM-CBM-1",
        "col_cbm": "1.45",
    }
    item_cbm = repo._row_to_invoice_item(invoice_id, row_cbm_fallback, is_adjustment=0)
    assert item_cbm.col_cbm_raw == "1.45"

    row_cbm_explicit = {
        "col_item": "ITEM-CBM-2",
        "col_cbm_raw": "2.50",
        "col_cbm": "1.45",
    }
    item_cbm2 = repo._row_to_invoice_item(invoice_id, row_cbm_explicit, is_adjustment=0)
    assert item_cbm2.col_cbm_raw == "2.50"

    # 3. Adjustment description and amount fallback when is_adjustment=1
    adj_row = {
        "description": "Discount Voucher",
        "amount": -75.50,
        "other_prop": "ignored",
    }
    item_adj = repo._row_to_invoice_item(invoice_id, adj_row, is_adjustment=1)
    assert item_adj.col_desc == "Discount Voucher"
    assert item_adj.col_amount == -75.50
    assert item_adj.is_adjustment == 1

    # 4. Fallback does not trigger when is_adjustment=0
    normal_row = {
        "description": "Not an adjustment desc",
        "amount": 100.0,
    }
    item_normal = repo._row_to_invoice_item(invoice_id, normal_row, is_adjustment=0)
    assert item_normal.col_desc == ""
    assert item_normal.col_amount == 0.0
    assert item_normal.is_adjustment == 0

    # 5. Timestamp handling
    custom_time = "2026-08-20T14:30:00"
    item_ts = repo._row_to_invoice_item(invoice_id, {"col_item": "TS", "timestamp": custom_time})
    assert item_ts.timestamp == datetime.datetime(2026, 8, 20, 14, 30, 0, tzinfo=ICT)


def test_accept_processed_invoice(db):
    repo = InvoiceRepository(db)
    filename = "test_accept_invoice.json"

    # Validation: invalid filename
    for bad_fn in [None, "", "   "]:
        with pytest.raises(ValueError, match=r"(?i)invalid filename"):
            repo.accept_processed_invoice(bad_fn, {"footer_data": {"grand_total": {}}})

    # 1. Validation: missing footer_data or grand_total
    with pytest.raises(ValueError, match=r"(?i)missing footer_data or grand_total"):
        repo.accept_processed_invoice(filename, {})

    with pytest.raises(ValueError, match=r"(?i)missing footer_data or grand_total"):
        repo.accept_processed_invoice(filename, {"footer_data": {}})

    with pytest.raises(ValueError, match=r"(?i)missing footer_data or grand_total"):
        repo.accept_processed_invoice(filename, {"footer_data": {"no_gt": 123}})

    # 2. Full accept flow with multi_table and price_adjustment
    invoice_payload = {
        "footer_data": {
            "grand_total": {
                "col_pallet_count": 4.0,
                "col_qty_sf": 2000.0,
                "col_net": 1000.0,
                "col_amount": 5000.0,
            }
        },
        "multi_table": [
            [
                {
                    "col_item": "ITEM-A1",
                    "col_qty_pcs": 20,
                    "col_qty_sf": 800.0,
                    "col_amount": 2000.0,
                    "col_cbm": "0.45",
                    "ignored_key": "val",
                },
                {
                    "col_item": "ITEM-A2",
                    "col_qty_pcs": 30,
                    "col_qty_sf": 1200.0,
                    "col_amount": 3000.0,
                    "col_cbm_raw": "0.65",
                },
            ]
        ],
        "price_adjustment": [
            {
                "description": "Special Rebate",
                "amount": -200.0,
            },
            {
                "description": "Hazardous Material Surcharge",
                "amount": 100.0,
            },
        ],
    }

    record = repo.accept_processed_invoice(filename, invoice_payload)

    # Verify ProcessedData record
    assert record.id is not None
    assert record.filename == filename
    assert record.status == "Accepted"
    assert record.item_count == 2
    assert record.total_pallets == 4.0
    assert record.total_sqft == 2000.0
    assert record.total_net == 1000.0
    assert record.total_amount == 5000.0  # grand_total already contains overall total
    assert record.data_payload == invoice_payload

    # Verify InvoiceItem records in database
    items = db.query(InvoiceItem).filter(InvoiceItem.invoice_id == filename).all()
    assert len(items) == 4

    line_items = [it for it in items if it.is_adjustment == 0]
    adj_items = [it for it in items if it.is_adjustment == 1]
    assert len(line_items) == 2
    assert len(adj_items) == 2

    # Check line items
    item_map = {it.col_item: it for it in line_items}
    assert "ITEM-A1" in item_map
    assert item_map["ITEM-A1"].col_cbm_raw == "0.45"
    assert item_map["ITEM-A1"].col_amount == 2000.0

    assert "ITEM-A2" in item_map
    assert item_map["ITEM-A2"].col_cbm_raw == "0.65"
    assert item_map["ITEM-A2"].col_amount == 3000.0

    # Check adjustments
    adj_map = {it.col_desc: it for it in adj_items}
    assert "Special Rebate" in adj_map
    assert adj_map["Special Rebate"].col_amount == -200.0
    assert "Hazardous Material Surcharge" in adj_map
    assert adj_map["Hazardous Material Surcharge"].col_amount == 100.0

    # 3. Existing records deleted on re-accept
    new_payload = {
        "footer_data": {
            "grand_total": {
                "col_pallet_count": 1.0,
                "col_qty_sf": 500.0,
                "col_net": 250.0,
                "col_amount": 1000.0,
            }
        },
        "multi_table": [
            [
                {
                    "col_item": "ITEM-REPLACE",
                    "col_qty_sf": 500.0,
                    "col_amount": 1000.0,
                }
            ]
        ],
    }
    updated_record = repo.accept_processed_invoice(filename, new_payload)
    assert updated_record.id == record.id  # Same ProcessedData updated
    assert updated_record.item_count == 1
    assert updated_record.total_amount == 1000.0

    new_items = db.query(InvoiceItem).filter(InvoiceItem.invoice_id == filename).all()
    assert len(new_items) == 1
    assert new_items[0].col_item == "ITEM-REPLACE"

    # 4. Empty multi_table handling
    empty_multi_payload = {
        "footer_data": {
            "grand_total": {
                "col_amount": 750.0,
            }
        },
        "multi_table": [],
    }
    fb_record = repo.accept_processed_invoice("test_fb_empty_multi.json", empty_multi_payload)
    assert fb_record.item_count == 0
    assert fb_record.total_amount == 750.0
    fb_items = db.query(InvoiceItem).filter(InvoiceItem.invoice_id == "test_fb_empty_multi.json").all()
    assert len(fb_items) == 0


def test_accept_processed_invoice_list_style_adjustments(db):
    repo = InvoiceRepository(db)
    filename = "test_list_adj.json"
    payload = {
        "footer_data": {
            "grand_total": {
                "col_pallet_count": 2.0,
                "col_qty_sf": 1000.0,
                "col_net": 500.0,
                "col_amount": 2500.0,
            }
        },
        "multi_table": [
            [
                {
                    "col_item": "ITEM-X",
                    "col_amount": 2500.0,
                }
            ]
        ],
        "price_adjustment": [
            ["Discount", -50.0],
            ["Packing Surcharge", 25.0],
        ],
    }
    record = repo.accept_processed_invoice(filename, payload)
    assert record.total_amount == 2500.0
    items = db.query(InvoiceItem).filter(InvoiceItem.invoice_id == filename).all()
    assert len(items) == 3
    adj_items = [it for it in items if it.is_adjustment == 1]
    assert len(adj_items) == 2
    adj_map = {it.col_desc: it.col_amount for it in adj_items}
    assert adj_map["Discount"] == -50.0
    assert adj_map["Packing Surcharge"] == 25.0


def test_generate_registry_csv(db):
    repo = InvoiceRepository(db)
    db.query(InvoiceItem).delete()
    db.commit()

    items = [
        {
            "col_item": "ITEM-001",
            "col_desc": "Widget A",
            "col_qty_pcs": 10,
            "col_qty_sf": 100.0,
            "col_pallet_count": 2.0,
            "col_net": 50.0,
            "col_gross": 55.0,
            "col_unit_price": 5.0,
            "col_amount": 500.0,
            "timestamp": "2026-09-01T10:00:00",
        },
        {
            "col_item": "ITEM-002",
            "col_desc": "Widget B",
            "col_qty_pcs": 20,
            "col_qty_sf": 200.0,
            "col_pallet_count": 3.0,
            "col_net": 100.0,
            "col_gross": 110.0,
            "col_unit_price": 6.0,
            "col_amount": 1200.0,
            "timestamp": "2026-09-02T10:00:00",
        },
    ]
    repo.save_invoice_items("csv_test.json", items, delete_existing=True)

    csv_data = repo.generate_registry_csv()
    assert csv_data.startswith("\ufeff")

    reader = list(csv.reader(io.StringIO(csv_data.lstrip("\ufeff"))))
    header = reader[0]
    assert "ID" in header
    assert "Invoice ID" in header
    assert "Amount" in header
    assert "Qty (SF)" in header

    # Verify rows
    data_rows = [r for r in reader[1:] if r and r[header.index("Invoice ID")] != "TOTAL"]
    assert len(data_rows) == 2

    # Verify summary row
    summary_rows = [r for r in reader[1:] if r and r[header.index("Invoice ID")] == "TOTAL"]
    assert len(summary_rows) == 1
    summary = summary_rows[0]
    assert float(summary[header.index("Qty (SF)")]) == 300.0
    assert float(summary[header.index("Amount")]) == 1700.0
    assert float(summary[header.index("Pallet Count")]) == 5.0
    assert float(summary[header.index("Net Weight")]) == 150.0
    assert float(summary[header.index("Gross Weight")]) == 165.0

    # With date filtering
    filtered_csv = repo.generate_registry_csv(start_date="2026-09-02")
    filtered_reader = list(csv.reader(io.StringIO(filtered_csv.lstrip("\ufeff"))))
    filtered_data = [r for r in filtered_reader[1:] if r and r[filtered_reader[0].index("Invoice ID")] != "TOTAL"]
    assert len(filtered_data) == 1
    assert filtered_data[0][filtered_reader[0].index("Item")] == "ITEM-002"


def test_query_items_invalid_date_raises(db):
    repo = InvoiceRepository(db)
    with pytest.raises(ValueError, match="Invalid start_date"):
        repo.query_items(start_date="invalid-date")
    with pytest.raises(ValueError, match="Invalid end_date"):
        repo.query_items(end_date="invalid-date")


def test_save_invoice_items_guard_non_dict(db):
    repo = InvoiceRepository(db)
    filename = "guard_test.json"
    items = [
        {"col_item": "VALID-1", "col_amount": 100.0},
        None,
        "not a dict",
        123,
        {"col_item": "VALID-2", "col_amount": 200.0},
    ]
    count = repo.save_invoice_items(filename, items, delete_existing=True)
    assert count == 2
    stored = db.query(InvoiceItem).filter(InvoiceItem.invoice_id == filename).all()
    assert len(stored) == 2
