import datetime
import decimal
import json
import pytest
from core.data_parser.util.serializer import Serializer
from core.data_parser.data_processor.single_table.serializer import format_aggregation_as_list


def test_serializer_default_types():
    # Test datetime and date serialization
    dt = datetime.datetime(2026, 7, 4, 12, 0, 0)
    d = datetime.date(2026, 7, 4)
    assert Serializer._serialize_unsupported_value(dt) == "2026-07-04T12:00:00"
    assert Serializer._serialize_unsupported_value(d) == "2026-07-04"

    # Test Decimal serialization
    dec = decimal.Decimal("123.45")
    assert Serializer._serialize_unsupported_value(dec) == 123.45

    # Test set serialization
    s = {1, 2, 3}
    assert sorted(Serializer._serialize_unsupported_value(s)) == [1, 2, 3]

    # Test unsupported type
    class Unsupported:
        pass

    with pytest.raises(TypeError):
        Serializer._serialize_unsupported_value(Unsupported())


def test_make_serializable():
    # Test nested structure with tuple keys
    input_data = {
        ("A", "B"): {
            "value": decimal.Decimal("10.0"),
            "date": datetime.date(2026, 7, 4),
        },
        "list_val": [
            {("Nested", "Tuple"): "ok"},
            "simple_str",
            None
        ]
    }

    expected_output = {
        "('A', 'B')": {
            "value": decimal.Decimal("10.0"),
            "date": datetime.date(2026, 7, 4),
        },
        "list_val": [
            {"('Nested', 'Tuple')": "ok"},
            "simple_str",
            None
        ]
    }

    assert Serializer._stringify_keys_recursively(input_data) == expected_output


def test_serialize_to_json():
    import json
    input_data = {
        ("A", "B"): {
            "value": decimal.Decimal("12.34"),
            "date": datetime.date(2026, 7, 4),
        }
    }

    json_str = Serializer.serialize_to_json(input_data)
    parsed = json.loads(json_str)

    assert parsed == {
        "('A', 'B')": {
            "value": 12.34,
            "date": "2026-07-04"
        }
    }


def test_format_aggregation_as_list():
    aggregation_map = {
        "item_1": {
            "item": "Product A",
            "quantity": 10,
            "unit_price": decimal.Decimal("15.50"),
            "total_amount": decimal.Decimal("155.00"),
        },
        "item_2": {
            "item": "Product B",
            "quantity": 3,
            "unit_price": decimal.Decimal("9.99"),
            "total_amount": decimal.Decimal("29.97"),
        },
    }

    # 1. Test convert_decimals=False (default): retains Decimal instances
    result_default = format_aggregation_as_list(aggregation_map)
    assert len(result_default) == 2
    assert result_default[0]["item"] == "Product A"
    assert isinstance(result_default[0]["unit_price"], decimal.Decimal)
    assert result_default[0]["unit_price"] == decimal.Decimal("15.50")
    assert isinstance(result_default[0]["total_amount"], decimal.Decimal)
    assert result_default[0]["total_amount"] == decimal.Decimal("155.00")

    result_explicit_false = format_aggregation_as_list(aggregation_map, convert_decimals=False)
    assert isinstance(result_explicit_false[0]["unit_price"], decimal.Decimal)

    # Standard json.dumps cannot serialize Decimal objects without custom encoder
    with pytest.raises(TypeError):
        json.dumps(result_default)

    # 2. Test convert_decimals=True: converts Decimal to float, directly serializable with standard json.dumps
    result_converted = format_aggregation_as_list(aggregation_map, convert_decimals=True)
    assert len(result_converted) == 2
    assert result_converted[0]["item"] == "Product A"
    assert isinstance(result_converted[0]["unit_price"], float)
    assert result_converted[0]["unit_price"] == 15.50
    assert isinstance(result_converted[0]["total_amount"], float)
    assert result_converted[0]["total_amount"] == 155.00

    # Ensure directly serializable with standard json.dumps
    json_str = json.dumps(result_converted)
    parsed = json.loads(json_str)
    assert parsed == [
        {
            "item": "Product A",
            "quantity": 10,
            "unit_price": 15.50,
            "total_amount": 155.00,
        },
        {
            "item": "Product B",
            "quantity": 3,
            "unit_price": 9.99,
            "total_amount": 29.97,
        },
    ]


