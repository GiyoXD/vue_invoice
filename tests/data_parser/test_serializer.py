import datetime
import decimal
import pytest
from core.data_parser.util.serializer import Serializer


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

