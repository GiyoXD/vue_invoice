import json
from sqlalchemy import Text, TypeDecorator

class JSONText(TypeDecorator):
    """Enables transparent JSON storage in SQLite by encoding/decoding on the fly."""
    impl = Text
    cache_ok = True

    def process_bind_param(self, value, dialect):
        if value is None:
            return None
        if isinstance(value, str):
            return value
        return json.dumps(value, ensure_ascii=False)

    def process_result_value(self, value, dialect):
        if value is None:
            return None
        if isinstance(value, str):
            try:
                return json.loads(value)
            except Exception:
                return value
        return value
