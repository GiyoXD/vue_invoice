import time
import copy
import threading
from typing import Any, Optional

class InMemoryCache:
    """Thread-safe simple TTL in-memory cache helper."""
    def __init__(self, ttl_seconds: float = 60.0):
        self._data: Optional[Any] = None
        self._timestamp: float = 0.0
        self.ttl = ttl_seconds
        self._lock = threading.Lock()

    def get(self) -> Optional[Any]:
        """Retrieve copy of cached data if it exists and has not expired."""
        with self._lock:
            now = time.time()
            if self._data is not None and (now - self._timestamp) < self.ttl:
                return copy.deepcopy(self._data)
            return None

    def set(self, data: Any) -> None:
        """Store copy of data and record timestamp."""
        with self._lock:
            self._data = copy.deepcopy(data)
            self._timestamp = time.time()

    def invalidate(self) -> None:
        """Reset cached data."""
        with self._lock:
            self._data = None
            self._timestamp = 0.0

# Singleton mapping configuration cache
mapping_cache = InMemoryCache(ttl_seconds=60.0)
