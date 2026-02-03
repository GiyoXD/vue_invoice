import functools
import logging
import uuid
import contextvars
import sys

# 1. The Global Context (Invisible Backpack)
_trace_id_ctx = contextvars.ContextVar("trace_id", default="NO-TRACE")

# Use centralized logger - no basicConfig here
logger = logging.getLogger("BACKEND_WORKFLOW")

def start_trace(custom_id=None):
    """
    Call this ONCE at the very top of the CLI/Script.
    Sets a trace ID for tracking execution flow.
    
    Note: File logging is now handled by core.logger_config.setup_logging()
    """
    tid = custom_id or f"run-{str(uuid.uuid4())[:8]}"
    _trace_id_ctx.set(tid)
    logger.info(f"[{tid}] Trace started")
    return tid

def get_trace_id():
    """Retrieve the current ID anywhere in the code."""
    return _trace_id_ctx.get()

def snitch(func):
    """Decorator to log entry/exit with the ID."""
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        tid = get_trace_id()
        func_name = func.__qualname__
        try:
            logger.info(f"[{tid}] >> ENTER: {func_name}")
            result = func(*args, **kwargs)
            logger.info(f"[{tid}] OK EXIT:  {func_name}")
            return result
        except Exception as e:
            logger.error(f"[{tid}] !! CRASH: {func_name} | {e}")
            raise e
    return wrapper
