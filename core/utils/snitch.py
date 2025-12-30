import functools
import logging
import uuid
import contextvars
import sys

# 1. The Global Context (Invisible Backpack)
_trace_id_ctx = contextvars.ContextVar("trace_id", default="NO-TRACE")

# 2. Console Logger Setup (Standardized)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
    force=True
)
logger = logging.getLogger("BACKEND_WORKFLOW")

def start_trace(custom_id=None):
    """Call this ONCE at the very top of the CLI/Script."""
    tid = custom_id or f"run-{str(uuid.uuid4())[:8]}"
    _trace_id_ctx.set(tid)
    
    # Setup File Logging
    try:
        from core.system_config import sys_config
        log_dir = sys_config.run_log_dir
        log_dir.mkdir(parents=True, exist_ok=True)
        
        log_file = log_dir / f"snitch_trace_{tid}.log"
        
        file_handler = logging.FileHandler(log_file, encoding='utf-8')
        file_handler.setLevel(logging.INFO)
        formatter = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", datefmt="%H:%M:%S")
        file_handler.setFormatter(formatter)
        
        # Attach to ROOT logger so we capture ALL logs from all modules (DataParser, etc.)
        root_logger = logging.getLogger()
        
        # Avoid adding duplicate handlers to Root if called multiple times
        # We check if a FileHandler for this exact file already exists
        already_has_file_handler = False
        for h in root_logger.handlers:
             if isinstance(h, logging.FileHandler) and str(h.baseFilename) == str(log_file):
                 already_has_file_handler = True
                 break
        
        if not already_has_file_handler:
            root_logger.addHandler(file_handler)
            logger.info(f"[{tid}] 📝 Logging to file: {log_file}")
        else:
             logger.info(f"[{tid}] 📝 Logging to EXISTING file: {log_file}")
        
    except Exception as e:
        logger.warning(f"[{tid}] ⚠️ Could not setup file logging: {e}")
        
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
