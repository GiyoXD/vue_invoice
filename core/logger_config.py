# core/logger_config.py
"""
Centralized logging configuration for the Invoice Generator.

This module provides a single source of truth for logging configuration.
Call setup_logging() ONCE at application startup before any other imports.

Usage:
    from core.logger_config import setup_logging
    from core.system_config import sys_config
    
    setup_logging(log_dir=sys_config.run_log_dir)
"""
import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Optional

# Track if logging has been configured to prevent double-init
_logging_initialized = False


def setup_logging(
    log_dir: Path,
    level: int = logging.INFO,
    log_filename: str = "invoice_generator.log",
    max_bytes: int = 5 * 1024 * 1024,  # 5 MB
    backup_count: int = 3
) -> None:
    """
    Configure logging for the entire application.
    
    Args:
        log_dir: Directory to write log files (from RUN_LOG_DIR env var)
        level: Console logging level (DEBUG, INFO, WARNING, etc.)
        log_filename: Name of the log file
        max_bytes: Maximum size per log file before rotation
        backup_count: Number of backup files to keep
    
    Note:
        - Console shows logs at the specified level
        - File always captures DEBUG level for thorough debugging
        - This function is idempotent; calling it multiple times has no effect
    """
    global _logging_initialized
    
    if _logging_initialized:
        logging.debug("Logging already initialized, skipping.")
        return
    
    # Ensure log directory exists
    log_dir = Path(log_dir)
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / log_filename
    
    # Unified format for all logs
    # Format: timestamp | LEVEL    | module:line | message
    formatter = logging.Formatter(
        fmt='%(asctime)s | %(levelname)-8s | %(name)s:%(lineno)d | %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    
    # Console Handler (stdout)
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    console_handler.setLevel(level)
    
    # Handler 1: Rolling History Log (Keeps everything, rotates)
    history_handler = RotatingFileHandler(
        log_file,
        mode='a',
        maxBytes=max_bytes,
        backupCount=backup_count,
        encoding='utf-8'
    )
    history_handler.setFormatter(formatter)
    history_handler.setLevel(logging.DEBUG)

    # Handler 2: Current Session Log (Wiped on every start)
    session_log_file = log_dir / "current_session.log"
    session_handler = logging.FileHandler(
        session_log_file,
        mode='w',  # Overwrite mode: Clean log for debugging current run
        encoding='utf-8'
    )
    session_handler.setFormatter(formatter)
    session_handler.setLevel(logging.DEBUG)
    
    # Configure Root Logger
    root_logger = logging.getLogger()
    root_logger.setLevel(logging.DEBUG)  # Allow all; handlers filter
    
    # Remove any existing handlers (in case of partial init)
    root_logger.handlers.clear()
    
    root_logger.addHandler(console_handler)
    root_logger.addHandler(history_handler)
    root_logger.addHandler(session_handler)
    
    _logging_initialized = True
    logging.info(f"Logging initialized.")
    logging.info(f"  History Log: {log_file}")
    logging.info(f"  Session Log: {session_log_file}")


def clear_session_log():
    """
    Manually clear the contents of current_session.log.
    Useful for persistent processes (like API servers) where setup_logging 
    only runs once, but we want a fresh log for each task run.
    """
    root_logger = logging.getLogger()
    for handler in root_logger.handlers:
        # Check if it's a FileHandler and looks like our session log
        if isinstance(handler, logging.FileHandler) and "current_session.log" in str(handler.baseFilename):
            try:
                # SAFER METHOD FOR WINDOWS:
                # Use the existing open stream to truncate, rather than opening a new handle
                # which can cause "binary" glitches or locking errors.
                if handler.stream and not handler.stream.closed:
                    handler.acquire()  # Thread-safe lock
                    try:
                        handler.stream.seek(0)
                        handler.stream.truncate(0)
                        handler.stream.flush()
                    finally:
                        handler.release()
            except Exception as e:
                print(f"Warning: Failed to clear session log: {e}")
            return


def get_logger(name: str) -> logging.Logger:
    """
    Get a logger instance for a module.
    
    This is a convenience wrapper around logging.getLogger().
    
    Args:
        name: Usually __name__ of the calling module
        
    Returns:
        Logger instance
    """
    return logging.getLogger(name)
