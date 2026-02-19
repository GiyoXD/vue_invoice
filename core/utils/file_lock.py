import os
import time
import logging
import subprocess
from pathlib import Path

logger = logging.getLogger(__name__)

def is_file_locked(filepath: str) -> bool:
    """
    Check if a file is locked by attempting to open it in append mode.
    """
    if not os.path.exists(filepath):
        return False
        
    try:
        # Try to open the file in append mode to check for write access
        # If open fails, it's likely locked
        with open(filepath, 'a'):
            pass
        return False
    except IOError:
        return True
    except Exception as e:
        logger.warning(f"Error checking file lock for {filepath}: {e}")
        return True

def ensure_file_unlocked(filepath: Path, max_retries: int = 3, retry_delay: float = 1.0):
    """
    Ensures a file is unlocked. If locked, attempts to kill the process holding it
    (specifically targeting Excel by window title) and retries.
    
    Args:
        filepath: Path object to the file
        max_retries: Number of times to retry checking after attempting to kill process
        retry_delay: Seconds to wait between retries
    """
    path_str = str(filepath)
    filename = filepath.name
    
    if not is_file_locked(path_str):
        return

    logger.warning(f"File '{filename}' is currently locked. Attempting to release lock...")

    # Construct PowerShell command to find and kill Excel process with matching window title
    # Get-Process | Where-Object {$_.MainWindowTitle -like "*filename*"} | Stop-Process -Force
    # We use -like "*name*" to handle "Microsoft Excel - filename" or "filename - Excel" etc.
    
    ps_command = f"""
    $target = "{filename}";
    Get-Process | Where-Object {{ $_.MainWindowTitle -like "*$target*" }} | Stop-Process -Force
    """
    
    try:
        logger.info(f"Attempting to kill process holding lock on '{filename}'...")
        result = subprocess.run(
            ["powershell", "-Command", ps_command],
            capture_output=True,
            text=True
        )
        
        if result.returncode == 0:
            logger.info("Process kill command executed successfully.")
        else:
            logger.warning(f"Process kill command failed: {result.stderr}")
            
    except Exception as e:
        logger.error(f"Failed to execute process kill command: {e}")

    # Wait a bit for OS to release handles
    time.sleep(retry_delay)
    
    # Verify strict unlocking
    for i in range(max_retries):
        if not is_file_locked(path_str):
            logger.info(f"Successfully released lock on '{filename}'.")
            return
        logger.warning(f"File still locked. Retrying check ({i+1}/{max_retries})...")
        time.sleep(retry_delay)
        
    # If still locked, raise error
    raise PermissionError(
        f"File '{filename}' is locked by another process (likely Excel). "
        "We attempted to close it but failed. Please close the file manually."
    )
