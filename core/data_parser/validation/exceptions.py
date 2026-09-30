class DataValidationError(Exception):
    """User-facing validation error for missing required data.
    
    This exception is caught separately in the API layer to return
    a clean, human-readable error message without traceback noise.
    """
    pass
