from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Any, Optional

@dataclass
class GenerationOptions:
    """Encapsulates all generation mode flags and output options."""
    daf_mode: bool = False
    custom_mode: bool = False
    enable_auto_fit: bool = True
    split_sheets: bool = False
    return_bytes: bool = False
    explicit_config_data: Optional[Dict[str, Any]] = None
    explicit_template_json_data: Optional[Dict[str, Any]] = None
    explicit_template_xlsx_bytes: Optional[bytes] = None

@dataclass
class InvoicePathConfig:
    """Configures primary input/output files and search directories."""
    input_data_path: Path
    output_path: Path
    template_dir: Optional[Path] = None
    config_dir: Optional[Path] = None

@dataclass
class ExplicitOverrides:
    """Manual overrides to skip resolution (direct data, manual templates)."""
    explicit_config_path: Optional[Path] = None
    explicit_template_path: Optional[Path] = None
    input_data_dict: Optional[Dict[str, Any]] = None

@dataclass
class InvoiceGenerationRequest:
    """Aggregates all inputs required to run the invoice generation pipeline."""
    paths: InvoicePathConfig
    overrides: ExplicitOverrides
    options: GenerationOptions = field(default_factory=GenerationOptions)
