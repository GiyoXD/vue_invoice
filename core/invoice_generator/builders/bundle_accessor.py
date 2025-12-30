"""
Bundle accessor class providing common bundle storage and access patterns.

This class provides shared bundle storage, property accessors, and helper methods
used across LayoutBuilder and the BuilderStyler classes to eliminate code duplication.

IMPORTANT: This is NOT a builder - it's pure infrastructure for bundle access.
The actual building and styling work happens in:
- LayoutBuilder (Director that coordinates the process)
- HeaderBuilderStyler (Builds + styles headers)
- DataTableBuilderStyler (Builds + styles data tables)
- FooterBuilderStyler (Builds + styles footers)

The 'style_config' bundle stored here contains styling configuration,
but this class doesn't DO any styling - it just provides ACCESS to the config.
Actual styling is delegated to the style_applier module.
"""
import logging
from typing import Any, Dict, Optional
from openpyxl.worksheet.worksheet import Worksheet

logger = logging.getLogger(__name__)

from ..styling.models import StylingConfigModel


class BundleAccessor:
    """
    Base class providing bundle storage and access patterns.
    
    Stores bundle dictionaries and provides common @property accessors for
    frequently used bundle values. Also includes shared helper methods.
    
    Child classes (builders) inherit this to get consistent bundle access
    without code duplication.
    """
    
    def __init__(
        self,
        worksheet: Worksheet,
        style_config: Dict[str, Any],
        context_config: Dict[str, Any],
        **kwargs
    ):
        """
        Initialize base builder with core bundles.
        
        Args:
            worksheet: The worksheet to build in
            style_config: Bundle containing styling_config
            context_config: Bundle containing contextual information (sheet_name, etc.)
            **kwargs: Additional bundles specific to child classes (layout_config, data_config, etc.)
        """
        self.worksheet = worksheet
        self.style_config = style_config
        self.context_config = context_config
        
        # Store additional bundles passed by child classes
        for key, value in kwargs.items():
            setattr(self, key, value)
    
    # ========== Common Properties ==========
    
    @property
    def sheet_name(self) -> str:
        """Sheet name from context config."""
        return self.context_config.get('sheet_name', '')
    
    @property
    def all_sheet_configs(self) -> Dict[str, Any]:
        """All sheet configurations from context config."""
        return self.context_config.get('all_sheet_configs', {})
    
    @property
    def sheet_styling_config(self) -> Optional[StylingConfigModel]:
        """
        Styling configuration from style config.
        Automatically converts dict to StylingConfigModel if needed.
        """
        styling_config = self.style_config.get('styling_config')
        if styling_config and not isinstance(styling_config, StylingConfigModel):
            try:
                styling_config = StylingConfigModel(**styling_config)
            except Exception as e:
                logger.warning(f"Could not create StylingConfigModel: {e}")
                styling_config = None
        return styling_config
    
    @property
    def args(self):
        """Command-line arguments from context config."""
        return self.context_config.get('args')
    
    # ========== Common Helper Methods ==========
    
    def _apply_footer_row_height(self, footer_row: int):
        """
        Apply footer height to a single footer row.
        
        This method handles the logic for determining footer height based on
        styling configuration, including matching header height if configured.
        
        Args:
            footer_row: The row number to apply footer height to
        """
        if not self.sheet_styling_config or not self.sheet_styling_config.rowHeights:
            return
        
        row_heights_cfg = self.sheet_styling_config.rowHeights
        footer_height_config = row_heights_cfg.get("footer")
        match_header_height_flag = row_heights_cfg.get("footer_matches_header_height", True)
        
        # Determine the footer height
        final_footer_height = None
        if match_header_height_flag:
            # Get header height from config
            header_height = row_heights_cfg.get("header")
            if header_height is not None:
                final_footer_height = header_height
        if final_footer_height is None and footer_height_config is not None:
            final_footer_height = footer_height_config
        
        # Apply the height
        if final_footer_height is not None and footer_row > 0:
            try:
                h_val = float(final_footer_height)
                if h_val > 0:
                    self.worksheet.row_dimensions[footer_row].height = h_val
            except (ValueError, TypeError):
                pass
    
    def _get_bool_flag(self, config_dict: Dict[str, Any], key: str, default: bool = False) -> bool:
        """
        Safely retrieve a boolean flag from a config dictionary.
        
        Args:
            config_dict: The configuration dictionary
            key: The key to retrieve
            default: Default value if key not found
            
        Returns:
            The boolean value or default
        """
        return config_dict.get(key, default)

