"""
Style Registry - Centralized ID-Driven Styling System

This module provides the core styling architecture where:
- Column IDs define WHAT (data format, alignment, width)
- Row contexts define HOW (visual emphasis, decoration)
- Cell style = Column base + Row context merged

Pattern:
    col_cbm (column) → {format: "0.00", alignment: "center"}
    header (context) → {bold: True, fill: "CCCCCC"}
    merged style     → {format: "0.00", alignment: "center", bold: True, fill: "CCCCCC"}
"""

import logging
from typing import Dict, Any, Optional
from dataclasses import dataclass

logger = logging.getLogger(__name__)


@dataclass
class ColumnStyle:
    """Base style definition for a column (data format and alignment)."""
    col_id: str
    format: Optional[str] = None  # Number format: "@", "0.00", "#,##0", etc.
    alignment: Optional[str] = None  # "left", "center", "right"
    vertical_alignment: Optional[str] = None  # "top", "center", "bottom"
    width: Optional[int] = None  # Column width
    wrap_text: bool = False
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for merging."""
        return {
            'format': self.format,
            'alignment': self.alignment,
            'vertical_alignment': self.vertical_alignment,
            'width': self.width,
            'wrap_text': self.wrap_text
        }


@dataclass
class RowContextStyle:
    """Context-specific style (visual emphasis for header/data/footer)."""
    context: str  # "header", "data", "footer", "grand_total"
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    font_size: Optional[int] = None
    font_name: Optional[str] = None
    fill_color: Optional[str] = None  # Hex color: "CCCCCC"
    border_style: Optional[str] = None  # "thin", "medium", "thick"
    row_height: Optional[int] = None
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for merging."""
        return {
            'bold': self.bold,
            'italic': self.italic,
            'font_size': self.font_size,
            'font_name': self.font_name,
            'fill_color': self.fill_color,
            'border_style': self.border_style,
            'row_height': self.row_height
        }


class StyleRegistry:
    """
    Centralized registry for ID-driven cell styling.
    
    Maps (column_id, row_context) → merged cell style
    
    Usage:
        registry = StyleRegistry(sheet_config)
        
        # Get style for specific cell
        style = registry.get_style('col_cbm', context='data')
        # Returns: {format: "0.00", alignment: "center", bold: False, ...}
        
        # Get style for header cell
        header_style = registry.get_style('col_cbm', context='header')
        # Returns: {format: "0.00", alignment: "center", bold: True, fill: "CCCCCC", ...}
    """
    
    def __init__(self, sheet_config: Dict[str, Any]):
        """
        Initialize registry from sheet configuration.
        
        Args:
            sheet_config: Sheet configuration containing:
                - columns: {col_id: {format, alignment, width}}
                - row_contexts: {context: {bold, fill_color, borders}}
        """
        self.sheet_config = sheet_config
        self.columns: Dict[str, ColumnStyle] = {}
        self.row_contexts: Dict[str, RowContextStyle] = {}
        
        self._load_columns()
        self._load_row_contexts()
    
    def _load_columns(self):
        """Load column definitions from config."""
        columns_config = self.sheet_config.get('columns', {})
        
        for col_id, col_def in columns_config.items():
            self.columns[col_id] = ColumnStyle(
                col_id=col_id,
                format=col_def.get('format'),
                alignment=col_def.get('alignment'),
                vertical_alignment=col_def.get('vertical_alignment'),
                width=col_def.get('width'),
                wrap_text=col_def.get('wrap_text', False)
            )
            logger.debug(f"Loaded column '{col_id}': alignment={col_def.get('alignment')}, vertical_alignment={col_def.get('vertical_alignment')}")
        
        logger.debug(f"Loaded {len(self.columns)} column styles: {list(self.columns.keys())}")
    
    def _load_row_contexts(self):
        """Load row context styles from config."""
        contexts_config = self.sheet_config.get('row_contexts', {})
        
        for context, context_def in contexts_config.items():
            self.row_contexts[context] = RowContextStyle(
                context=context,
                bold=context_def.get('bold'),
                italic=context_def.get('italic'),
                font_size=context_def.get('font_size'),
                font_name=context_def.get('font_name'),
                fill_color=context_def.get('fill_color'),
                border_style=context_def.get('border_style'),
                row_height=context_def.get('row_height')
            )
        
        logger.debug(f"Loaded {len(self.row_contexts)} row contexts: {list(self.row_contexts.keys())}")
        
        # Check if border_style is missing from all contexts
        has_any_border = any(ctx.border_style for ctx in self.row_contexts.values())
        if not has_any_border and self.row_contexts:
            logger.warning(f"ℹ️  No 'border_style' found in any row_contexts")
            logger.warning(f"   Cells will have NO borders unless added to config")
            logger.warning(f"   To add borders: Add 'border_style: thin' (or 'medium'/'thick') to row_contexts")
    
    def get_style(self, col_id: str, context: str = 'data', overrides: Optional[Dict] = None) -> Dict[str, Any]:
        """
        Get merged style for a specific cell.
        
        Merge priority: Column base → Row context → Overrides
        
        Args:
            col_id: Column identifier (e.g., "col_cbm", "col_po")
            context: Row context (e.g., "header", "data", "footer")
            overrides: Optional style overrides for special cases
        
        Returns:
            Merged style dictionary with all properties
        
        Example:
            style = registry.get_style('col_cbm', context='header')
            # Returns: {
            #     'format': '0.00',           # from column
            #     'alignment': 'center',      # from column
            #     'bold': True,               # from context
            #     'fill_color': 'CCCCCC'      # from context
            # }
        """
        merged_style = {}
        
        # Define column-owned properties (NEVER override these from context)
        COLUMN_OWNED = {'format', 'alignment', 'vertical_alignment', 'width', 'wrap_text'}
        
        # 1. Get column base style (WHAT: format, alignment)
        if col_id in self.columns:
            col_style = self.columns[col_id].to_dict()
            logger.debug(f"Column '{col_id}' style dict: {col_style}")
            merged_style.update({k: v for k, v in col_style.items() if v is not None})
            logger.debug(f"After column merge: {merged_style}")
        else:
            logger.warning(f"❌ Column '{col_id}' not found in StyleRegistry!")
            logger.warning(f"   Available columns: {list(self.columns.keys())}")
            logger.warning(f"   Please add column definition to config with: format, alignment, width")
        
        # 2. Merge row context style (HOW: emphasis, decoration)
        # CRITICAL: Only merge properties that are NOT column-owned
        if context in self.row_contexts:
            context_style = self.row_contexts[context].to_dict()
            for key, value in context_style.items():
                if value is not None and key not in COLUMN_OWNED:
                    merged_style[key] = value
        else:
            logger.warning(f"❌ Row context '{context}' not found in StyleRegistry!")
            logger.warning(f"   Available contexts: {list(self.row_contexts.keys())}")
            logger.warning(f"   Please add context definition to config with: bold, font_size, font_name, etc.")
        
        # 3. Apply overrides (special cases - can override anything)
        if overrides:
            merged_style.update(overrides)
        
        # 4. STRICT VALIDATION: Verify all required properties exist
        required_props = {
            'alignment': f"Add 'alignment' to styling_bundle.{self.sheet_config.get('sheet_name', 'Sheet')}.columns.{col_id}",
            'format': f"Add 'format' to styling_bundle.{self.sheet_config.get('sheet_name', 'Sheet')}.columns.{col_id}",
            'font_name': f"Add 'font_name' to styling_bundle.{self.sheet_config.get('sheet_name', 'Sheet')}.row_contexts.{context}",
            'font_size': f"Add 'font_size' to styling_bundle.{self.sheet_config.get('sheet_name', 'Sheet')}.row_contexts.{context}"
        }
        
        missing_props = []
        for prop, instruction in required_props.items():
            if prop not in merged_style or merged_style[prop] is None:
                missing_props.append(prop)
                logger.warning(f"❌ StyleRegistry.get_style(col_id='{col_id}', context='{context}'): Missing required '{prop}'")
                logger.warning(f"   → {instruction}")
        
        if missing_props:
            logger.error(f"BROKEN INCOMPLETE STYLE: col_id='{col_id}', context='{context}' - missing {missing_props}")
            logger.error(f"   → Merged style keys: {list(merged_style.keys())}")
            logger.error(f"   → This will cause CellStyler to skip applying this style!")
        
        return merged_style
    
    def get_column_width(self, col_id: str) -> Optional[int]:
        """Get column width for a specific column ID."""
        if col_id in self.columns:
            return self.columns[col_id].width
        return None
    
    def get_row_height(self, context: str) -> Optional[int]:
        """Get row height for a specific context."""
        if context in self.row_contexts:
            return self.row_contexts[context].row_height
        return None
    
    def has_column(self, col_id: str) -> bool:
        """Check if column ID exists in registry."""
        return col_id in self.columns
    
    def has_context(self, context: str) -> bool:
        """Check if row context exists in registry."""
        return context in self.row_contexts
    
    @classmethod
    def create_from_styling_bundle(cls, styling_config: Dict[str, Any], sheet_name: str) -> 'StyleRegistry':
        """
        Factory method to create registry from styling_bundle config.
        
        Args:
            styling_config: The styling_bundle section from config
            sheet_name: Name of sheet to load styles for
        
        Returns:
            StyleRegistry instance
        """
        sheet_config = styling_config.get(sheet_name, {})
        return cls(sheet_config)
