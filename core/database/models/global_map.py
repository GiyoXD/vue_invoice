from sqlalchemy import Column, Integer, String, ForeignKey, UniqueConstraint
from sqlalchemy.orm import relationship
from core.database.session import Base

class GlobalMapColumn(Base):
    __tablename__ = "global_map_columns"
    
    col_id = Column(String(64), primary_key=True)
    excel_format = Column(String(64), default="@")
    
    keywords = relationship("GlobalMapColumnKeyword", back_populates="column", cascade="all, delete-orphan")
    overrides = relationship("GlobalMapHeaderTextMapping", back_populates="column", cascade="all, delete-orphan")

class GlobalMapColumnKeyword(Base):
    __tablename__ = "global_map_column_keywords"
    
    id = Column(Integer, primary_key=True, autoincrement=True)
    col_id = Column(String(64), ForeignKey("global_map_columns.col_id", ondelete="CASCADE"), nullable=False)
    keyword = Column(String(128), nullable=False)
    
    column = relationship("GlobalMapColumn", back_populates="keywords")
    
    __table_args__ = (
        UniqueConstraint("col_id", "keyword", name="uq_global_map_col_keyword"),
    )

class GlobalMapHeaderTextMapping(Base):
    __tablename__ = "global_map_header_text_mappings"
    
    raw_text = Column(String(256), primary_key=True)
    canonical_col_id = Column(
        String(64), 
        ForeignKey("global_map_columns.col_id", ondelete="CASCADE"), 
        nullable=False
    )
    
    column = relationship("GlobalMapColumn", back_populates="overrides")

class GlobalMapSheet(Base):
    __tablename__ = "global_map_sheets"
    
    sheet_name = Column(String(128), primary_key=True)
    processing_type = Column(String(64), nullable=False)  # 'aggregation' or 'processed_tables'

class GlobalMapFooterLabelKeyword(Base):
    __tablename__ = "global_map_footer_label_keywords"
    
    keyword = Column(String(128), primary_key=True)

class GlobalMapFallbackStrategy(Base):
    __tablename__ = "global_map_fallback_strategies"
    
    strategy_key = Column(String(128), primary_key=True)
    strategy_value = Column(String(256), nullable=False)
