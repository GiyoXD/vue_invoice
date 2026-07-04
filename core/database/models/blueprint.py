from sqlalchemy import Column, Integer, String, Text, DateTime, ForeignKey, LargeBinary
from sqlalchemy.orm import relationship
from core.database.session import Base
from core.database.json_util import JSONText
from core.utils.clock import now as ict_now

class Blueprint(Base):
    __tablename__ = "blueprints"

    id = Column(Integer, primary_key=True, index=True)
    customer_code = Column(String, nullable=False)
    locale = Column(String, nullable=False, default="KH")
    description = Column(Text, nullable=True)
    config_json = Column(JSONText, nullable=False)
    template_json = Column(JSONText, nullable=False)
    created_at = Column(DateTime, default=ict_now)
    updated_at = Column(DateTime, default=ict_now, onupdate=ict_now)

    template_binary = relationship("BlueprintTemplate", uselist=False, back_populates="blueprint", cascade="all, delete-orphan")

class BlueprintTemplate(Base):
    __tablename__ = "blueprint_templates"

    blueprint_id = Column(Integer, ForeignKey("blueprints.id", ondelete="CASCADE"), primary_key=True)
    filename = Column(String, nullable=False)
    xlsx_blob = Column(LargeBinary, nullable=False)

    blueprint = relationship("Blueprint", back_populates="template_binary")
