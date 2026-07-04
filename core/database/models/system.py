from sqlalchemy import Column, String, Text
from core.database.session import Base

class SystemSetting(Base):
    __tablename__ = "system_settings"

    key = Column(String, primary_key=True)
    value_json = Column(Text, nullable=False)
