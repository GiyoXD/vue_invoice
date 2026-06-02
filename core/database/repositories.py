# core/database/repositories.py
import json
import logging
from typing import Optional, List
from sqlalchemy.orm import Session
from core.database.db_manager import Blueprint, BlueprintTemplate

logger = logging.getLogger(__name__)

class BlueprintRepository:
    """
    Encapsulates all database query and storage operations for Blueprints.
    This separates database access from core business logic (resolvers, generators, etc.).
    """

    def __init__(self, db: Session):
        self.db = db

    def get_blueprint(self, customer_code: str, locale: str = "KH") -> Optional[Blueprint]:
        """Fetch a single blueprint by customer code and locale."""
        try:
            return self.db.query(Blueprint).filter(
                Blueprint.customer_code == customer_code,
                Blueprint.locale == locale
            ).first()
        except Exception as e:
            logger.error(f"Database query failed for customer {customer_code} and locale {locale}: {e}")
            return None

    def get_customer_variants(self, customer_code: str) -> List[Blueprint]:
        """Fetch all locale variants of a blueprint for a customer."""
        try:
            return self.db.query(Blueprint).filter(
                Blueprint.customer_code == customer_code
            ).all()
        except Exception as e:
            logger.error(f"Database query failed for customer variants of {customer_code}: {e}")
            return []

    def get_all_blueprints(self) -> List[Blueprint]:
        """Fetch all blueprints in the database."""
        try:
            return self.db.query(Blueprint).all()
        except Exception as e:
            logger.error(f"Database query failed for all blueprints: {e}")
            return []

    def delete_blueprint(self, customer_code: str, locale: str = "KH") -> bool:
        """Delete a blueprint by customer code and locale."""
        try:
            row = self.get_blueprint(customer_code, locale)
            if row:
                self.db.delete(row)
                self.db.commit()
                return True
            return False
        except Exception as e:
            self.db.rollback()
            logger.error(f"Database deletion failed for customer {customer_code} and locale {locale}: {e}")
            raise e

    def save_blueprint(
        self,
        customer_code: str,
        locale: str,
        config_data: dict,
        template_json_data: dict,
        xlsx_bytes: bytes,
        filename: str
    ) -> Blueprint:
        """
        Saves or updates a blueprint template and configuration in the database.
        """
        try:
            config_str = json.dumps(config_data, ensure_ascii=False)
            template_str = json.dumps(template_json_data, ensure_ascii=False)
            description = config_data.get("_meta", {}).get("description", f"Generated blueprint for {customer_code}_{locale}")

            existing = self.get_blueprint(customer_code, locale)

            if existing:
                existing.description = description
                existing.config_json = config_str
                existing.template_json = template_str
                if existing.template_binary:
                    existing.template_binary.filename = filename
                    existing.template_binary.xlsx_blob = xlsx_bytes
                else:
                    existing.template_binary = BlueprintTemplate(
                        filename=filename,
                        xlsx_blob=xlsx_bytes
                    )
                blueprint = existing
            else:
                blueprint = Blueprint(
                    customer_code=customer_code,
                    locale=locale,
                    description=description,
                    config_json=config_str,
                    template_json=template_str
                )
                blueprint.template_binary = BlueprintTemplate(
                    filename=filename,
                    xlsx_blob=xlsx_bytes
                )
                self.db.add(blueprint)

            self.db.commit()
            return blueprint
        except Exception as e:
            self.db.rollback()
            logger.error(f"Database save failed for customer {customer_code} and locale {locale}: {e}")
            raise e
