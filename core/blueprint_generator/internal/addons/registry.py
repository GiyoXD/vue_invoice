from typing import Dict, Type
from core.blueprint_generator.internal.scanner.models.addons import BaseAddonFact
from .base import BaseAddonBuilder

class AddonRegistry:
    """Registry to map Addon facts to their respective builders."""
    
    _builders: Dict[str, BaseAddonBuilder] = {}

    @classmethod
    def register(cls, fact_type: str, builder: BaseAddonBuilder) -> None:
        cls._builders[fact_type] = builder
        
    @classmethod
    def get_builder(cls, fact: BaseAddonFact) -> BaseAddonBuilder:
        builder = cls._builders.get(fact.fact_type)
        if not builder:
            raise ValueError(f"No AddonBuilder registered for fact_type: {fact.fact_type}")
        return builder
