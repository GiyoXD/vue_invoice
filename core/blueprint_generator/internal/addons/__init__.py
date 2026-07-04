from .base import BaseAddonBuilder
from .leather import LeatherAddonBuilder
from .registry import AddonRegistry

# Initialize registry with all known builders
AddonRegistry.register("leather_summary", LeatherAddonBuilder())

__all__ = [
    "BaseAddonBuilder",
    "AddonRegistry"
]
