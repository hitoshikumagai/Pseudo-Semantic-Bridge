from typing import Callable, Dict
from src.schema.definitions import ProcessorType

# Registry
_PROCESSOR_REGISTRY: Dict[str, Callable] = {}

def register_processor(processor_id: str):
    def decorator(func: Callable):
        _PROCESSOR_REGISTRY[processor_id] = func
        return func
    return decorator

def get_processor(processor_id: str) -> Callable:
    if processor_id not in _PROCESSOR_REGISTRY:
        # Lazy-load: import modules if not loaded yet
        # (Simpler approach: import all here)
        try:
            import src.catalog.handlers.basic
            import src.catalog.handlers.document
            import src.catalog.handlers.archive
            import src.catalog.workflows.mail_router  # important: workflows
            import src.catalog.agents  # register agents
        except ImportError as e:
            print(f"⚠️ Import Warning: {e}")

    # Re-check
    if processor_id in _PROCESSOR_REGISTRY:
        return _PROCESSOR_REGISTRY[processor_id]
    
    raise KeyError(f"Logic ID '{processor_id}' is not implemented.")
