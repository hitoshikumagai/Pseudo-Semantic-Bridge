from typing import Callable, Dict
from src.schema.definitions import ProcessorType

# レジストリ
_PROCESSOR_REGISTRY: Dict[str, Callable] = {}

def register_processor(processor_id: str):
    def decorator(func: Callable):
        _PROCESSOR_REGISTRY[processor_id] = func
        return func
    return decorator

def get_processor(processor_id: str) -> Callable:
    if processor_id not in _PROCESSOR_REGISTRY:
        # 遅延ロード: まだ読み込まれていないモジュールがあればここでimportする
        # (簡易的に全て読み込んでしまうのが楽です)
        try:
            import src.catalog.handlers.basic
            import src.catalog.handlers.document
            import src.catalog.handlers.archive
            import src.catalog.workflows.mail_router # ★ここが重要
        except ImportError as e:
            print(f"⚠️ Import Warning: {e}")

    # 再チェック
    if processor_id in _PROCESSOR_REGISTRY:
        return _PROCESSOR_REGISTRY[processor_id]
    
    raise KeyError(f"ロジックID '{processor_id}' は実装されていません。")