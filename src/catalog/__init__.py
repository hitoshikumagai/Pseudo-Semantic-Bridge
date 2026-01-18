from typing import Callable, Dict, Any

# ロジックを格納する辞書（シングルトン）
_PROCESSOR_REGISTRY: Dict[str, Callable] = {}

def register_processor(processor_id: str):
    """関数をカタログに登録するためのデコレータ"""
    def decorator(func: Callable):
        if processor_id in _PROCESSOR_REGISTRY:
            raise ValueError(f"ID '{processor_id}' は既に登録されています。")
        _PROCESSOR_REGISTRY[processor_id] = func
        return func
    return decorator

def get_processor(processor_id: str) -> Callable:
    """エンジンがロジックを取り出すための関数"""
    logic = _PROCESSOR_REGISTRY.get(processor_id)
    if not logic:
        # 未実装の場合は安全策としてエラーにするか、デフォルトを返す
        raise KeyError(f"ロジックID '{processor_id}' は実装されていません。")
    return logic