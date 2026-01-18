from abc import ABC, abstractmethod
from typing import List, Any

class UnifiedItem(ABC):
    """
    全てのデータソース（Outlook, SharePoint, FileSystem...）
    を同じように扱うための共通規格
    """
    def __init__(self, raw_item: Any):
        self._raw_item = raw_item

    @property
    @abstractmethod
    def name(self) -> str:
        """表示名（Subject または Filename）"""
        pass

    @property
    @abstractmethod
    def extension(self) -> str:
        """拡張子（.msg, .pdf, .xlsx...）"""
        pass

    @property
    @abstractmethod
    def is_container(self) -> bool:
        """フォルダやメールのように、中に子要素を持つか？"""
        pass

    @abstractmethod
    def get_children(self) -> List['UnifiedItem']:
        """子要素（添付ファイルなど）を取得"""
        pass

    @abstractmethod
    def save_to(self, directory: str) -> str:
        """指定フォルダに保存する（パス問題を解決済みであること）"""
        pass

class BaseAdapter(ABC):
    @abstractmethod
    def fetch_items(self, keyword: str) -> List[UnifiedItem]:
        pass