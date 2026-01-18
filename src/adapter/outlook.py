import os
import shutil
import tempfile
import win32com.client
from typing import List
from .base import BaseAdapter, UnifiedItem

class OutlookItem(UnifiedItem):
    """Outlook固有の事情（COM操作など）を吸収するラッパー"""
    
    @property
    def name(self) -> str:
        # Subjectがあればそれを、なければFileNameを、それもなければUnknown
        return getattr(self._raw_item, "Subject", None) or getattr(self._raw_item, "FileName", "Unknown")

    @property
    def extension(self) -> str:
        # コンテナ（メール本体）なら .msg
        if self.is_container:
            return ".msg"
        
        # 添付ファイルならファイル名から拡張子を抽出
        fname = self.name
        _, ext = os.path.splitext(fname)
        return ext.lower() if ext else ""

    @property
    def is_container(self) -> bool:
        # Class 43 = MailItem, Class 2 = Contact, etc.
        # Attachmentsプロパティを持っていて、かつ添付ファイルオブジェクトではないものをコンテナとみなす
        # (簡易判定: Subjectがあるならメール本体とみなす)
        return hasattr(self._raw_item, "Subject") and hasattr(self._raw_item, "Attachments")

    def get_children(self) -> List['UnifiedItem']:
        children = []
        if self.is_container:
            try:
                for att in self._raw_item.Attachments:
                    children.append(OutlookItem(att))
            except Exception as e:
                print(f"      ⚠️ Failed to get children: {e}")
        return children

    def save_to(self, directory: str) -> str:
        """
        Tempリレー方式を用いて、指定ディレクトリに確実に保存する
        """
        try:
            # 1. 保存先を絶対パス化
            abs_dir = os.path.abspath(directory)
            os.makedirs(abs_dir, exist_ok=True)
            
            filename = self.name
            # ファイル名に使えない文字を除去（簡易版）
            invalid_chars = '<>:"/\\|?*'
            for char in invalid_chars:
                filename = filename.replace(char, '_')

            final_path = os.path.join(abs_dir, filename)

            # --- メール本体(.msg)として保存する場合 ---
            if self.is_container:
                # Outlookの仕様上、メール自体のSaveAsは絶対パス必須
                # Type 3 = olMSG
                self._raw_item.SaveAs(final_path, 3)
                print(f"      (System) 📧 メール保存完了: {filename}")
                return final_path

            # --- 添付ファイルとして保存する場合 (Tempリレー) ---
            # Tempに保存
            temp_dir = tempfile.gettempdir()
            temp_path = os.path.join(temp_dir, filename)
            
            # クリーンアップ
            if os.path.exists(temp_path):
                try: os.remove(temp_path)
                except: pass

            # 保存実行
            if hasattr(self._raw_item, "save_as"): # Wrapper対応
                self._raw_item.save_as(temp_path)
            elif hasattr(self._raw_item, "SaveAsFile"): # Raw COM対応
                self._raw_item.SaveAsFile(temp_path)
            else:
                raise Exception("保存メソッドが見つかりません")

            # 移動
            shutil.move(temp_path, final_path)
            print(f"      (System) 💾 ファイル保存完了: {filename}")
            return final_path

        except Exception as e:
            print(f"      ❌ Save Error ({self.name}): {e}")
            raise e

class OutlookAdapter(BaseAdapter):
    def __init__(self):
        try:
            self.outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            print(">> [Adapter] Outlookに接続しました")
        except Exception as e:
            print(f"❌ Outlook接続エラー: {e}")
            self.outlook = None

    def fetch_items(self, keyword: str) -> List[UnifiedItem]:
        if not self.outlook:
            return []
            
        folder = self.outlook.GetDefaultFolder(6) # Inbox
        # 検索 (Subjectのみ)
        try:
            items = folder.Items.Restrict(f"@SQL=\"urn:schemas:httpmail:subject\" like '%{keyword}%'")
            results = []
            for item in items:
                results.append(OutlookItem(item))
            return results
        except Exception as e:
            print(f"⚠️ Search Error: {e}")
            return []