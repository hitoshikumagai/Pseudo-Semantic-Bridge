import win32com.client
from typing import Protocol, Any

class AttachmentWrapper:
    """COMオブジェクトを隠蔽し、保存機能だけを提供するラッパー"""
    def __init__(self, com_attachment: Any, filename: str):
        self._com = com_attachment
        self.filename = filename

    def save_as(self, path: str):
        self._com.SaveAsFile(path)

class OutlookAdapter:
    """ローカルのOutlook (win32com) を操作するクラス"""
    def __init__(self):
        self.outlook = None
        self.namespace = None

    def connect(self):
        try:
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
            print(">> [Adapter] Outlookに接続しました")
        except Exception as e:
            print(f"!! [Adapter] 接続失敗: {e}")
            raise

    def fetch_attachments(self, keyword: str):
        inbox = self.namespace.GetDefaultFolder(6) # 6 = Inbox
        filter_str = f"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{keyword}%'"
        items = inbox.Items.Restrict(filter_str)

        print(f">> [Adapter] 検索 '{keyword}': {items.Count} 件ヒット")

        for item in items:
            try:
                if item.Attachments.Count == 0:
                    continue
                
                print(f"  - メール発見: {item.Subject}")
                for attachment in item.Attachments:
                    yield AttachmentWrapper(attachment, attachment.FileName)
            except Exception as e:
                print(f"  !! エラー: {e}")