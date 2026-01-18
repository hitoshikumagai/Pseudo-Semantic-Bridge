import os
import logging
from src.catalog import get_processor
from src.schema.definitions import OutlookConfig, AttachmentRule, ProcessorType

# ログ設定（本番運用を見越してprintではなくlogger推奨だが、今回は分かりやすくprint併用）
logger = logging.getLogger(__name__)

class GenericEtlEngine:
    def __init__(self, config: OutlookConfig, adapter):
        self.config = config
        self.adapter = adapter

    def run(self):
        """ジョブの実行メインループ"""
        print(f"🚀 Engine Start: {self.config.job_name} (v{self.config.version})")
        
        # 0. 保存先の準備
        if not os.path.exists(self.config.destination_path):
            os.makedirs(self.config.destination_path)
            print(f"    📁 Created dir: {self.config.destination_path}")
        
        # 1. Outlook接続
        try:
            self.adapter.connect()
        except Exception as e:
            print(f"❌ Outlook接続エラー: {e}")
            return

        # 2. 検索ループ
        for keyword in self.config.search_keywords:
            # アダプターから添付ファイルを順次取得
            attachments = self.adapter.fetch_attachments(keyword)
            
            for attachment in attachments:
                self._process_single_attachment(attachment)

        print("✅ Engine Finished.")

    def _process_single_attachment(self, attachment):
        """添付ファイル1つに対する処理"""
        
        # A. ルール検索 (どのルールに当てはまるか？)
        rule = self._find_matching_rule(attachment.filename)
        
        if not rule:
            # マッチするルールがない場合はスキップ（またはデフォルト処理）
            print(f"    ⏭️ Skip: {attachment.filename} (No matching rule)")
            return

        # B. ロジック（関数）の取得
        handler = get_processor(rule.processor_id)
        if not handler:
            print(f"    ⚠️ Warning: 未実装のロジックIDです ({rule.processor_id})")
            return

        # C. ロジックの実行 (★v2.0: パラメータを渡す)
        try:
            print(f"    ⚙️ Executing {rule.processor_id} for {attachment.filename}...")
            
            # ハンドラには「添付ファイル」「保存先」「パラメータ(辞書)」の3つを渡す
            handler(
                attachment, 
                self.config.destination_path, 
                rule.parameters  # ★ここが重要！Excel/JSONに書かれた指示書を渡す
            )
            
        except TypeError as e:
            # 古いハンドラ（引数が2つしかない場合）への救済措置
            print(f"    ⚠️ v1互換モードで実行します: {e}")
            handler(attachment, self.config.destination_path)
        except Exception as e:
            print(f"    💥 Error in handler: {e}")

    def _find_matching_rule(self, filename: str) -> AttachmentRule:
        """
        ファイル名に基づいて適用すべきルールオブジェクトを返す
        """
        ext = os.path.splitext(filename)[1].lower()
        
        # 設定されたルールを上から順に走査
        for rule in self.config.rules:
            # 現在は拡張子の一致のみを見ているが、
            # 将来ここで「ファイル名正規表現」などの条件判定も追加可能
            if rule.extension.lower() == ext:
                return rule
        
        return None