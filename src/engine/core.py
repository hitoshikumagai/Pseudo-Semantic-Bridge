import os
from src.schema.definitions import OutlookConfig
# Adapterは抽象クラスとして受け取るのが理想だが、便宜上型ヒント等は省略
from src.catalog import get_processor

class GenericEtlEngine:
    def __init__(self, config: OutlookConfig, adapter):
        self.config = config
        self.adapter = adapter

    def run(self):
        print(f"🚀 Engine Start: {self.config.job_name} (v{self.config.version})")
        
        for keyword in self.config.search_keywords:
            items = self.adapter.fetch_items(keyword)
            print(f">> [Adapter] 検索 '{keyword}': {len(items)} 件ヒット")

            for item in items:
                self._process_recursive(item)
        
        print("✅ Engine Finished.")

    def _process_recursive(self, item):
        """
        UnifiedItem を受け取り、再帰的に処理する
        """
        # 1. ルール適合チェック
        # アイテムの拡張子 (.msg, .pdf 等) を見て、対応するルールがあれば実行
        rule_executed = self._try_execute_rule(item)
        
        # 2. ルールが実行されたら、そのハンドラーに全権委任（子要素処理はハンドラー次第）
        if rule_executed:
            return

        # 3. ルールが実行されず、かつコンテナ（フォルダ/メール）なら、自動で中身を掘る
        if item.is_container:
            # print(f"   📂 Opening Container: {item.name}")
            for child in item.get_children():
                self._process_recursive(child)

    def _try_execute_rule(self, item) -> bool:
        """
        アイテムの拡張子を見て、適合するルールがあれば実行する
        """
        target_rule = None
        # UnifiedItemから拡張子を取得（例: .pdf, .msg）
        ext = item.extension.lower()
        
        # 設定(rules)から、この拡張子に対応するルールを探す
        for rule in self.config.rules:
            if rule.extension.lower() == ext:
                target_rule = rule
                break
        
        # ルールがなければ何もしない
        if not target_rule:
            return False

        # --- 修正ポイント: Enum対策 ---
        # processor_id が Enum(ProcessorType) の場合と、文字列の場合があるため吸収する
        # (JSONから読み込んだ場合は文字列、コード定義の場合はEnumの可能性がある)
        raw_id = target_rule.processor_id
        
        # Enumなら .value ("mail_workflow") を取り出し、文字列ならそのまま使う
        processor_id = raw_id.value if hasattr(raw_id, "value") else raw_id

        # ログ出力 (Enumではなく変換後のIDを表示)
        print(f"   ⚙️  Running Rule [{processor_id}] for: {item.name} ({ext})")

        try:
            # IDに対応する関数（Handler/Workflow）を取得
            # ここで文字列の "mail_workflow" などが渡されるので KeyError にならない
            handler = get_processor(processor_id)
            
            # 実行 (UnifiedItemをそのまま渡す)
            handler(
                item, 
                self.config.destination_path, 
                target_rule.parameters
            )
            return True
        except Exception as e:
            print(f"   ❌ Engine Error: {e}")
            import traceback
            traceback.print_exc()
            return False