import os
from src.schema.definitions import OutlookConfig, ProcessorType
from src.adapter.outlook import OutlookAdapter
from src.catalog import get_processor

class GenericEtlEngine:
    def __init__(self, config: OutlookConfig, adapter: OutlookAdapter):
        self.config = config
        self.adapter = adapter

    def run(self):
        # 0. 準備
        if not os.path.exists(self.config.destination_path):
            os.makedirs(self.config.destination_path)
        self.adapter.connect()

        # 1. 検索ループ
        for keyword in self.config.search_keywords:
            attachments = self.adapter.fetch_attachments(keyword)
            
            # 2. 処理ループ
            for attachment in attachments:
                processor_id = self._decide_processor(attachment.filename)
                
                # 3. カタログから関数を取得して実行
                try:
                    handler = get_processor(processor_id)
                    handler(attachment, self.config.destination_path)
                except Exception as e:
                    print(f"    !! 処理失敗 ({processor_id}): {e}")

    def _decide_processor(self, filename: str) -> str:
        ext = os.path.splitext(filename)[1].lower()
        
        for rule in self.config.rules:
            if rule.extension.lower() == ext:
                return rule.processor_id
        
        return ProcessorType.SAVE_ONLY