from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timezone
import time
from uuid import uuid4

from src.schema.definitions import OutlookConfig
from src.catalog import get_processor
from src.telemetry.run_logger import append_run

class GenericEtlEngine:
    def __init__(self, config: OutlookConfig, adapter):
        self.config = config
        self.adapter = adapter
        self._executor = self._build_executor()
        self._futures = []

    def _build_executor(self):
        max_workers = 1
        for rule in self.config.rules:
            params = rule.parameters or {}
            try:
                candidate = int(params.get("max_concurrency", 1))
            except Exception:
                candidate = 1
            if candidate > max_workers:
                max_workers = candidate

        if max_workers > 1:
            return ThreadPoolExecutor(max_workers=max_workers)
        return None

    def run(self):
        print(f"🚀 Engine Start: {self.config.job_name} (v{self.config.version})")
        
        for keyword in self.config.search_keywords:
            items = self.adapter.fetch_items(keyword)
            print(f">> [Adapter] Search '{keyword}': {len(items)} items")

            for item in items:
                self._process_recursive(item)

        if self._executor:
            for future in self._futures:
                try:
                    future.result()
                except Exception as e:
                    print(f"   ❌ Engine Error (async): {e}")
            self._executor.shutdown(wait=True)

        print("✅ Engine Finished.")

    def _process_recursive(self, item):
        """
        Process a UnifiedItem recursively.
        """
        # 1. Try rule by extension (e.g., .msg, .pdf).
        rule_executed = self._try_execute_rule(item)
        
        # 2. If a rule ran, delegate all further processing to that handler.
        if rule_executed:
            return

        # 3. If no rule ran and this is a container, drill into children.
        if item.is_container:
            for child in item.get_children():
                self._process_recursive(child)

    def _try_execute_rule(self, item) -> bool:
        """
        Execute a rule that matches the item's extension.
        """
        target_rule = None
        # Read extension from item (e.g., .pdf, .msg).
        ext = item.extension.lower()
        
        # Find a rule for the extension.
        for rule in self.config.rules:
            if rule.extension.lower() == ext:
                target_rule = rule
                break
        
        # No rule, no-op.
        if not target_rule:
            return False

        # Normalize processor_id: Enum -> value, string -> as-is.
        raw_id = target_rule.processor_id
        
        processor_id = raw_id.value if hasattr(raw_id, "value") else raw_id

        print(f"   ⚙️  Running Rule [{processor_id}] for: {item.name} ({ext})")

        def build_record():
            timestamp = datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")
            run_id = f"{int(time.time() * 1000)}-{uuid4().hex[:8]}"
            has_attachment = False
            if item.is_container:
                try:
                    has_attachment = len(item.get_children()) > 0
                except Exception:
                    has_attachment = False
            return {
                "run_id": run_id,
                "timestamp": timestamp,
                "workflow": "engine",
                "processor_id": processor_id,
                "action_id": processor_id,
                "input": {
                    "subject": item.name,
                    "has_attachment": has_attachment,
                    "attachment_ext": ext,
                },
                "result": {"status": None, "output_path": None, "error": None},
                "quality": {
                    "label": None,
                    "score": None,
                    "notes": None,
                    "feedback_by": None,
                    "feedback_at": None,
                },
            }

        def run_and_log():
            record = build_record()
            try:
                handler(
                    item,
                    self.config.destination_path,
                    params,
                )
                record["result"]["status"] = "success"
                append_run(record)
            except Exception as e:
                record["result"]["status"] = "error"
                record["result"]["error"] = str(e)
                append_run(record)
                raise

        try:
            handler = get_processor(processor_id)
            params = target_rule.parameters or {}
            max_concurrency = int(params.get("max_concurrency", 1)) if params else 1
            log_success = processor_id != "mail_workflow"

            if self._executor and max_concurrency > 1:
                if log_success:
                    future = self._executor.submit(run_and_log)
                else:
                    future = self._executor.submit(
                        handler,
                        item,
                        self.config.destination_path,
                        params,
                    )
                self._futures.append(future)
            else:
                if log_success:
                    run_and_log()
                else:
                    handler(
                        item,
                        self.config.destination_path,
                        params,
                    )
            return True
        except Exception as e:
            print(f"   ❌ Engine Error: {e}")
            import traceback
            traceback.print_exc()
            if processor_id == "mail_workflow":
                record = build_record()
                record["result"]["status"] = "error"
                record["result"]["error"] = str(e)
                append_run(record)
            return False
