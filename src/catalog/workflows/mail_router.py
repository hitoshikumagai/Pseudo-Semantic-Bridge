import os
import json
import time
from datetime import datetime, timezone
from pathlib import Path
from uuid import uuid4

from src.catalog import register_processor
from src.telemetry.run_logger import append_run

from src.catalog.handlers.document import pdf_to_text_ocr
from src.catalog.handlers.basic import save_only
from src.catalog.handlers.archive import unzip_file

PROCESSOR_MAP = {
    "ocr_process": pdf_to_text_ocr,
    "save_process": save_only,
    "unzip_process": unzip_file
}

def _resolve_rule_file_path(rule_file_path: str) -> str:
    if not rule_file_path:
        return rule_file_path
    candidate = Path(rule_file_path)
    if candidate.is_absolute() and candidate.exists():
        return str(candidate)
    # Try current working directory
    cwd_candidate = (Path.cwd() / candidate).resolve()
    if cwd_candidate.exists():
        return str(cwd_candidate)
    # Try repo root inferred from this file: src/catalog/workflows/mail_router.py
    try:
        repo_root = Path(__file__).resolve().parents[3]
        repo_candidate = (repo_root / candidate).resolve()
        if repo_candidate.exists():
            return str(repo_candidate)
    except Exception:
        pass
    return rule_file_path

@register_processor("mail_workflow")
def mail_workflow(*args, **kwargs):
    item = args[0]
    output_dir = args[1]
    params = args[2] if len(args) > 2 else kwargs.get("params", {})
    
    # Read rule file path
    rule_file_path = params.get("rule_file")
    rule_file_path = _resolve_rule_file_path(rule_file_path)
    
    if not rule_file_path or not os.path.exists(rule_file_path):
        print(f"      ⚠️ Rule file not found: {rule_file_path}")
        return

    # Load as JSON
    try:
        with open(rule_file_path, 'r', encoding='utf-8') as f:
            rules_list = json.load(f)
    except Exception as e:
        print(f"      ❌ Rule load error (JSON): {e}")
        return

    mail_subject = item.name
    matched_rule = None

    # Find first matching rule
    for rule in rules_list:
        f_filter = str(rule.get("subject_filter", "*"))
        
        if f_filter == "*" or f_filter in mail_subject:
            matched_rule = rule
            break
    
    if matched_rule is None:
        return

    task_name = matched_rule["task_name"]
    action_id = matched_rule["action_id"]
    
    require_attachment = matched_rule.get("require_attachment")
    if isinstance(require_attachment, str):
        require_attachment = require_attachment.lower() == "true"

    print(f"🔄 [Workflow: {task_name}] Check: {mail_subject}")

    final_output_dir = os.path.join(output_dir, task_name)
    os.makedirs(final_output_dir, exist_ok=True)

    children = item.get_children()
    has_attachment = len(children) > 0

    if require_attachment and not has_attachment:
        return

    def log_run(attachment_ext: str, action_executed: str = None):
        timestamp = datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")
        run_id = f"{int(time.time() * 1000)}-{uuid4().hex[:8]}"
        record = {
            "run_id": run_id,
            "timestamp": timestamp,
            "workflow": "mail_workflow",
            "processor_id": "mail_workflow",
            "action_id": action_id,
            "input": {
                "subject": mail_subject,
                "has_attachment": has_attachment,
                "attachment_ext": attachment_ext,
            },
            "result": {
                "status": "success",
                "output_path": None,
                "error": None,
                "action_executed": action_executed,
            },
            "quality": {
                "label": None,
                "score": None,
                "notes": None,
                "feedback_by": None,
                "feedback_at": None,
            },
        }
        append_run(record)

    def get_child_ext(child) -> str:
        ext = getattr(child, "extension", None)
        if not ext:
            _, ext = os.path.splitext(getattr(child, "name", "") or "")
        return (ext or "").lower()

    if has_attachment:
        target_func = PROCESSOR_MAP.get(action_id, save_only)
        for child in children:
            try:
                target_func(child, final_output_dir, params)
                log_run(get_child_ext(child), action_executed=action_id)
            except Exception as e:
                error_record = {
                    "run_id": f"{int(time.time() * 1000)}-{uuid4().hex[:8]}",
                    "timestamp": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
                    "workflow": "mail_workflow",
                    "processor_id": "mail_workflow",
                    "action_id": action_id,
                    "input": {
                        "subject": mail_subject,
                        "has_attachment": has_attachment,
                        "attachment_ext": get_child_ext(child),
                    },
                    "result": {"status": "error", "output_path": None, "error": str(e)},
                    "quality": {
                        "label": None,
                        "score": None,
                        "notes": None,
                        "feedback_by": None,
                        "feedback_at": None,
                    },
                }
                append_run(error_record)
    else:
        try:
            item.save_to(final_output_dir)
            print(f"      📝 Saved body: {item.name}")
            log_run("", action_executed="save_only")
        except Exception as e:
            error_record = {
                "run_id": f"{int(time.time() * 1000)}-{uuid4().hex[:8]}",
                "timestamp": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
                "workflow": "mail_workflow",
                "processor_id": "mail_workflow",
                "action_id": action_id,
                "input": {
                    "subject": mail_subject,
                    "has_attachment": has_attachment,
                    "attachment_ext": "",
                },
                "result": {"status": "error", "output_path": None, "error": str(e)},
                "quality": {
                    "label": None,
                    "score": None,
                    "notes": None,
                    "feedback_by": None,
                    "feedback_at": None,
                },
            }
            append_run(error_record)
