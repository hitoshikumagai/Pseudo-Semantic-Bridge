import os
import json # pandas の代わりに json を使う
from src.catalog import register_processor

from src.catalog.handlers.document import pdf_to_text_ocr
from src.catalog.handlers.basic import save_only
from src.catalog.handlers.archive import unzip_file

PROCESSOR_MAP = {
    "ocr_process": pdf_to_text_ocr,
    "save_process": save_only,
    "unzip_process": unzip_file
}

@register_processor("mail_workflow")
def mail_workflow(*args, **kwargs):
    item = args[0]
    output_dir = args[1]
    params = args[2] if len(args) > 2 else kwargs.get("params", {})
    
    # JSONファイルのパスを受け取る
    rule_file_path = params.get("rule_file")
    
    if not rule_file_path or not os.path.exists(rule_file_path):
        print(f"      ⚠️ ルールファイルが見つかりません: {rule_file_path}")
        return

    # ★ JSONとして読み込む (高速！)
    try:
        with open(rule_file_path, 'r', encoding='utf-8') as f:
            rules_list = json.load(f)
    except Exception as e:
        print(f"      ❌ ルール読み込みエラー(JSON): {e}")
        return

    mail_subject = item.name
    matched_rule = None

    # リストをループしてチェック
    for rule in rules_list:
        f_filter = str(rule.get("subject_filter", "*"))
        
        if f_filter == "*" or f_filter in mail_subject:
            matched_rule = rule
            # print(f"      🔍 ルール適合: {f_filter}")
            break
    
    if matched_rule is None:
        return

    # 以下、ロジックは同じ
    task_name = matched_rule["task_name"]
    action_id = matched_rule["action_id"]
    
    # JSONの boolean はそのまま Python の bool になるので変換不要だが念のため
    require_attachment = matched_rule.get("require_attachment")
    if isinstance(require_attachment, str):
        require_attachment = require_attachment.lower() == "true"

    print(f"🔄 [Workflow: {task_name}] Check: {mail_subject}")

    final_output_dir = os.path.join(output_dir, task_name)
    os.makedirs(final_output_dir, exist_ok=True)

    children = item.get_children()
    has_attachment = len(children) > 0

    if require_attachment and not has_attachment:
        # print(f"      ⏭️  Skipping: 添付必須")
        return

    if has_attachment:
        target_func = PROCESSOR_MAP.get(action_id, save_only)
        for child in children:
            target_func(child, final_output_dir, params)
    else:
        item.save_to(final_output_dir)
        print(f"      📝 本文保存: {item.name}")