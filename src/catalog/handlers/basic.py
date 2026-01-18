import os
from src.catalog import register_processor
# ❌ from src.adapter.outlook import AttachmentWrapper <-- これを消す！

@register_processor("save_only")
def save_only(*args, **kwargs):
    item = args[0]
    output_dir = args[1]
    
    try:
        item.save_to(output_dir)
        # print(f"      (Child) 💾 保存完了: {item.name}")
    except Exception as e:
        print(f"      ❌ Save Error: {e}")