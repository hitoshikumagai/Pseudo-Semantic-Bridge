import os
from src.catalog import register_processor
# ❌ from src.adapter.outlook import AttachmentWrapper <-- これを消す！

@register_processor("unzip_file")
def unzip_file(*args, **kwargs):
    item = args[0]
    output_dir = args[1]
    params = args[2] if len(args) > 2 else kwargs.get("params", {})
    mode = params.get("mode", "auto")
    
    try:
        saved_path = item.save_to(output_dir)
        print(f"      (Child) 📦 ZIP解凍: {os.path.basename(saved_path)} (Mode: {mode})")
    except Exception as e:
        print(f"      ❌ Zip Error: {e}")