import os
from src.catalog import register_processor
# ❌ from src.adapter.outlook import AttachmentWrapper  <-- これを消す！

@register_processor("pdf_to_text_ocr")
def pdf_to_text_ocr(*args, **kwargs):
    # 新しい引数の受け取り方 (*args)
    item = args[0] # ここに来るのはもう Wrapper ではなく UnifiedItem です
    output_dir = args[1]
    params = args[2] if len(args) > 2 else kwargs.get("params", {})
    lang = params.get("lang", "eng")
    
    try:
        # UnifiedItem なので .save_to() が使えます
        saved_path = item.save_to(output_dir)
        filename = os.path.basename(saved_path)
        
        print(f"      (Child) 👁️ OCR処理: {filename} [Lang: {lang}]")
        # ここに OCR ロジック...
        
    except Exception as e:
        print(f"      ❌ OCR Error: {e}")