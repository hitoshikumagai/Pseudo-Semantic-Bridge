import os
from src.catalog import register_processor
from src.schema.definitions import ProcessorType
from src.adapter.outlook import AttachmentWrapper

@register_processor(ProcessorType.PDF_OCR)
def logic_pdf_ocr(attachment: AttachmentWrapper, output_dir: str):
    abs_output_dir = os.path.abspath(output_dir)
    
    # 1. 保存
    save_path = os.path.join(abs_output_dir, attachment.filename)
    attachment.save_as(save_path)
    
    # 2. OCRシミュレーション
    txt_filename = attachment.filename + ".txt"
    txt_path = os.path.join(abs_output_dir, txt_filename)
    
    with open(txt_path, "w", encoding="utf-8") as f:
        f.write(f"【OCR済み】{attachment.filename}\n金額: 10,000円")
        
    print(f"    [OCR ] {attachment.filename} をテキスト化しました -> {txt_filename}")