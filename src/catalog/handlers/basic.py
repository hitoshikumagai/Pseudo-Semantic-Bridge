import os
from src.catalog import register_processor
from src.schema.definitions import ProcessorType
from src.adapter.outlook import AttachmentWrapper

@register_processor(ProcessorType.SAVE_ONLY)
def logic_save_only(attachment: AttachmentWrapper, output_dir: str):
    abs_output_dir = os.path.abspath(output_dir)
    save_path = os.path.join(abs_output_dir, attachment.filename)
    
    attachment.save_as(save_path)
    print(f"    [保存] {attachment.filename} を保存しました。")