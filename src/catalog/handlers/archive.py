import os
from src.catalog import register_processor
from src.schema.definitions import ProcessorType
from src.adapter.outlook import AttachmentWrapper

@register_processor(ProcessorType.UNZIP)
def logic_unzip(attachment: AttachmentWrapper, output_dir: str):
    abs_output_dir = os.path.abspath(output_dir)
    
    # 1. 保存
    save_path = os.path.join(abs_output_dir, attachment.filename)
    attachment.save_as(save_path)
    
    # 2. 解凍シミュレーション
    unzip_dir = os.path.join(abs_output_dir, "unzipped_" + attachment.filename)
    if not os.path.exists(unzip_dir):
        os.makedirs(unzip_dir)
        
    print(f"    [Unzip] {attachment.filename} を解凍しました -> {unzip_dir}")