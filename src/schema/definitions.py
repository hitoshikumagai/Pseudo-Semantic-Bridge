from enum import Enum
from typing import List
from pydantic import BaseModel

class ProcessorType(str, Enum):
    SAVE_ONLY = "save_only"       # そのまま保存
    PDF_OCR = "pdf_to_text_ocr"   # PDFをOCR処理
    UNZIP = "unzip_file"          # ZIP解凍

class AttachmentRule(BaseModel):
    extension: str
    processor_id: ProcessorType

class OutlookConfig(BaseModel):
    job_name: str
    search_keywords: List[str]
    destination_path: str
    rules: List[AttachmentRule]