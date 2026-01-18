from enum import Enum
from typing import List, Dict, Any
from pydantic import BaseModel, Field

class ProcessorType(str, Enum):
    SAVE_ONLY = "save_only"
    PDF_OCR = "pdf_to_text_ocr"
    UNZIP = "unzip_file"
    MAIL_WORKFLOW = "mail_workflow"

class AttachmentRule(BaseModel):
    extension: str
    processor_id: ProcessorType
    # ★New: 振る舞いを制御するためのパラメータ (デフォルトは空の辞書)
    parameters: Dict[str, Any] = Field(default_factory=dict)

class OutlookConfig(BaseModel):
    job_name: str
    version: str = "2.0"  # バージョンアップ
    domain: str
    search_keywords: List[str]
    destination_path: str
    rules: List[AttachmentRule]