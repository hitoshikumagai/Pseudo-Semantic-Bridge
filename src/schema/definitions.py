from enum import Enum
from typing import List, Dict, Any
from pydantic import BaseModel, Field, model_validator

class ProcessorType(str, Enum):
    SAVE_ONLY = "save_only"
    PDF_OCR = "pdf_to_text_ocr"
    UNZIP = "unzip_file"
    MAIL_WORKFLOW = "mail_workflow"

class AttachmentRule(BaseModel):
    extension: str
    # Use string to allow extension beyond fixed ProcessorType
    processor_id: str
    # Parameters to control behavior (defaults to empty dict)
    parameters: Dict[str, Any] = Field(default_factory=dict)

class OutlookConfig(BaseModel):
    job_name: str
    version: str = "2.0"  # Version bump
    domain: str
    search_keywords: List[str]
    destination_path: str
    rules: List[AttachmentRule]


class SpecificationStep(BaseModel):
    id: str = Field(min_length=1)
    action: str = Field(min_length=1)
    params: Dict[str, Any] = Field(default_factory=dict)


class SpecificationVerification(BaseModel):
    required_fields: List[str] = Field(default_factory=list)
    min_quality_score: float = Field(ge=0.0, le=1.0, default=0.8)


class SpecificationFallback(BaseModel):
    on_failure: str = "route_manual_review"


class IntentSpecification(BaseModel):
    spec_id: str = Field(min_length=1)
    spec_version: str = "1.0"
    domain: str = Field(min_length=1)
    intent: str = Field(min_length=1)
    inputs: Dict[str, Any] = Field(default_factory=dict)
    steps: List[SpecificationStep] = Field(min_length=1)
    verification: SpecificationVerification = Field(default_factory=SpecificationVerification)
    fallback: SpecificationFallback = Field(default_factory=SpecificationFallback)

    @model_validator(mode="after")
    def validate_unique_step_ids(self):
        step_ids = [step.id for step in self.steps]
        if len(step_ids) != len(set(step_ids)):
            raise ValueError("step ids must be unique")
        return self
