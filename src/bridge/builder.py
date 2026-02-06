import os
import json
import pandas as pd
from src.bridge.excel_parser import parse_excel_spec

# =========================================================
# Internal Compilation Logic
# =========================================================

def _compile_system_spec(excel_path: str, json_out_path: str):
    """
    Compile system spec (invoice_bot_v2.xlsx).
    Returns: Pydantic Model -> JSON File
    """
    print(f"📖 System Spec parse: {excel_path}")
    
    # Parse Excel into OutlookConfig
    spec_data = parse_excel_spec(excel_path)
    
    if spec_data:
        os.makedirs(os.path.dirname(json_out_path), exist_ok=True)
        with open(json_out_path, "w", encoding='utf-8') as f:
            
            # Serialize Pydantic models with the right method
            if hasattr(spec_data, "model_dump_json"):
                # Pydantic v2
                f.write(spec_data.model_dump_json(indent=2))
            elif hasattr(spec_data, "json"):
                # Pydantic v1
                f.write(spec_data.json(indent=2))
            else:
                # Fallback for dict
                json.dump(spec_data, f, indent=2, ensure_ascii=False)
                
        print(f"    ✅ JSON Saved: {json_out_path}")
    else:
        print(f"    ❌ Spec Data is Empty: {excel_path}")

def _compile_business_rules(excel_path: str, json_out_path: str):
    """
    Compile business rules (mail_business_rules.xlsx).
    Returns: Pandas DataFrame -> List[Dict] -> JSON File
    """
    print(f"📖 Business Rule parse: {excel_path}")
    
    if not os.path.exists(excel_path):
        print(f"    ⚠️ File Not Found (Skip): {excel_path}")
        return

    try:
        # Read Excel
        df = pd.read_excel(excel_path)
        
        # Replace NaN with None for JSON
        df = df.where(pd.notnull(df), None)
        
        # Convert to list[dict]
        rules_list = df.to_dict(orient='records')
        
        os.makedirs(os.path.dirname(json_out_path), exist_ok=True)
        with open(json_out_path, 'w', encoding='utf-8') as f:
            json.dump(rules_list, f, indent=2, ensure_ascii=False)
            
        print(f"    ✅ Rules JSON Saved: {json_out_path}")
        
    except Exception as e:
        print(f"    ⚠️ Rule Compile Error (Skip): {e}")

# =========================================================
# Public Facade (Main calls this)
# =========================================================

def build_all_configs():
    """
    Build all configs from Excel specs in the project.
    """
    print("🏗️  Building Configurations...")

    # 1. System spec (bot settings)
    _compile_system_spec(
        "specs/accounting/invoice_bot_v2.xlsx", 
        "configs/accounting/invoice_bot_v2.json"
    )

    # 2. Business rules (mail routing)
    _compile_business_rules(
        "specs/accounting/mail_business_rules.xlsx",
        "configs/accounting/mail_business_rules.json"
    )
