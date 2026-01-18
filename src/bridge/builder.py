import os
import json
import pandas as pd
from src.bridge.excel_parser import parse_excel_spec

# =========================================================
# 🏗️ Internal Compilation Logic
# =========================================================

def _compile_system_spec(excel_path: str, json_out_path: str):
    """
    システム仕様書 (invoice_bot_v2.xlsx) のコンパイル
    Returns: Pydantic Model -> JSON File
    """
    print(f"📖 System Spec解析: {excel_path}")
    
    # 1. Excelを解析してオブジェクトを取得 (OutlookConfig Object)
    spec_data = parse_excel_spec(excel_path)
    
    if spec_data:
        os.makedirs(os.path.dirname(json_out_path), exist_ok=True)
        with open(json_out_path, "w", encoding='utf-8') as f:
            
            # ★ ここが修正ポイント ★
            # オブジェクトが Pydantic モデルの場合、専用メソッドで JSON 化する
            if hasattr(spec_data, "model_dump_json"):
                # Pydantic v2用
                f.write(spec_data.model_dump_json(indent=2))
            elif hasattr(spec_data, "json"):
                # Pydantic v1用 (互換性維持)
                f.write(spec_data.json(indent=2))
            else:
                # ただの辞書(dict)なら標準ライブラリでOK
                json.dump(spec_data, f, indent=2, ensure_ascii=False)
                
        print(f"    ✅ JSON Saved: {json_out_path}")
    else:
        print(f"    ❌ Spec Data is Empty: {excel_path}")

def _compile_business_rules(excel_path: str, json_out_path: str):
    """
    業務ルール (mail_business_rules.xlsx) のコンパイル
    Returns: Pandas DataFrame -> List[Dict] -> JSON File
    """
    print(f"📖 Business Rule解析: {excel_path}")
    
    if not os.path.exists(excel_path):
        print(f"    ⚠️ File Not Found (Skip): {excel_path}")
        return

    try:
        # Excel読込
        df = pd.read_excel(excel_path)
        
        # NaN (空欄) を None に置換してJSONでエラーにならないようにする
        df = df.where(pd.notnull(df), None)
        
        # 辞書のリストに変換
        rules_list = df.to_dict(orient='records')
        
        os.makedirs(os.path.dirname(json_out_path), exist_ok=True)
        with open(json_out_path, 'w', encoding='utf-8') as f:
            json.dump(rules_list, f, indent=2, ensure_ascii=False)
            
        print(f"    ✅ Rules JSON Saved: {json_out_path}")
        
    except Exception as e:
        print(f"    ⚠️ Rule Compile Error (Skip): {e}")

# =========================================================
# 🚀 Public Facade (Mainから呼ぶのはこれだけ)
# =========================================================

def build_all_configs():
    """
    プロジェクト内のすべてのExcel仕様書を探してJSONに変換する
    """
    print("🏗️  Building Configurations...")

    # 1. システム仕様書 (Botの基本設定)
    _compile_system_spec(
        "specs/accounting/invoice_bot_v2.xlsx", 
        "configs/accounting/invoice_bot_v2.json"
    )

    # 2. 業務ルール (メール振り分け設定)
    _compile_business_rules(
        "specs/accounting/mail_business_rules.xlsx",
        "configs/accounting/mail_business_rules.json"
    )