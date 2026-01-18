import pandas as pd
import json
import os
import sys
from typing import List, Dict, Any

# プロジェクトルートにパスを通す（Jupyterなどで実行する場合の安全策）
if os.getcwd() not in sys.path:
    sys.path.append(os.getcwd())

from src.schema.definitions import OutlookConfig, AttachmentRule, ProcessorType

def parse_excel_spec(excel_path: str) -> OutlookConfig:
    """
    Excel仕様書を読み込み、OutlookConfigオブジェクトを生成する
    
    Args:
        excel_path (str): 仕様書Excelファイル(.xlsx)へのパス
        
    Returns:
        OutlookConfig: Pydanticでバリデーションされた設定オブジェクト
    """
    
    # 0. ファイル存在チェック
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"❌ 仕様書Excelが見つかりません: {excel_path}")

    print(f"📖 Excel仕様書を解析中...: {excel_path}")

    # =========================================================
    # 1. Settingsシート (基本設定) の読み込み
    # =========================================================
    try:
        # header=None: A列をKey, B列をValueとして読む
        df_settings = pd.read_excel(excel_path, sheet_name="Settings", header=None)
        # 辞書化: { "Job Name": "Invoice_Bot", ... }
        settings = dict(zip(df_settings[0], df_settings[1]))
    except Exception as e:
        raise ValueError(f"❌ Settingsシートの読み込みに失敗しました: {e}")

    # 必須項目の取得とデフォルト値（.get()を使用）
    # str()変換を入れるのは、Excelが数値を勝手にfloat等で返すのを防ぐため
    job_name = str(settings.get("Job Name", "Unnamed_Job")).strip()
    domain = str(settings.get("Domain", "common")).strip()
    destination = str(settings.get("Destination", "./data/output")).strip()
    
    # キーワード: "請求書, Invoice" -> ["請求書", "Invoice"]
    raw_keywords = str(settings.get("Keywords", ""))
    keywords = [k.strip() for k in raw_keywords.split(",") if k.strip()]

    # =========================================================
    # 2. Rulesシート (処理ルール) の読み込み
    # =========================================================
    try:
        df_rules = pd.read_excel(excel_path, sheet_name="Rules")
    except Exception as e:
        raise ValueError(f"❌ Rulesシートの読み込みに失敗しました: {e}")

    rules_list = []
    
    # 各行を処理
    for index, row in df_rules.iterrows():
        # Extensionが空の行（Excelの装飾や空行など）はスキップ
        if pd.isna(row.get("Extension")):
            continue
            
        # 値の取得とクリーニング
        ext = str(row["Extension"]).strip()
        proc_id = str(row["Processor ID"]).strip()
        
        # -----------------------------------------------------
        # ★ JSONパラメータ解析 (Bridgeの重要ロジック)
        # Excelの「Parameters」列に書かれたJSON文字列を辞書に変換する
        # -----------------------------------------------------
        raw_params = row.get("Parameters")
        parameters = {}
        
        # NaNチェック (pd.notna) かつ 空文字でないか確認
        if pd.notna(raw_params) and str(raw_params).strip() != "":
            try:
                # 文字列化してからJSONパース
                # Excelが数値を勝手に数値型に変換していても str() で吸収
                param_str = str(raw_params).strip()
                parameters = json.loads(param_str)
                
                # パースは成功したが、辞書じゃない場合（リストなど）への対策
                if not isinstance(parameters, dict):
                    print(f"⚠️ [Warning] 行{index+2}: ParametersはJSONオブジェクト(辞書)である必要があります。無視します。")
                    parameters = {}
                    
            except json.JSONDecodeError as e:
                # 構文ミスはログに出して、デフォルト(空)で続行させる（処理を止めない）
                print(f"⚠️ [Warning] 行{index+2}: ParametersのJSON記述が不正です。")
                print(f"   Value: {raw_params}")
                print(f"   Error: {e}")
                parameters = {}

        # -----------------------------------------------------
        # ルールオブジェクトの生成 (Pydanticバリデーション)
        # -----------------------------------------------------
        try:
            # ProcessorType(Enum)のマッチングもここで行われる
            rule = AttachmentRule(
                extension=ext, 
                processor_id=proc_id, # ここで Enum にない値ならエラーになる
                parameters=parameters
            )
            rules_list.append(rule)
            
        except Exception as e:
            print(f"⚠️ [Warning] 行{index+2}: ルールが無効なためスキップします。({ext} -> {proc_id})")
            print(f"   Reason: {e}")

    # =========================================================
    # 3. 最終Configオブジェクトの生成
    # =========================================================
    config = OutlookConfig(
        job_name=job_name,
        domain=domain,
        search_keywords=keywords,
        destination_path=destination,
        rules=rules_list
    )
    
    print(f"✅ 解析完了: Job='{job_name}' / Domain='{domain}' / Rules={len(rules_list)}件")
    return config