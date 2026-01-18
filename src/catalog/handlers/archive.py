import os
import time
from src.catalog import register_processor
from src.schema.definitions import ProcessorType

# 依存ライブラリ: zipfileは標準ライブラリ
import zipfile

@register_processor(ProcessorType.UNZIP)
def logic_unzip(attachment, output_dir, params: dict = {}):
    """
    Zip解凍ロジック (v2.0 Parameter対応)
    
    Args:
        attachment: AttachmentWrapperオブジェクト
        output_dir: 保存先パス
        params (dict): 動作制御パラメータ
            - mode (str): 'auto' | 'fixed' | 'manual'
            - password (str): mode='fixed' の場合のパスワード
    """
    
    # 1. まずZIPファイル自体を保存する
    # (解凍に失敗しても原本は残すため)
    save_path = os.path.join(output_dir, attachment.filename)
    attachment.save_as(save_path)
    
    # 解凍先のフォルダ名を作成 (例: ./data/extracted_filename/)
    extract_folder_name = f"extracted_{os.path.splitext(attachment.filename)[0]}"
    extract_path = os.path.join(output_dir, extract_folder_name)

    # -------------------------------------------------------
    # ★ ビジネスルールの適用 (Policy Injection)
    # パラメータに応じて「パスワードをどう調達するか」を決定する
    # -------------------------------------------------------
    mode = params.get("mode", "auto")
    password_str = None

    if mode == "fixed":
        # ケースA: 社内システム連携など、パスワードが決まっている場合
        password_str = params.get("password")
        if not password_str:
            print("    ⚠️ [Warning] mode=fixed ですが password が指定されていません。")
        else:
            print(f"    🔓 [Policy] 固定パスワードを使用します: {'*' * len(password_str)}")

    elif mode == "manual":
        # ケースB: 人間判断が必要な場合（不定期な取引先など）
        print(f"\n    🛑 【Action Required】 {attachment.filename} は保護されています。")
        print("    (ヒント: メール本文を確認してください)")
        # 実行時にコンソールで入力を待つ
        password_str = input("    🔑 パスワードを入力してください >> ").strip()
        print("    🔓 [Policy] ユーザー入力を受け付けました。")
    
    else:
        # ケースC: パスワードなし (auto)
        print("    📦 [Policy] パスワードなしとして処理します。")

    # -------------------------------------------------------
    # 3. 解凍実行 (Mechanism)
    # -------------------------------------------------------
    try:
        # パスワードはbytes型にする必要がある
        pwd_bytes = password_str.encode('utf-8') if password_str else None
        
        if not os.path.exists(extract_path):
            os.makedirs(extract_path)

        # 実際に解凍を試みる (疑似コードではなく標準ライブラリを使用)
        # ※注意: 暗号化ZIPの解凍には本来 'pyzipper' 等が必要な場合があるが、
        # ここでは標準の zipfile で実装
        with zipfile.ZipFile(save_path, 'r') as zip_ref:
            # パスワード付きかどうかをチェックするロジックを入れても良いが
            # ここではシンプルにトライ＆エラー
            zip_ref.extractall(extract_path, pwd=pwd_bytes)
            
        print(f"    ✅ 解凍成功: {extract_path}")
        
        # 中身のファイル一覧を表示（デバッグ用）
        files = os.listdir(extract_path)
        print(f"       -> 展開されたファイル: {files}")

    except RuntimeError as e:
        if 'Bad password' in str(e):
            print(f"    ❌ 解凍失敗: パスワードが間違っています。")
        else:
            print(f"    ❌ 解凍失敗: パスワードが必要か、ファイルが壊れています。({e})")
    except Exception as e:
        print(f"    💥 解凍プロセスで予期せぬエラー: {e}")