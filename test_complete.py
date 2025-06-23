#!/usr/bin/env python3
"""
Excel MCP Server の完全テスト
"""

import json
import os
import sys
from pathlib import Path

# main.pyから関数をインポート
sys.path.append('.')
from main import app

def test_mcp_tools_with_sample():
    """サンプルファイルを使ってMCPツールをテスト"""
    
    sample_file = "sample_data.xlsx"
    
    if not os.path.exists(sample_file):
        print("❌ サンプルファイルが見つかりません。create_sample.py を実行してください。")
        return
    
    print("🧪 Excel MCP Server完全テスト開始")
    print("=" * 60)
      # 関数を直接インポート
    import main
    
    # テスト1: ファイル読み取り
    print("\n1. 📖 ファイル読み取りテスト")
    try:
        result = main.read_excel_file(sample_file, "従業員データ")
        if "error" not in result:
            print(f"✅ 成功: {result['shape']['rows']}行, {result['shape']['columns']}列")
            print(f"   列: {result['columns']}")
        else:
            print(f"❌ エラー: {result['error']}")
    except Exception as e:
        print(f"❌ 例外: {e}")
    
    # テスト2: シート一覧
    print("\n2. 📄 シート一覧テスト")
    try:
        result = main.list_excel_sheets(sample_file)
        if "error" not in result:
            print(f"✅ 成功: {result['total_sheets']}シート")
            print(f"   シート名: {result['sheets']}")
        else:
            print(f"❌ エラー: {result['error']}")
    except Exception as e:
        print(f"❌ 例外: {e}")
    
    # テスト3: データ概要
    print("\n3. 📊 データ概要テスト")
    try:
        result = main.excel_data_summary(sample_file, "売上データ")
        if "error" not in result:
            print(f"✅ 成功: {result['shape']}")
            print(f"   データ型: {result['data_types']}")
            if "numeric_summary" in result:
                print(f"   数値列の統計: あり")
        else:
            print(f"❌ エラー: {result['error']}")
    except Exception as e:
        print(f"❌ 例外: {e}")
    
    # テスト4: セル更新
    print("\n4. ✏️ セル更新テスト")
    try:
        # バックアップ作成
        import shutil
        backup_file = "sample_data_backup.xlsx"
        shutil.copy2(sample_file, backup_file)
        
        result = main.update_excel_cell(sample_file, "従業員データ", 1, "F", "更新テスト")
        if "error" not in result:
            print(f"✅ 成功: セル{result['cell']}を更新")
            print(f"   値: {result['value']}")
        else:
            print(f"❌ エラー: {result['error']}")
        
        # バックアップから復元
        shutil.copy2(backup_file, sample_file)
        os.remove(backup_file)
    except Exception as e:
        print(f"❌ 例外: {e}")
    
    # テスト5: 新しいシート追加
    print("\n5. ➕ シート追加テスト")
    try:
        test_data = [
            {"項目": "テスト1", "値": 100},
            {"項目": "テスト2", "値": 200}
        ]
        
        # バックアップ作成
        backup_file = "sample_data_backup.xlsx"
        shutil.copy2(sample_file, backup_file)
        
        result = main.add_excel_sheet(sample_file, "テストシート", test_data)
        if "error" not in result:
            print(f"✅ 成功: シート'{result['sheet_name']}'を追加")
            print(f"   行数: {result['rows_added']}")
            
            # 確認
            sheets_result = main.list_excel_sheets(sample_file)
            print(f"   現在のシート: {sheets_result['sheets']}")
        else:
            print(f"❌ エラー: {result['error']}")
        
        # バックアップから復元
        shutil.copy2(backup_file, sample_file)
        os.remove(backup_file)
    except Exception as e:
        print(f"❌ 例外: {e}")
    
    # テスト6: 新しいファイル作成
    print("\n6. 📝 新ファイル作成テスト")
    try:
        new_file = "test_output.xlsx"
        test_data = [
            {"名前": "新規太郎", "年齢": 25, "職業": "テスター"},
            {"名前": "作成花子", "年齢": 30, "職業": "エンジニア"}
        ]
        
        result = main.write_excel_file(new_file, test_data, "新規データ")
        if "error" not in result:
            print(f"✅ 成功: ファイル'{result['file_path']}'を作成")
            print(f"   行数: {result['rows_written']}")
            print(f"   列: {result['columns']}")
            
            # ファイルを削除
            if os.path.exists(new_file):
                os.remove(new_file)
                print(f"   🗑️ テストファイルを削除")
        else:
            print(f"❌ エラー: {result['error']}")
    except Exception as e:
        print(f"❌ 例外: {e}")
    
    print("\n" + "=" * 60)
    print("🎉 テスト完了")

def print_usage_examples():
    """使用例を表示"""
    print("\n📚 Excel MCP Server 使用例")
    print("=" * 40)
    
    examples = [
        {
            "title": "📖 ファイル読み取り",
            "tool": "read_excel_file",
            "args": {
                "file_path": "sample_data.xlsx",
                "sheet_name": "従業員データ"
            }
        },
        {
            "title": "📄 シート一覧取得",
            "tool": "list_excel_sheets", 
            "args": {
                "file_path": "sample_data.xlsx"
            }
        },
        {
            "title": "✏️ セル更新",
            "tool": "update_excel_cell",
            "args": {
                "file_path": "sample_data.xlsx",
                "sheet_name": "従業員データ",
                "row": 1,
                "column": "A",
                "value": "更新された値"
            }
        },
        {
            "title": "🎨 セルフォーマット",
            "tool": "format_excel_cells",
            "args": {
                "file_path": "sample_data.xlsx",
                "sheet_name": "従業員データ",
                "cell_range": "A1:E1",
                "font_color": "FFFFFF",
                "bg_color": "4472C4",
                "bold": True,
                "font_size": 12
            }
        },
        {
            "title": "📊 データ概要",
            "tool": "excel_data_summary",
            "args": {
                "file_path": "sample_data.xlsx",
                "sheet_name": "売上データ"
            }
        }
    ]
    
    for example in examples:
        print(f"\n{example['title']}")
        print(f"ツール: {example['tool']}")
        print("引数:")
        print(json.dumps(example['args'], ensure_ascii=False, indent=2))

if __name__ == "__main__":
    test_mcp_tools_with_sample()
    print_usage_examples()
