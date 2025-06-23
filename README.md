# Excel File Editor MCP Server

このMCPサーバーは、Excelファイルの読み書きや編集機能を提供します。

## 機能

### 基本操作
- `read_excel_file`: Excelファイルを読み取り
- `write_excel_file`: データをExcelファイルに書き込み
- `list_excel_sheets`: ファイル内のシート一覧を取得

### セル操作
- `update_excel_cell`: 特定のセルを更新
- `format_excel_cells`: セル範囲のフォーマット設定

### シート操作
- `add_excel_sheet`: 新しいシートを追加
- `delete_excel_sheet`: シートを削除

### データ分析
- `excel_data_summary`: データの概要と統計情報を取得

## インストール

```bash
pip install -r requirements.txt
```

## 使用方法

```bash
python main.py
```

## 例

### ファイルの読み取り
```json
{
  "tool": "read_excel_file",
  "arguments": {
    "file_path": "data.xlsx",
    "sheet_name": "Sheet1"
  }
}
```

### セルの更新
```json
{
  "tool": "update_excel_cell",
  "arguments": {
    "file_path": "data.xlsx",
    "sheet_name": "Sheet1",
    "row": 1,
    "column": "A",
    "value": "新しい値"
  }
}
```

### フォーマット設定
```json
{
  "tool": "format_excel_cells",
  "arguments": {
    "file_path": "data.xlsx",
    "sheet_name": "Sheet1",
    "cell_range": "A1:C3",
    "font_color": "FF0000",
    "bg_color": "FFFF00",
    "bold": true,
    "font_size": 14
  }
}
```

## 依存関係

- fastmcp: MCP サーバーフレームワーク
- pandas: データ処理
- openpyxl: Excel ファイル操作
- xlrd: 旧形式Excel読み取り
- xlsxwriter: Excel書き込み
