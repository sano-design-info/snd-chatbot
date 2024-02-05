from googleapiclient.discovery import Resource
from googleapiclient.errors import HttpError


def write_data_to_sheet(gsheet_service: Resource, spreadsheet_id, data, cell_mapping):
    """
    Google Sheetsにデータを書き込む

    データの例
    write_data = {
        "name": "山田太郎",
        "age": 30,
        "email": "taro.yamada@example.internal",
        "department": "DX推進部",
        "assign_project": [
            {"name": "プロジェクトA", "role": "PM"},
            {"name": "プロジェクトB", "role": "PL"},
        ],
    }

    cell_mappingの例
    cell_mapping = {
        "singlecell": {
            "name": "A1",
            "age": "A2",
            "email": "A3",
            "department": "A4",
        },
        "tables": {
            "assign_project": {
                "startRow": 6,
                "endRow": 10,
                "columns": {
                    "name": "A",
                    "role": "B",
                },
            },
        },
    }
    """
    # TODO:2024-02-05 docstringを書く
    single_updates, table_updates = [], []

    # 単一セルへのデータ書き込み
    for key, cell in cell_mapping.get("singlecell", {}).items():
        value = data.get(key, "")
        single_updates.append({"range": cell, "values": [[value]]})

    # テーブルデータへの書き込み
    # TODO:2024-02-05 セルマッピングよりデータのほうが多い場合の処理を追加
    for table_name, table_info in cell_mapping.get("tables", {}).items():
        items = data.get(table_name, [])
        start_row = table_info["startRow"]
        for i, item in enumerate(items):
            row = start_row + i
            for column_name, column_letter in table_info["columns"].items():
                value = item.get(column_name, "")
                cell_address = f"{column_letter}{row}"
                table_updates.append({"range": cell_address, "values": [[value]]})

    # 全ての更新をバッチで適用
    body = {"valueInputOption": "USER_ENTERED", "data": single_updates + table_updates}
    try:
        result = (
            gsheet_service.spreadsheets()
            .values()
            .batchUpdate(spreadsheetId=spreadsheet_id, body=body)
            .execute()
        )
        print(f"{result.get('totalUpdatedCells')} cells updated.")
    except HttpError as err:
        print(err)
