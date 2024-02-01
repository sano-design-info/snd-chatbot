from googleapiclient.discovery import Resource
from googleapiclient.errors import HttpError


def write_data_to_sheet(
    gsheet_service: Resource, spreadsheet_id, sheet_name, data, cell_mapping
):
    # 単一セルへのデータ書き込み
    single_updates = []
    for key, cell in cell_mapping["singlecell"].items():
        value = data.get(key, "")
        single_updates.append({"range": f"{sheet_name}!{cell}", "values": [[value]]})

    # テーブルデータへの書き込み
    table_updates = []
    for table_name, table_info in cell_mapping["tables"].items():
        items = data[table_name]
        start_row = table_info["startRow"]
        for i, item in enumerate(items):
            row = start_row + i
            for column_name, column_letter in table_info["columns"].items():
                value = item.get(column_name, "")
                table_updates.append(
                    {"range": f"{sheet_name}!{column_letter}{row}", "values": [[value]]}
                )

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
