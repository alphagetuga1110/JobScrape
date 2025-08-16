# 2_excel_writer.py

import openpyxl

def append_to_excel(filepath, sheet_name, data_list, msg_queue):
    """
    指定されたExcelファイルの指定シートに、データを追記する関数
    """
    try:
        workbook = openpyxl.load_workbook(filepath)
        sheet = workbook[sheet_name]
        start_row = sheet.max_row + 1
        msg_queue.put(f"「{sheet_name}」シートの{start_row}行目から追記を開始します...")
        for row_data in data_list:
            sheet.append(row_data)
        workbook.save(filepath)
        msg_queue.put(f"データの追記が完了しました。ファイルを保存しました: {filepath}")
    except Exception as e:
        msg_queue.put(f"Excelへの書き込み中にエラーが発生しました: {e}")