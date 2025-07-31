import os
from collections import namedtuple
import openpyxl

# Datファイルとシートの紐づけ情報
DatSheetInfo = namedtuple("dat_sheet_info", ["dat_path", "sheet_name"])


def run(
        dat_sheet_pairs: list[DatSheetInfo],
        workbook_file_path: str,
        output_dir: str,
        start_row: int,
        head_pos_col: int,
        length_pos_col: int,
        encoding: str = "utf-8"
):
    dat_sheet_info = dat_sheet_pairs[0]
    dat_data = read_dat(dat_sheet_info.dat_path, encoding)

    # データ全体の長さを計算
    dat_length = len(dat_data) - dat_data.count("\n")

    # ファイルレイアウト情報を取得
    sheet_values = read_xlsx_sheet(workbook_file_path, dat_sheet_info.sheet_name)

    # ファイルレイアウトの末尾の情報を取得
    last_head_pos = None
    last_length_pos = None
    for row in sheet_values[start_row-1:]:
        head_pos = row[head_pos_col-1]
        length_pos = row[length_pos_col-1]

        if head_pos is None and length_pos is None:
            break

        last_head_pos = head_pos
        last_length_pos = length_pos

    # 1行当たりの長さを計算
    line_length = last_head_pos + last_length_pos

    # データを分割
    sliced_data = []
    for i in range(0, dat_length, line_length):
        sliced_data.append(dat_data[i: i + line_length])

    # 分割したデータをcsvとして保存
    with open(os.path.join(output_dir, dat_sheet_info.sheet_name + ".csv"), "w", encoding=encoding) as f:
        f.writelines("\n".join(sliced_data))

def read_dat(file_path: str, encoding: str):
    with open(file_path, encoding=encoding) as f:
        return f.read()


def read_xlsx_sheet(file_path: str, sheet_name: str):
    workbook = openpyxl.load_workbook(file_path, data_only=True)
    sheet = workbook[sheet_name]

    values = []
    for row in sheet.iter_rows():
        row_values = []
        for cell in row:
            row_values.append(cell.value)
        values.append(row_values)
    return values

