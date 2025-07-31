import os
import configparser
import const
import dat2Csv
from dat2Csv import DatSheetInfo

config = configparser.ConfigParser()

def check_paths(paths: list[str]):
    for path in paths:
        if not os.path.exists(path):
            raise FileNotFoundError(f"Error: Path not exists. Path={path}")
    return


if __name__ == "__main__":
    # 設定ファイル読み込み
    config.read(const.__CONFIG_FILE_PATH__, const.__CONFIG_FILE_ENCODING__)
    config_dat_sheet_pairs = config["datSheetPair"]
    config_paths = config["path"]
    config_positions = config["position"]

    # 入出力先パスの取得
    dat_dir = config_paths["datDir"]
    workbook_file_path = config_paths["workbook"]
    output_dir = config_paths["outputDir"]

    # 存在チェック
    check_paths([dat_dir, workbook_file_path])

    # datファイルとExcelシートの紐づけ
    dat_sheet_pairs: list[DatSheetInfo] = []
    for dat_name, sheet_name in config_dat_sheet_pairs.items():
        dat_file_path = os.path.join(dat_dir, dat_name + ".dat")

        # 存在するdatファイルのみを対象とする
        if not os.path.exists(dat_file_path):
            continue

        dat_sheet_pairs.append(DatSheetInfo(dat_file_path, sheet_name))

    # 出力先フォルダの作成
    if not os.path.exists(output_dir):
        os.mkdir(output_dir)

    dat2Csv.run(
        dat_sheet_pairs,
        workbook_file_path,
        output_dir,
        int(config_positions["startRow"]),
        int(config_positions["headPosCol"]),
        int(config_positions["lengthPosCol"])
    )
