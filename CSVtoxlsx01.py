import datetime
import pathlib
from tkinter import filedialog

import openpyxl
import pandas as pd


def main():
    """csvファイルの一覧を取得するメイン関数"""

    # csvファイルの選択画面を開く
    csv_files = filedialog.askopenfilenames(
        title="csvファイルを選択する", filetypes=[("csvファイル", ".csv")]
    )

    # 選択されたcsvファイルの一覧を取得し、csv用とExcel用のファイルパスを生成する
    for file in csv_files:
        csv_path = pathlib.Path(file)
        excel_path = csv_path.with_suffix(".xlsx")

        # 既に同名のExcelファイルが存在する場合、上書きを防ぐためファイル名にタイムスタンプをつける
        if pathlib.Path.exists(excel_path) is True:

            # タイムスタンプを生成する
            now = datetime.datetime.now()
            time_num = now.strftime("%Y%m%d%H%M%S")

            # ファイルパスにタイムスタンプをつける
            excel_path = (
                str(excel_path.parent)
                + "\\"
                + str(excel_path.stem)
                + "_"
                + time_num
                + ".xlsx"
            )

        conversion_to_excel(csv_path, excel_path)


def conversion_to_excel(csv_path, excel_path):
    """csvファイルをExcelに変換する関数

    Args:
        csv_path (pathlib.WindowsPath): csv用のファイルパス
        excel_path (pathlib.WindowsPath): Excel用のファイルパス

    Todo:
        現状では文字コードがutf-8とSJISの場合しか対応していないので、
        より柔軟に多数の文字コードに対応できるよう改善したい
    """

    # pandasでcsvを読み込む
    # ヘッダーがないcsvも扱うので、header=Noneを指定する
    try:
        df = pd.read_csv(csv_path, header=None, encoding="utf-8")  # 最初はutf-8で読み込み、
    except UnicodeDecodeError:
        df = pd.read_csv(
            csv_path, header=None, encoding="cp932"
        )  # 文字コードでエラーが発生したら、SJISで読み込む

    # データフレームをExcelとして保存する
    # Excelには既に行番号があるので、index=Falseを指定する
    df.to_excel(excel_path, index=False)

    # csvをpandasで読み込むとき、header=Noneを指定すると1行目に列番号が生成される
    # Excelで扱う上では不要な行なので、1行目を削除する
    wb = openpyxl.load_workbook(excel_path)
    ws = wb["Sheet1"]
    ws.delete_rows(1)
    wb.save(excel_path)


if __name__ == "__main__":
    main()