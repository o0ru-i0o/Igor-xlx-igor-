import pandas as pd
import chardet
import os
import tkinter.filedialog

def read_path():
    # ファイル選択
    csv_file_path = tkinter.filedialog.askopenfilename(
        title="CSVファイルを選んでね！",
        filetypes=[("CSV files", "*.csv;*.CSV")]
    )

    if not csv_file_path:
        print("キャンセルされたよ〜")
        return

    print(f"✅ 選択されたファイル：{csv_file_path}")

if __name__ == "__main__":
    read_path()