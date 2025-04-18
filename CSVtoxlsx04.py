import pandas as pd
import chardet
import os
import tkinter.filedialog

def csv_to_excel_with_pandas():
    # ファイル選択
    csv_file_path = tkinter.filedialog.askopenfilename(
        title="CSVファイルを選んでね！",
        filetypes=[("CSV files", "*.csv;*.CSV")]
    )

    if not csv_file_path:
        print("キャンセルされたよ〜")
        return

    # 文字コードを自動検出
    with open(csv_file_path, 'rb') as f:
        raw_data = f.read()
        result = chardet.detect(raw_data)
        detected_encoding = result['encoding']
        print(f"✅ 検出された文字コード：{detected_encoding}")

    # pandasで読み込んで → Excelに出力！
    try:
        df = pd.read_csv(csv_file_path, encoding=detected_encoding, engine='python', on_bad_lines='skip')

        # 拡張子を安全に置き換え
        filename_root, _ = os.path.splitext(csv_file_path)
        excel_file_path = filename_root + ".xlsx"

        df.to_excel(excel_file_path, index=False)
        print(f"✅ pandasで変換完了！: {excel_file_path}")
    except Exception as e:
        print("❌ pandasでの読み込みエラー：", e)


if __name__ == '__main__':
    csv_to_excel_with_pandas()
