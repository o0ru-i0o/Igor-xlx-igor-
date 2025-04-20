import tkinter.messagebox
import pandas as pd
import chardet
import os
import tkinter.filedialog
import csv
import collections
from collections import Counter

def csv_to_excel_with_pandas_with_argument(path):
    #tkinter.messagebox.showinfo('メッセージ', "csv_to_excel_with_pandas_with_argument" + str(path));

    #print(path);
    #print("csv_to_excel_with_pandas_with_argument");

    csv_file_path = path;

    if not csv_file_path:
        print("キャンセルされたよ〜")
        return
    


    # 文字コードを自動検出
    print(f"{csv_file_path}の文字コードを検出中...");
    with open(csv_file_path, 'rb') as f:
        raw_data = f.read()
        result = chardet.detect(raw_data)
        detected_encoding = result['encoding']
        print(f"✅ 検出された文字コード：{detected_encoding}")

    

    #前処理(” を削除)
    print(f"{csv_file_path}の前処理中...");
    
    with open(csv_file_path, "r", encoding=detected_encoding) as f:
        lines = [line.replace('"', '') for line in f]

    filename_root, _ = os.path.splitext(csv_file_path)
    #csv_file_path_cleaned = filename_root + "_cleaned.xlsx"
    csv_file_path_cleaned = filename_root + "_cleaned.csv"
    


    with open(csv_file_path_cleaned, "w", encoding=detected_encoding) as f:
        f.writelines(lines)
        #print(lines)

    #df = pd.read_csv("cleaned_file.csv", encoding="utf-8")    

    #区切り文字を自動判断
    #detected_delimiter = detecting_delimiter(csv_file_path_cleaned, detected_encoding,count_csv_lines_fast(csv_file_path_cleaned, detected_encoding));
    #print(f"✅ 検出された区切り文字：{detected_delimiter}")

    #最大列数を自動検出
    max_cols = detect_max_columns(csv_file_path_cleaned, detected_encoding)
    print(f"📏 最大列数は {max_cols} 列です！")    


    # pandasで読み込んで → Excelに出力！
    print(f"{csv_file_path_cleaned}の読み込み中...");
    try:
        colnames = [f"col{i+1}" for i in range(max_cols)]
        df = pd.read_csv(
            csv_file_path_cleaned, 
            encoding=detected_encoding, 
            #delimiter=detected_delimiter,  # 自動検出した区切り文字を使用
            #delimiter=',',
            engine='python', 
            names=colnames,
            #header=35,  # ← ヘッダー行を指定
            #quotechar='"',        # ← これを追加
            quotechar="|",  # ← これを追加
            #on_bad_lines='skip',   # ← 必要に応じて調整
            #quoting=csv.QUOTE_NONNUMERIC, 
            dtype=str   # 一旦「全部文字列」として読む（header=Noneなど調整してね！）
        )
        # セルごとに safe_convert を適用！
        df = df.applymap(safe_convert)
        #print(f"{df[1:40]}");


        # 拡張子を安全に置き換え
        filename_root, _ = os.path.splitext(csv_file_path)
        excel_file_path = filename_root + ".xlsx"

        df.to_excel(excel_file_path, index=False)
        print(f"✅ pandasで変換完了！: {excel_file_path}")
    except Exception as e:
        print("❌ pandasでの読み込みエラー：", e)

#区切り文字判断
def detecting_delimiter(file_path, encoding, sample_lines):
    with open(file_path, 'r', encoding=encoding, errors='ignore') as f:
        sample = [next(f) for _ in range(sample_lines)]

    delimiters = [',', '\t', ';', '|']
    delimiter_counts = {}

    for delim in delimiters:
        counts = [len(line.split(delim)) for line in sample]
        # 一番よく一致する数の出現回数を使ってスコア化
        most_common_count = Counter(counts).most_common(1)[0][1]
        delimiter_counts[delim] = most_common_count

    best_delim = max(delimiter_counts, key=delimiter_counts.get)
    print(f"🔍 判定結果：{delimiter_counts} → 最終判定：{repr(best_delim)}")
    return best_delim

#最大列数を自動検出
def detect_max_columns(file_path, encoding):
    delimiter=','
    max_columns = 0
    with open(file_path, 'r', encoding=encoding, newline='') as f:
        reader = csv.reader(f, delimiter=delimiter)
        for row in reader:
            max_columns = max(max_columns, len(row))
    return max_columns

#最大行数を自動検出
def count_csv_lines_fast(file_path, encoding):
    with open(file_path, encoding=encoding, errors='ignore') as f:
        return sum(1 for _ in f)

# セルごとに変換を適用する関数
def safe_convert(val):
    try:
        return float(val)
    except (ValueError, TypeError):
        return val  # 数値変換できなければそのまま返す



"""
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
"""

"""
if __name__ == '__main__':
    csv_to_excel_with_pandas()
"""