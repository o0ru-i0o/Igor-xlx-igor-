import tkinter.messagebox
import pandas as pd
import chardet
import os
import tkinter.filedialog
import csv
import collections
from collections import Counter
from tkinter import messagebox
import traceback

file_path = None  # グローバル変数にファイルパスを格納

#対象ファイルのパス，検出した文字コードをGUIに返す，進捗バーをGUIに返す
def csv_to_excel_with_pandas_with_argument(path, notify_encoding=None, progress_callback=None):
    global file_path;
    #tkinter.messagebox.showinfo('メッセージ', "csv_to_excel_with_pandas_with_argument" + str(path));

    #print(path);
    #print("csv_to_excel_with_pandas_with_argument");

    csv_file_path = path;

    if not csv_file_path:
        print("キャンセルされたよ〜")
        return
    


    # 文字コードを自動検出
    print(f"{csv_file_path}の文字コードを検出中...");
    try:
        with open(csv_file_path, 'rb') as f:
            raw_data = f.read()
            result = chardet.detect(raw_data)
            detected_encoding = result['encoding']
            print(f"✅ 検出された文字コード：{detected_encoding}")
            #GUIテスト.encord_label.config(text="✅ 検出された文字コード：" + detected_encoding)
            #notify_user(f"検出された文字コード：{detected_encoding}")
            if notify_encoding:
                notify_encoding(detected_encoding)  # ← GUI側に通知！
            if progress_callback:
                progress_callback(10)
    except Exception as e:
        print("❌ 文字コード検出エラー：", e)
        tb = traceback.extract_tb(e.__traceback__)
        last_trace = tb[-1]  # 最後のトレース（エラーが起きた場所）
        line_number = last_trace.lineno
        filename = last_trace.filename
        error_message = f"💥 不具合が起きたようです...以下の内容を小松路易に伝えてください:\n{e=}\n📄 ファイル: {filename}\n📍 行番号: {line_number}"
        messagebox.showerror("エラー発生！", error_message)
        return
    

    #前処理(” を削除)
    print(f"{csv_file_path}の前処理中...");
    try:
        with open(csv_file_path, "r", encoding=detected_encoding) as f:
            lines = [line.replace('"', '') for line in f]

            if progress_callback:
                progress_callback(20)
    except Exception as e:
        print("❌ 文字コード検出エラー：", e)
        tb = traceback.extract_tb(e.__traceback__)
        last_trace = tb[-1]  # 最後のトレース（エラーが起きた場所）
        line_number = last_trace.lineno
        filename = last_trace.filename
        error_message = f"💥 不具合が起きたようです...以下の内容を小松路易に伝えてください:\n{e=}\n📄 ファイル: {filename}\n📍 行番号: {line_number}"
        messagebox.showerror("エラー発生！", error_message)
        return
    
    filename_root, _ = os.path.splitext(csv_file_path)
    #csv_file_path_cleaned = filename_root + "_cleaned.xlsx"
    csv_file_path_cleaned = filename_root + "_cleaned.csv"
    

    try:
        with open(csv_file_path_cleaned, "w", encoding=detected_encoding) as f:
            f.writelines(lines)
    except Exception as e:
        tb = traceback.extract_tb(e.__traceback__)
        last_trace = tb[-1]  # 最後のトレース（エラーが起きた場所）
        line_number = last_trace.lineno
        filename = last_trace.filename
        error_message = f"💥 不具合が起きたようです...以下の内容を小松路易に伝えてください:\n{e=}\n📄 ファイル: {filename}\n📍 行番号: {line_number}"
        messagebox.showerror("エラー発生！", error_message)
        return

        #print(lines)

        if progress_callback:
            progress_callback(30)

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

        if progress_callback:
            progress_callback(50)

        # セルごとに safe_convert を適用！
        df = df.applymap(safe_convert)
        #print(f"{df[1:40]}");

        if progress_callback:
            progress_callback(75)

        # 拡張子を安全に置き換え
        filename_root, _ = os.path.splitext(csv_file_path)
        excel_file_path = filename_root + ".xlsx"

        dname = os.path.dirname(excel_file_path);
        fname = os.path.basename(excel_file_path);
        excel_file_path = dname + "/output/cleaned_" + fname;


        if progress_callback:
            progress_callback(100)

        file_path = excel_file_path;  # グローバル変数にファイルパスを格納
        df.to_excel(excel_file_path, index=False)
        print(f"✅ pandasで変換完了！: {excel_file_path}")

        # _cleaned.csvファイルを削除
        if os.path.exists(csv_file_path_cleaned):   
            os.remove(csv_file_path_cleaned)
            print(f"{csv_file_path_cleaned} を削除しました✅")
        else:
            print("削除対象のファイルが見つかりませんでした❌")

    except Exception as e:
        print("❌ pandasでの読み込みエラー：", e)
        tb = traceback.extract_tb(e.__traceback__)
        last_trace = tb[-1]  # 最後のトレース（エラーが起きた場所）
        line_number = last_trace.lineno
        filename = last_trace.filename
        error_message = f"💥 不具合が起きたようです...以下の内容を小松路易に伝えてください:\n{e=}\n📄 ファイル: {filename}\n📍 行番号: {line_number}"
        messagebox.showerror("エラー発生！", error_message)
        return


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

#
def return_xlsx_file_path():
    global file_path;
    #tkinter.messagebox.showinfo('メッセージ', "return_xlsx_file_path");
    #print(file_path);
    return file_path;
#def csv_to_excel_by_csvreader():

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