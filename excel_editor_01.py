import openpyxl;
import os;
import os.path;
import tkinter;
import tkinter.filedialog;
import pyperclip
import csv

import re

#グローバル変数の定義
wb = None;
sheet_names = None;
ws = None;
file_path = None;
mass_number = None;


import pandas
import chardet
import os
import tkinter.filedialog

def csv_to_excel_with_pandas():
    #global wb;
    #global sheet_names;
    #global ws;
    global file_path;

    # ファイル選択
    csv_file_path = tkinter.filedialog.askopenfilename(
        title="CSVファイルを選んでね！",
        filetypes=[("CSV files", "*.csv;*.CSV")]
    )

    if not csv_file_path:
        notify_user("キャンセルされたよ〜");
        print("キャンセルされたよ〜")
        return

    # 文字コードを自動検出
    with open(csv_file_path, 'rb') as f:
        raw_data = f.read()
        result = chardet.detect(raw_data)
        detected_encoding = result['encoding']
        #notify_user(f"検出された文字コード：{detected_encoding}")
        print(f"✅ 検出された文字コード：{detected_encoding}")

    # pandasで読み込んで → Excelに出力！
    try:
        df = pandas.read_csv(csv_file_path, encoding=detected_encoding, engine='python', on_bad_lines='skip')

        # 拡張子を安全に置き換え
        filename_root, _ = os.path.splitext(csv_file_path)
        excel_file_path = filename_root + ".xlsx"

        df.to_excel(excel_file_path, index=False)
        #notify_user(f"✅ pandasで変換完了！: {excel_file_path}")
        print(f"✅ pandasで変換完了！: {excel_file_path}")
    except Exception as e:
        #notify_user(f"❌ pandasでの読み込みエラー：{e}")
        print("❌ pandasでの読み込みエラー：", e)

    file_path = excel_file_path;    # グローバル変数にファイルパスを格納


def csv_to_excel_by_csvreader():
    global wb;
    global sheet_names;
    global ws;
    global file_path;

    # CSVファイルの読み込み
    # ファイルダイアログを表示してCSVファイルパスを取得
    csv_file_path = tkinter.filedialog.askopenfilename(
        title="CSVファイルを選択してください",
        filetypes=[("CSV files", "*.csv")]
    )
    excel_file_path = csv_file_path.replace(".csv",".xlsx")
    file_path = excel_file_path;    # グローバル変数にファイルパスを格納


    if csv_file_path:
        csv_file_path = csv_file_path.replace("C:", "");
        csv_file_path = csv_file_path.replace("D:", "");


        wb = openpyxl.Workbook();
        print("選択されたファイル：", csv_file_path);
        ws = wb.active

        

        with open(csv_file_path) as f:
            reader = csv.reader(f, delimiter=',')
            for row in reader:
                ws.append(row)

        # 出力先のExcelファイル名を生成
        excel_file_path = csv_file_path.replace(".csv","_convert.xlsx")


        wb.save(excel_file_path)
    
    else:
        print("CSVファイルが選択されませんでした。")
def csv_to_excel_test():
    global wb;
    global sheet_names;
    global ws;
    global file_path;

    # ファイルダイアログを表示してCSVファイルパスを取得
    csv_file_path = tkinter.filedialog.askopenfilename(
        title="CSVファイルを選択してください",
        filetypes=[("CSV files", "*.csv")]
    )

    # ファイルが選択された場合のみ処理
    if csv_file_path:
        # 出力先のExcelファイル名を生成
        excel_file_path = os.path.splitext(csv_file_path)[0] + "_converted.xlsx"

        # CSVファイルの読み込みと置換
        with open(csv_file_path, 'r', newline='', encoding='utf-8') as file, \
                open('file_out.csv', 'w', newline='', encoding='utf-8') as fileout:
            text = re.sub(r'\s* ', ',', file.read())
            print(text, file=fileout)
            print('置換完了')

        # CSVファイルの読み込み
        data = pandas.read_csv('file_out.csv', encoding='utf-8')

        # Excel形式で出力
        data.to_excel(excel_file_path, encoding='utf-8', index=False)

        print(f'CSV > Excel変換完了: {excel_file_path}')

        file_path = excel_file_path;    # グローバル変数にファイルパスを格納
    else:
        print("CSVファイルが選択されませんでした。")

# Tkinterのウィンドウを非表示にする
root = tkinter.Tk();
root.withdraw();


def read_excel_file():
    global wb;
    global sheet_names;
    global ws;
    global file_path;

    """
    # ファイルダイアログを表示してファイルパスを取得
    file_path = tkinter.filedialog.askopenfilename(
        title="Excelファイルを選択してください",
        filetypes=[("Excel files", "*.xlsx *.xlsm")]
    )
    """

    # ファイルが選択された場合のみ処理
    if file_path:
        # Excelファイルを読み込む
        wb = openpyxl.load_workbook(file_path);
        print("選択されたファイル：", file_path);

        sheet_names = wb.sheetnames;    # シート名のリストを取得
        
        notify_user(f"{str(file_path)}'\n を読み込みます")
        print(f"選択されたファイル：{file_path}");    # 選択されたファイル名を表示
            
        for i, sheet_name in enumerate(sheet_names):
            ws = wb[sheet_name];    # シートを取得
            print(f"{i+1}番目のシート名：{sheet_name}");    # シート名を表示
            print("先頭セルの値：", ws.cell(row=1, column=1).value)
            print(f"最大行数：{ws.max_row}");    # A列の行数を表示
          
    else:

        if tkinter.messagebox.askyesno("エラー", "ファイルが選択されてないよ！今ここで選択する？"):
            # ファイルダイアログを表示してファイルパスを取得
            file_path = tkinter.filedialog.askopenfilename(
                title="Excelファイルを選択してください",
                filetypes=[("Excel files", "*.xlsx *.xlsm")]
            )
            read_excel_file();
        else:
            tkinter.messagebox.showinfo("終了", "ファイルが選択されませんでした");
            print("ファイルが選択されませんでした");


def edit_excel_file_mass():
    global wb;
    global sheet_names;
    global ws;
    global file_path;
    global mass_number;

    mass_number_row = 9;
    header_end_row = 39;

    

    ws = wb[sheet_names[0]];

    col = ws["A"];    # A列を取得
    for cell in col:
        if cell.value == "測定質量数              : ":
            mass_number_row = cell.row;    # 行番号を取得
            print(f"質量数の行番号：{mass_number_row}");    # 行番号を表示
        if cell.value == "測定回数":
            header_end_row = cell.row;
            print(f"ヘッダーの終了行番号：{header_end_row}");    # 行番号を表示 
            break;    # ループを抜ける


    

    mass_number = ws[mass_number_row];# ラベル行目を取得
    print(type(mass_number));    # 取得した行の型を表示
    #print(f"質量数：{mass_number}");


    mass_number_listed = list(mass_number);    # セルの値を取得

    #mass_number_edited = [i for i in mass_number if type(i) == int];    # int型だけ残す
    print(mass_number_listed);    # int型の質量数を表示
    print(type(mass_number_listed));
    print(f"{mass_number_listed[0]=}");
    print(f"{mass_number_listed[0].value=}");    # セルの値を表示

    mass_number_excerpted = [cell.value for cell in mass_number_listed if type(cell.value) == int];
    print(f"{mass_number_excerpted=}");    # int型の質量数を表示

    ws.delete_rows(1, header_end_row);    # 1行目から39行目まで削除
    ws.delete_cols(1,1);
    ws.delete_cols(2,4);

    for cell in ws["A"]:
        cell.value = cell.value[1:12];    # A列の値をスライスして上書き

    ws.insert_rows(1, 1);    # 1行目に1行追加

    ws["A1"].value = "Elapsed Time (s)";
    for i in range(len(mass_number_excerpted)):
        ws.cell(row=1, column=i+2).value = "m=" + str(mass_number_excerpted[i]);    # 1行目に質量数を追加
    ws.delete_cols(len(mass_number_excerpted)+2, ws.max_column);



def save_excel_file():
    global wb;
    global sheet_names;
    global ws;
    global file_path;
    
    print("Excelファイルを保存します。");
    if wb is not None:
        dname = os.path.dirname(file_path);
        fname = os.path.basename(file_path);
        outputFilePath = dname + "/output/edited_" + fname;
        print(f"出力ファイルパス：{outputFilePath}");
        os.makedirs(dname + "/output", exist_ok=True);    # 出力先のディレクトリを作成
        wb.save(outputFilePath) # Excelファイルの保存
        print(f"Excelファイルが保存されました：{outputFilePath}");


    else:
        print("Excelファイルが読み込まれていません。先にread_excel_file()を実行してください。");



def excel_to_csv():
    global wb;
    global sheet_names;
    global ws;
    global file_path;

    excel_file = os.path.dirname(file_path) + "/output/edited_" + os.path.basename(file_path)
    csv_file = os.path.dirname(file_path) + "/output/edited_" + os.path.basename(file_path).replace('.xlsx', '.csv').replace('.xlsm', '.csv')

    
    # Excelファイルを読み込む
    df = pandas.read_excel(excel_file)
    
    # CSVファイルに書き込む
    df.to_csv(csv_file, index=False)

    """
    tkinter.Tk().withdraw()
    tkinter.messagebox.showinfo('メッセージ', "読み込んだxlsxをCSVに変換しました！/n(「output」フォルダに保存されています)")
    """
    print("読み込んだxlsxをCSVに変換しました！ /n (「output」フォルダに保存されています)")

def copy_command_for_Igor():

    csv_file_path_with_collon = os.path.dirname(file_path) + "/output/edited_" + os.path.basename(file_path).replace('.xlsx', '.csv').replace('.xlsm', '.csv')
    csv_file_path_with_collon = csv_file_path_with_collon.replace(":", "")
    csv_file_path_with_collon = csv_file_path_with_collon.replace("/", ":")
    print(f"{csv_file_path_with_collon=}");
    
    # クリップボードにコピー
    #pyperclip.copy('LoadWave/J/D/W/A/E=1/K=0 "D:DQM:学習:openpyxl:インスト:pythonOpenpyxlのまとめ:SelfCreate:Igor提携:output:edited_S1_241017_221354.csv"');
    pyperclip.copy(f'LoadWave/J/D/W/A/E=1/K=0 "{csv_file_path_with_collon}"');


def notify_user(message):
    import tkinter as tk
    from tkinter import messagebox

    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    messagebox.showinfo('メッセージ', message, parent=root)
    root.destroy()