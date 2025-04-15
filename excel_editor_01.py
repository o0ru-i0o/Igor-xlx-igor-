import openpyxl;
import os;
import os.path;
import tkinter;
import tkinter.filedialog;

# Tkinterのウィンドウを非表示にする
root = tkinter.Tk();
root.withdraw();

#グローバル変数の定義
wb = None;
sheet_names = None;
ws = None;
file_path = None;
mass_number = None;

def read_excel_file():
    global wb;
    global sheet_names;
    global ws;
    global file_path;

    # ファイルダイアログを表示してファイルパスを取得
    file_path = tkinter.filedialog.askopenfilename(
        title="Excelファイルを選択してください",
        filetypes=[("Excel files", "*.xlsx *.xlsm")]
    )

    # ファイルが選択された場合のみ処理
    if file_path:
        # Excelファイルを読み込む
        wb = openpyxl.load_workbook(file_path);
        print("選択されたファイル：", file_path);

        sheet_names = wb.sheetnames;    # シート名のリストを取得
            
        for i, sheet_name in enumerate(sheet_names):
            ws = wb[sheet_name];    # シートを取得
            print(f"{i+1}番目のシート名：{sheet_name}");    # シート名を表示
            print("先頭セルの値：", ws.cell(row=1, column=1).value)
            print(f"最大行数：{ws.max_row}");    # A列の行数を表示
          
    else:
        print("ファイルが選択されませんでした");

def edit_excel_file_mass():
    global wb;
    global sheet_names;
    global ws;
    global file_path;
    global mass_number;

    ws = wb[sheet_names[0]];

    mass_number = ws[9];## 9行目を取得
    print(f"質量数：{mass_number}");


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