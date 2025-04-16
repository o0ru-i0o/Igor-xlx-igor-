import openpyxl;
import os;
import os.path;
import tkinter;
import tkinter.filedialog;
import pandas;

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

        tkinter.Tk().withdraw()
        tkinter.messagebox.showinfo('メッセージ', str(file_path)+' \nを読み込みます')

            
        for i, sheet_name in enumerate(sheet_names):
            ws = wb[sheet_name];    # シートを取得
            print(f"{i+1}番目のシート名：{sheet_name}");    # シート名を表示
            print("先頭セルの値：", ws.cell(row=1, column=1).value)
            print(f"最大行数：{ws.max_row}");    # A列の行数を表示
          
    else:

        tkinter.Tk().withdraw()
        tkinter.messagebox.showinfo('メッセージ', 'ファイルが選択されませんでした')
        print("ファイルが選択されませんでした");

def edit_excel_file_mass():
    global wb;
    global sheet_names;
    global ws;
    global file_path;
    global mass_number;

    ws = wb[sheet_names[0]];

    mass_number = ws[9];## 9行目を取得
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

    ws.delete_rows(1, 39);    # 1行目から39行目まで削除
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
    csv_file = os.path.dirname(file_path) + "/output/edited_" + os.path.basename(file_path) + "CSVver"

    
    # Excelファイルを読み込む
    df = pandas.read_excel(excel_file)
    
    # CSVファイルに書き込む
    df.to_csv(csv_file, index=False)

    tkinter.Tk().withdraw()
    tkinter.messagebox.showinfo('メッセージ', "読み込んだxlsxをCSVに変換しました！/n(「output」フォルダに保存されています)")
