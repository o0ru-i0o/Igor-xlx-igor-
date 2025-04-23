import tkinter as tk
from tkinter import ttk, filedialog
import time
import threading
import os

#------------------.pyimport------------------
import CSVtoxlsx04
import excel_editor_01

#---------------------------------------------

selected_file_path = None  # ← グローバルに保持


def choose_file():
    global selected_file_path
    path = filedialog.askopenfilename(filetypes=[("CSVファイル", "*.csv;*.CSV")])
    if path:
        selected_file_path = path
        file_label.config(text=f"📄 選択中：{os.path.basename(path)}")
        process_button["state"] = "normal"
        progress["value"] = 0
        progress_label["text"] = "待機中..."

def process_file():
    global thread1; #スレッドが終了したかどうか別関数から確認するために必要
    if not selected_file_path:
        return
    progress["value"] = 5
    progress_label["text"] = "読み込んでいます..."
    # ボタン無効化
    process_button["state"] = "disabled"
    choose_button["state"] = "disabled"

    # 別スレッドで処理開始
        #CSV処理のため呼び出し,進捗バーの更新のため呼び出し

    thread1 = threading.Thread   (target=lambda: (
                                CSVtoxlsx04.csv_to_excel_with_pandas_with_argument(selected_file_path, notify_encoding=show_encoding_on_gui,progress_callback=update_progress),
                                run_processing()
                                ))
    thread1.start()
    #thread.join()  
    #join()は完全にブロッキングなの！
    #つまり「スレッドが終わるまでPythonの処理が完全停止しちゃう」！
    #TkinterみたいなGUIではメインループ（mainloop()）が止まるとGUIも固まるから、
    #→ 「フリーズした！」って見えるってわけ！
    check_thread_then_start_thread2()  # スレッドが終了したら次の処理を開始
    
    #excel_editor_01.read_excel_file(path=excel_editor_01.return_xlsx_file_path())  # ← Excelファイルを開く

def check_thread_then_start_thread2():
    global thread1; #スレッドが終了したかどうか別関数から確認するために必要

    if thread1.is_alive():   #thread=宣言してから使ってね！
        root.after(100, check_thread_then_start_thread2) # 100ms後に再チェック
    else:
        print("スレッドが終了しました！(もしくは、スレッドが無いかも...)")
        start_thread2()  # スレッドが終了したら次の処理を開始

def start_thread2():
    global thread2; #スレッドが終了したかどうか別関数から確認するために必要
    xlsx_path = CSVtoxlsx04.return_xlsx_file_path()
    thread2 = threading.Thread  (target=lambda:(
                                excel_editor_01.xlsx_to_csv_to_igor_integrated(path=xlsx_path, progress_callback=update_progress), # ← Excelファイルを開く
                                run_processing()
                                ))  
    thread2.start()

"""
def check_thread_then_start_thread3():
    global thread2; #スレッドが終了したかどうか別関数から確認するために必要

    if thread2.is_alive():   #thread=宣言してから使ってね！
        root.after(100, check_thread_then_start_thread3) # 100ms後に再チェック
    else:
        print("スレッドが終了しました！(もしくは、スレッドが無いかも...)")
        start_thread3()  # スレッドが終了したら次の処理を開始
"""



def show_encoding_on_gui(enc):
    encode_label.config(text=f"✅ 検出された文字コード：{enc}")
    encode_label.pack(after=progress)  # ← 改めて表示！


def update_progress(percent):
    progress["value"] = percent
    if percent == 100:
        progress_label["text"] = "✅ 完了！"
        finish_button["state"] = "normal"
    else:
        progress_label["text"] = f"処理中... {percent}%"
    progress.update()
    progress_label.update()
    root.update_idletasks()

def run_processing():
    pass
    #progress["value"] = 5
    #progress_label["text"] = "📂 ファイル読み込み中..."
    #progress.update()
    #progress_label.update()
        
"""
    for i in range(6,101):
        time.sleep(0.02)  # ← 擬似的な処理時間
        progress["value"] = i
        progress_label["text"] = f"処理中... {i}%"
        progress.update()
        progress_label.update()
    progress_label["text"] = "✅ 完了！"
    finish_button["state"] = "normal"
"""

# GUIセットアップ
root = tk.Tk()
root.title("CSV処理ツール")
root.geometry("450x600")
#root.geometry("400x285")
root.resizable(False, False)

#test_button = tk.Button(root, text="テスト", command=lambda:CSVtoxlsx04.csv_to_excel_with_pandas_with_argument(selected_file_path), width=10, height=1)
#test_button.pack(pady=(0,2.5))


tk.Label(root, text="📂 .CSVファイルを選択してください\n(MASS(Qulee)から出力されたもの)", font=("メイリオ", 12)).pack(pady=(5,5));
#grid(row=0, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W+tk.E)
#tk.pack(pady=10)
#tk.grid(row=0, column=0, padx=5, pady=5)

separator01 = ttk.Separator(root, orient='horizontal')
separator01.pack(fill='x', padx=20, pady=0)

choose_button = tk.Button(root, text="ファイル選択", command=choose_file, width=12, height=1)
choose_button.pack(pady=(5,2.5))
#choose_button.grid(row=1, column=3, padx=5, pady=1)

file_label = tk.Label(root, text="📄 まだファイルが選ばれていません", font=("メイリオ", 8))
file_label.pack(pady=(2.5,5))
#file_label.grid(row=1, column=1, padx=5, pady=5)

arrow_label = tk.Label(root, text="▼", font=("Arial", 16))  # サイズ調整もできるよ！
arrow_label.pack()


process_button = tk.Button(root, text="処理スタート", command=process_file, state="disabled", width=12, height=1)
process_button.pack(pady=(5,2.5))
#process_button.grid(row=2, column=2, padx=5, pady=5)


progress = ttk.Progressbar(root, length=300, mode="determinate")
progress.pack(pady=(2.5,2.5))
#progress.grid(row=3, column=2, padx=5, pady=5)

encode_label = tk.Label(root, text="✅ 検出された文字コード：", font=("メイリオ", 8))
encode_label.pack(after=progress,pady=(1.5,1))
encode_label.pack_forget()  # ← 最初は隠す！



progress_label = tk.Label(root, text="待機中...", font=("メイリオ", 8))
progress_label.pack(pady=(1,5))
#progress_label.grid(row=4, column=2, padx=5, pady=5)

arrow2 = tk.Label(root, text="▼", font=("Arial", 16))
arrow2.pack()

save_label01 = tk.Label(root, text="📂 変換後のファイル名：\n   変換後の保存場所 ： ", font=("メイリオ", 8), fg="green")
save_label01.pack(after=progress,pady=(5,5))
save_label01.pack_forget()  # ← 最初は隠す！

finish_button = tk.Button(root, text="終了", command=root.destroy, state="disabled", width=10, height=1)
finish_button.pack(pady=10)
#finish_button.grid(row=5, column=2, padx=5, pady=5)


# 🔥 ここ重要！他ファイルからもアクセス可能にしたいので、
# mainloopを直で起動する構成にしておく
#if __name__ == "__main__":
root.mainloop()
