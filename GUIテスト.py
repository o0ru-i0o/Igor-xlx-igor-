import tkinter as tk
from tkinter import ttk, filedialog
import time
import threading
import os

#------------------.pyimport------------------
import CSVtoxlsx04

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
    if not selected_file_path:
        return

    # ボタン無効化
    process_button["state"] = "disabled"
    choose_button["state"] = "disabled"

    # 別スレッドで処理開始
    thread = threading.Thread(target=run_processing)
    thread.start()

def run_processing():

    progress["value"] = 5
    progress_label["text"] = "📂 ファイル読み込み中..."
    progress.update()
    progress_label.update()
        
    #CSV処理のため呼び出し
    CSVtoxlsx04.csv_to_excel_with_pandas_with_argument(selected_file_path)

    for i in range(6,101):
        time.sleep(0.02)  # ← 擬似的な処理時間
        progress["value"] = i
        progress_label["text"] = f"処理中... {i}%"
        progress.update()
        progress_label.update()
    progress_label["text"] = "✅ 完了！"
    finish_button["state"] = "normal"

# GUIセットアップ
root = tk.Tk()
root.title("CSV処理ツール")
root.geometry("450x600")
#root.geometry("400x285")
root.resizable(False, False)

test_button = tk.Button(root, text="テスト", command=lambda:CSVtoxlsx04.csv_to_excel_with_pandas_with_argument(selected_file_path), width=10, height=1)
test_button.pack(pady=(0,2.5))


tk.Label(root, text="📂 MASSの .CSVファイルを .xlsxファイルに変換します", font=("メイリオ", 12)).pack(pady=10);
#grid(row=0, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W+tk.E)
#tk.pack(pady=10)
#tk.grid(row=0, column=0, padx=5, pady=5)

choose_button = tk.Button(root, text="ファイル選択", command=choose_file, width=12, height=1)
choose_button.pack(pady=(0,2.5))
#choose_button.grid(row=1, column=3, padx=5, pady=1)

file_label = tk.Label(root, text="📄 まだファイルが選ばれていません", font=("メイリオ", 10))
file_label.pack(pady=(2.5,5))
#file_label.grid(row=1, column=1, padx=5, pady=5)




process_button = tk.Button(root, text="処理スタート", command=process_file, state="disabled", width=12, height=1)
process_button.pack(pady=(5,2.5))
#process_button.grid(row=2, column=2, padx=5, pady=5)


progress = ttk.Progressbar(root, length=300, mode="determinate")
progress.pack(pady=(2.5,2.5))
#progress.grid(row=3, column=2, padx=5, pady=5)

progress_label = tk.Label(root, text="待機中...", font=("メイリオ", 10))
progress_label.pack(pady=(2.5,5))
#progress_label.grid(row=4, column=2, padx=5, pady=5)

finish_button = tk.Button(root, text="終了", command=root.destroy, state="disabled", width=10, height=1)
finish_button.pack(pady=10)
#finish_button.grid(row=5, column=2, padx=5, pady=5)



root.mainloop()
