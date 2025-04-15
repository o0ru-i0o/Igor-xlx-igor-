import subprocess
import pyperclip

#import excel_editor_00;
#print(excel_editor_00.test);

#subprocess.run(["python", "excel_editor_00.py"])
def launch_igor():
    subprocess.run  (
                    '"C:\\Program Files\\WaveMetrics\\Igor Pro 9 Folder\\IgorBinaries_x64\\Igor64.exe" '
                    '/I "D:\\DQM\\学習\\openpyxl\\インスト\\pythonOpenpyxlのまとめ\\SelfCreate\\Igor提携\\AutoPlot.ipf" ',
                    shell=True
                    )


def to_clip_board():
    # コピーしたい関数名（たとえば Igor の関数名）
    function_call = "xloadwave()"

    # クリップボードにコピー
    pyperclip.copy(function_call)

    print("関数がクリップボードにコピーされたよ！プロシージャーウィンドウに貼り付けてね！")