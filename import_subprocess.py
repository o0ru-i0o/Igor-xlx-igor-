import subprocess
import pyperclip
import tkinter;
import tkinter.filedialog;


#import excel_editor_00;
#print(excel_editor_00.test);

#subprocess.run(["python", "excel_editor_00.py"])
def launch_igor_with_command():
    subprocess.run  (
                    '"C:\\Program Files\\WaveMetrics\\Igor Pro 9 Folder\\IgorBinaries_x64\\Igor64.exe" '
                    '/I "D:\\DQM\\学習\\openpyxl\\インスト\\pythonOpenpyxlのまとめ\\SelfCreate\\Igor提携\\AutoPlot.ipf" ',
                    shell=True
                    )

def launch_igor():
    tkinter.Tk().withdraw()
    tkinter.messagebox.showinfo('メッセージ', "Igorが起動したら，Ctrl+Vで貼り付けてEnter！")

    subprocess.Popen('"C:\\Program Files\\WaveMetrics\\Igor Pro 9 Folder\\IgorBinaries_x64\\Igor64.exe" ')


def to_clip_board_ipfFunction():
    # コピーしたい関数名（たとえば Igor の関数名）
    function_call = "xloadwave()"

    # クリップボードにコピー
    pyperclip.copy(function_call)

    print("関数がクリップボードにコピーされたよ！プロシージャーウィンドウに貼り付けてね！")

def to_clip_boardan_Procedure():
    #tkinter.Tk().withdraw()
    #tkinter.messagebox.showinfo('メッセージ', "Ctrl+Vで貼り付けてEnter！")


    # クリップボードにコピー
    pyperclip.copy('LoadWave/J/D/W/A/E=1/K=0 "D:DQM:学習:openpyxl:インスト:pythonOpenpyxlのまとめ:SelfCreate:Igor提携:output:edited_S1_241017_221354.csv"');

    
