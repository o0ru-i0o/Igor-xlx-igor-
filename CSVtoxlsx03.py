import csv
import openpyxl

#csv_path = "/DQM/学習/openpyxl/インスト/pythonOpenpyxlのまとめ/SelfCreate/Igor提携/S1_230424_203149.csv"
#excel_path = "/DQM/学習/openpyxl/インスト/pythonOpenpyxlのまとめ/SelfCreate/Igor提携/S1_230424_203149.xlsx"
#csv_path = "/DQM/大学関連/四年次/研究/Qulee/komatsu/231220_⑭-×1.2VN0p9_Fe_thin/S1_231219_124924.CSV"
#excel_path = "/DQM/大学関連/四年次/研究/Qulee/komatsu/231220_⑭-×1.2VN0p9_Fe_thin/S1_230424_203149.xlsx"

csv_path    = "/DQM/学習/openpyxl/インスト/pythonOpenpyxlのまとめ/SelfCreate/Igor提携/文字コードテスト/ANSI_Shift-JIS.csv"
print(f"今回読み込んだcsv_path: {csv_path}")
excel_path  = "/DQM/学習/openpyxl/インスト/pythonOpenpyxlのまとめ/SelfCreate/Igor提携/文字コードテスト/ANSI_Shift-JIS.xlsx"
print(f"今回保存するexcel_path: {excel_path}")

wb = openpyxl.Workbook()
ws = wb.active
 
with open(csv_path) as f:
    reader = csv.reader(f, delimiter=',')
    for row in reader:
        ws.append(row)
 
wb.save(excel_path)

"""
if __name__ == '__main__':
    #csv_to_excel_test()
    #csv_to_excel()
    pass
    #csv_to_excel()

"""