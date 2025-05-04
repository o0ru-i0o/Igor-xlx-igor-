#pragma rtGlobals=3

Function MyAutoPlotter()
    
    LoadWave/J/D/O "D:\\DQM\\学習\\openpyxl\\インスト\\pythonOpenpyxlのまとめ\\SelfCreate\\Igor提携\\Book1.csv"
    //Wave weight
    //Display weight
    //Delimited text load from "D:\\DQM\\学習\\openpyxl\\インスト\\pythonOpenpyxlのまとめ\\SelfCreate\\Igor提携\\Book1.csv"
    //Data length: 11, waves: timeW, weight, radius
        
End

Function MacSayHello()
	
    DoAlert 0, "Hello from Igor!"
End

Function xloadwave()
    XLLoadWave/S="Sheet1"/R=(A1,C12)/W=1/D/T "D:DQM:学習:openpyxl:インスト:pythonOpenpyxlのまとめ:SelfCreate:Igor提携:Book1.xlsx"
End