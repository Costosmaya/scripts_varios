@echo off
forfiles /P "C:\Users\User\Documents\Analisis  Desarrollo Costos\Scripts\Python\imgs_qa" /M * /C "cmd /c if @isdir==FALSE del @file"
echo activate conda main
"py" "c:/Users/User/Documents/Analisis  Desarrollo Costos/Scripts/Python/ocr_QA_envios.py"
pause