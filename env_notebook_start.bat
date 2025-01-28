ECHO minicondaの設定および仮想環境の自動起動
ECHO 仮想環境を終了する場合は"deactivate"を実行

set kankyo=notebook
%windir%\System32\cmd.exe /K "C:\Users\fx22228.DC00\AppData\Local\miniconda3\Scripts\activate.bat C:\Users\fx22228.DC00\AppData\Local\miniconda3 & activate %kankyo%"
