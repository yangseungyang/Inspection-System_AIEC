@echo off
rem <ドライブCの検査データをネットワークサーバにコピー>
xcopy "C:\標準化検査装置" "\\10.36.14.211\aiec\AIEC_B\標準化検査装置\" /D /S /E /C /F /H /R /K /Y /Z
rem pause
