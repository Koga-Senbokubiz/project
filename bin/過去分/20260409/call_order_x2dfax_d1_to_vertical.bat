@echo on
set DFAXDATA_DIR=D:\Van\bigboss\data\out
python "D:\github\project\tools\common\d1_to_vertical.py" "%DFAXDATA_DIR%\DFAXDATA_PHP.txt" "%DFAXDATA_DIR%\DFAXDATA_PHP_VT.txt"
pause
