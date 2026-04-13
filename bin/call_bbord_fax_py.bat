@echo off

call common_env.bat v1_3
set MAIN_PY=%TOOLS_DIR%\domains\order\bbord_fax.py
set JOBID=%DATE:~0,4%%DATE:~5,2%%DATE:~8,2%

%PYTHON_CMD% %MAIN_PY%  ^
    "J%JOBID%" ^
    "D:\github\project\work\order_x2dfax\in\J9682128\FAX" ^
    "D:\github\project\work\order_x2dfax\out\DFAXDATA_py.txt" ^
    --ini "D:\github\project\ini\bbord_fax.ini"

if %ERRORLEVEL% neq 0 (
    echo ÉGÉâÅ[î≠ê∂
pause
    exit /b %ERRORLEVEL%
)

pause
exit /b 0
