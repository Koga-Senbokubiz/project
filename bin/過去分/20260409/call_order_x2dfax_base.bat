@echo off

REM ==========================================
call "%~dp0common_env.bat" v1_3

set param_i=%WORK_DIR%\order_x2f\00030_20260114062503.xml
set mapping=%DICTIONARY_DIR%\BigBoss_DFAX_•ÏŠ·’è‹`.xlsx
set out=%WORK_DIR%\order_x2dfax\dfax_d1_preview.txt

set SCRIPT=%PROJECT_ROOT%\tools\domains\order\order_x2dfax_base.py
rem python "%SCRIPT%" "D:\github\project\work\order_x2f\00030_20260114062503.xml" --mapping "D:\github\project\dictionary\BigBoss_DFAX_•ÏŠ·’è‹`.xlsx" --out "D:\github\project\work\order_x2dfax\dfax_d1_preview.txt" --page-no 1
python "%SCRIPT%" "D:\github\project\work\order_x2f\00030_20260114062503.xml" --out "D:\github\project\work\order_x2dfax\dfax_d1_preview.txt" --bbfax "0661234567" --now "26.04.05 22:55"

if errorlevel 1 (
    echo ERROR : order_x2dfax_base failed
    call "%~dp0common_log.bat" ERROR : order_x2dfax_base failed
pause
    exit /b 1
)

echo.
echo Step3 order_x2dfax_base : %OUTPUT_XML%
call "%~dp0common_log.bat" order_x2dfax_base finished : %OUTPUT_XML%

pause
exit /b 0