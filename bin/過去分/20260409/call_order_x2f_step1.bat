@echo off

REM ==========================================
REM Order-X Å® ê≥ãKâªXML Step1
REM ==========================================

call "%~dp0common_env.bat" v1_3

set INPUT_XML=%WORK_DIR%\order_x2f\00030_20260114062503.xml
set OUTPUT_XML=%WORK_DIR%\order_x2f\order_x2f_step1.xml

set SCRIPT=%PROJECT_ROOT%\tools\domains\order\order_x2f_step1.py

echo.
echo ==========================================
echo Order-X2F Step1 start
call "%~dp0common_log.bat" Order-X2F Step1 start
echo ==========================================

python "%SCRIPT%" "%INPUT_XML%" "%OUTPUT_XML%"

if errorlevel 1 (
    echo ERROR : Step1 failed
    call "%~dp0common_log.bat" ERROR : Step1 failed
    exit /b 1
)

echo.
echo Step1 finished : %OUTPUT_XML%
call "%~dp0common_log.bat" Step1 finished : %OUTPUT_XML%

pause
exit /b 0
