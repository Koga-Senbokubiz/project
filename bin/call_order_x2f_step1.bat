@echo off

REM ==========================================
REM Order-X Å® ê≥ãKâªXML Step1
REM ==========================================

set PROJECT_ROOT=D:\github\project

set INPUT_XML=%PROJECT_ROOT%\data\input_order.xml
set OUTPUT_XML=%PROJECT_ROOT%\work\order_step1.xml

set SCRIPT=%PROJECT_ROOT%\tools\domains\order\order_x2f_step1.py

echo.
echo ==========================================
echo Order-X2F Step1 start
echo ==========================================

python "%SCRIPT%" "%INPUT_XML%" "%OUTPUT_XML%"

if errorlevel 1 (
    echo ERROR : Step1 failed
    pause
    exit /b 1
)

echo.
echo Step1 finished
echo output :
echo %OUTPUT_XML%

pause