@echo off

REM ==========================================
REM Order-X2F Step3
REM 顧客データと流通BMS基本形との差異をもとに
REM 変換元レイアウトを作成する。
REM ==========================================

call "%~dp0common_env.bat" v1_3

set param_i=%WORK_DIR%\order_x2f\order_x2f_step1.xml
set param_d=%WORK_DIR%\order_x2f\order_x2f_step2_diff_report.xlsx
set param_t=%TEMPLATE_DIR%\Setting_流通BMS基本形%VERSION_DIR%\Layout\基本形1_3：発注Ver1_3.xml
set param_o=%WORK_DIR%\order_x2f\bbord_xml.xml
set param_p=%WORK_DIR%\order_x2f\bbord_xml_paths.txt
set param_r=%WORK_DIR%\order_x2f\bbord_xml_decisions.txt

set SCRIPT=%PROJECT_ROOT%\tools\domains\order\order_x2f_step3.py

echo.
echo ==========================================
echo Order-X2F Step3 start
call "%~dp0common_log.bat" Order-X2F Step3 start
echo ==========================================

echo param_i = %param_i%
echo param_d = %param_d%
echo param_t = %param_t%
echo param_o = %param_o%
rem echo param_p = %param_p%
rem echo param_r = %param_r%

python "%SCRIPT%" ^
  -i "%param_i%" ^
  -d "%param_d%" ^
  -t "%param_t%" ^
  -o "%param_o%" ^
  -p "%param_p%" ^
  -r "%param_r%"

if errorlevel 1 (
    echo ERROR : Step3 failed
    call "%~dp0common_log.bat" ERROR : Step3 failed
pause
    exit /b 1
)

echo.
echo Step3 finished : %OUTPUT_XML%
call "%~dp0common_log.bat" Step3 finished : %OUTPUT_XML%

pause
exit /b 0