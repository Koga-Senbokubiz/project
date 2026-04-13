@echo off

REM ==========================================
REM Order-X2F Step6
REM 顧客データと流通BMS基本形との差異をもとに
REM settingXMLを作成する。
REM ==========================================

call "%~dp0common_env.bat" v1_3

set INPUT_XML=%WORK_DIR%\order_x2f\order_x2f_step1.xml
set MAPPING_XLSX=%WORK_DIR%\order_x2f\order_x2f_step2.xlsx
set OUTPUT_XML=%WORK_DIR%\order_x2f\order_x2f_step6_mapping.xml
set SCRIPT=%PROJECT_ROOT%\tools\domains\order\order_x2f_step6.py
set TEMPLATE_FILE=%TEMPLATE_DIR%\Setting_流通BMS基本形%VERSION_DIR%\基本形1_3：発注Ver1_3（XML→JCA128）.xml

echo.
echo ==========================================
echo Order-X2F Step6 start
call "%~dp0common_log.bat" Order-X2F Step6 start
echo ==========================================

echo INPUT_XML    = %INPUT_XML%
echo MAPPING_XLSX = %MAPPING_XLSX%
echo OUTPUT_XML   = %OUTPUT_XML%
echo TEMPLATE_FILE= %TEMPLATE_FILE%
echo SCRIPT       = %SCRIPT%

python "%SCRIPT%" ^
  --input-xml "%INPUT_XML%" ^
  --mapping-xlsx "%MAPPING_XLSX%" ^
  --template-file "%TEMPLATE_FILE%" ^
  --output-xml "%OUTPUT_XML%"

if errorlevel 1 (
    echo ERROR : Step6	 failed
    call "%~dp0common_log.bat" ERROR : Step6 failed
pause
    exit /b 1
)

echo.
echo Step6 finished : %OUTPUT_XML%
call "%~dp0common_log.bat" Step6 finished : %OUTPUT_XML%

pause
exit /b 0