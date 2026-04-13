@echo off

REM ==========================================
REM Order-X2F Step5
REM 顧客データと流通BMS基本形との差異をもとに
REM 変換元レイアウトを作成する。
REM ==========================================

call "%~dp0common_env.bat" v1_3

set SCRIPT=%PROJECT_ROOT%\tools\domains\order\order_x2f_step5.py
set INPUT_XML=%WORK_DIR%\order_x2f\order_x2f_step1.xml
set DICTIONARY=%DICTIONARY_DIR%\order_bigboss_dictionary.xlsx
set TEMPLATE_FILE=%TEMPLATE_DIR%\Setting_流通BMS基本形%VERSION_DIR%\Logic\基本形1_3：発注Ver1_3（XML→JCA128）_基本形1_3：発注Ver1_3_基本形1_3：発注JCA128.xml
set OUTPUT_XML=%WORK_DIR%\order_x2f\order_x2f_step5_logic.xml

echo.
echo ==========================================
echo Order-X2F Step5 start
call "%~dp0common_log.bat" Order-X2F Step5 start
echo ==========================================

echo INPUT_XML    = %INPUT_XML%
echo DICTIONARY   = %DICTIONARY%
echo TEMPLATE_FILE= %TEMPLATE_FILE%
echo OUTPUT_XML   = %OUTPUT_XML%

python "%SCRIPT%" ^
  -i "%INPUT_XML%" ^
  -d "%DICTIONARY%" ^
  -t "%TEMPLATE_FILE%" ^
  -o "%OUTPUT_XML%"

if errorlevel 1 (
    echo ERROR : Step5	 failed
    call "%~dp0common_log.bat" ERROR : Step5 failed
pause
    exit /b 1
)

echo.
echo Step5 finished : %OUTPUT_XML%
call "%~dp0common_log.bat" Step5 finished : %OUTPUT_XML%

pause
exit /b 0