@echo off

REM ==========================================
REM Order-X2F Step7
REM Deployフォルダに、Step3～Step6のxmlファイル
REM をコピー
REM ==========================================

call "%~dp0common_env.bat" v1_3

set INPUT_ROOT=%WORK_DIR%\order_x2f
set SETTING_ROOT=C:\Intercom\EasyExchange\Setting

set MAP_NAME=BigBoss_xml2fix

echo on
copy %INPUT_ROOT%\order_x2f_step3_from_layout.xml %SETTING_ROOT%\Layout\%MAP_NAME%_xml.xml
rem copy %INPUT_ROOT%\order_x2f_step4_to_layout.xml %SETTING_ROOT%\Layout\%MAP_NAME%_fix.xml
rem copy %INPUT_ROOT%\order_x2f_step5_logic.xml %SETTING_ROOT%\Logic\%MAP_NAME%_%MAP_NAME%_xml_%MAP_NAME%_fix.xml




pause
exit /b 0
