@echo off
setlocal
REM =========================================================
REM Order-X2F Step3 : 顧客マッピング原形の作成
REM 使用例:
REM   call call_order_x2f_step3.bat v1_3
REM =========================================================

REM ----- VERSION引数受け取り（未指定ならデフォルト） -----
set DEF_VERSION=v1_3
set VERSION=%~1
if "%VERSION%"=="" set VERSION=%DEF_VERSION%

REM ----- VERSION正規化（v1.3 -> v1_3） -----
set VERSION=%VERSION:.=_%

REM ----- 共通環境変数読込 -----
call common_env.bat %VERSION%
if errorlevel 1 (
    echo [ERROR] common_env.bat の呼び出しに失敗しました。
    exit /b 1
)

set PRJ=bbord

copy "%TEMPLATE_DIR%\Setting_流通BMS基本形Ver1_3\Layout\基本形1_3：発注Ver1_3.xml" "%EE_SETTING_DIR%\Layout\bbord_xml.xml"
copy "%TEMPLATE_DIR%\Setting_流通BMS基本形Ver1_3\Layout\基本形1_3：発注JCA128.xml" "%EE_SETTING_DIR%\Layout\bbord_jca128.xml"
copy "%TEMPLATE_DIR%\Setting_流通BMS基本形Ver1_3\基本形1_3：発注Ver1_3（XML→JCA128）.xml" "%EE_SETTING_DIR%\bbord_xml-jca128.xml"
copy "%TEMPLATE_DIR%\Setting_流通BMS基本形Ver1_3\Logic\基本形1_3：発注Ver1_3（XML→JCA128）_基本形1_3：発注Ver1_3_基本形1_3：発注JCA128.xml" ^
 "%EE_SETTING_DIR%\Logic\bbord_xml-jca128_bbord_xmlbbord_jca128.xml"

echo.
echo [INFO] Step3 終了
echo.

pause
endlocal
exit /b 0