@echo off
setlocal

REM =========================================================
REM Order-X2DFAX Step2 : 変換先フォーマットXMLの作成
REM 使用例:
REM   call call_order_x2dfax_step2.bat v1_3
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

REM =========================================================
REM パス定義
REM =========================================================

REM 作業フォルダ
set PROCESS_DIR=%WORK_DIR%\order_x2dfax

REM 入出力
set base-xml=D:\github\project\template\Setting_流通BMS基本形Ver1_3\Layout\基本形1_3：発注Ver1_3.xml
set customer-xml=%PROCESS_DIR%\order_x2dfax_step1.xml
set dictionary=%DICTIONARY_DIR%\order_bigboss_dictionary.xlsx
set out=%PROCESS_DIR%\order_x2dfax_step2_mapping_list.xlsx

REM Python
set MAIN_PY=%TOOLS_DIR%\domains\order\order_x2dfax_step2.py

REM =========================================================
REM 事前チェック
REM =========================================================
if not exist "%PROCESS_DIR%" (
    mkdir "%PROCESS_DIR%"
)

if not exist "%MAIN_PY%" (
    echo [ERROR] Pythonファイルが存在しません: %MAIN_PY%
    exit /b 1
)
REM =========================================================
REM 実行
REM =========================================================
echo.
echo ============================================
echo Order-X2DFAX Step2 
echo ============================================
echo [base-xml] %base-xml%
echo [dictionary] %dictionary%
echo [customer-xml] %customer-xml%
echo [out] %out%
echo.

%PYTHON_CMD% "%MAIN_PY%" ^
  --base-xml "%base-xml%" ^
  --dictionary "%dictionary%" ^
  --customer-xml "%customer-xml%" ^
  --out "%out%"

if errorlevel 1 (
    echo [ERROR] Step2 %out%の作成に失敗しました。
pause
    exit /b 1
)

echo.
echo [INFO] Step2 %out%を作成しました。
echo [INFO] %OUTPUT%
echo.

pause
endlocal
exit /b 0