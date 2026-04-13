@echo off
setlocal

REM =========================================================
REM Order-X2F Step2 : 差異表の作成
REM 使用例:
REM   call call_order_x2f_step2.bat v1_3
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
set PROCESS_DIR=%WORK_DIR%\order_x2f

REM 入出力
set INPUT_XML=%PROCESS_DIR%\order_x2f_step1.xml
set INPUT_DICT=%DICTIONARY_DIR%\order_bigboss_dictionary.xlsx
set OUTPUT_XLSX=%PROCESS_DIR%\order_x2f_step2_diff_report.xlsx

REM Python
set MAIN_PY=%TOOLS_DIR%\domains\order\order_x2f_step2.py

REM =========================================================
REM 事前チェック
REM =========================================================
if not exist "%PROCESS_DIR%" (
    mkdir "%PROCESS_DIR%"
)

if not exist "%INPUT_XML%" (
    echo [ERROR] 入力XMLが存在しません: %INPUT_XML%
    exit /b 1
)

if not exist "%INPUT_DICT%" (
    echo [ERROR] 顧客項目辞書が存在しません: %INPUT_DICT%
    exit /b 1
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
echo Create Order-X2F Step2 Diff Table
echo ============================================
echo [INPUT XML ] %INPUT_XML%
echo [DICTIONARY] %INPUT_DICT%
echo [OUTPUT    ] %OUTPUT_XLSX%
echo.

%PYTHON_CMD% "%MAIN_PY%" ^
    "%INPUT_XML%" ^
    "%INPUT_DICT%" ^
    -o "%OUTPUT_XLSX%"

if errorlevel 1 (
    echo [ERROR] Step2 差異表の作成に失敗しました。
pause
    exit /b 1
)

echo.
echo [INFO] Step2 差異表を作成しました。
echo [INFO] %OUTPUT_XLSX%
echo.

pause
endlocal
exit /b 0