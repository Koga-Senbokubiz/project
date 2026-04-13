@echo off
setlocal

REM =========================================================
REM Order-X2DFAX Step3
REM 使用例:
REM   call call_order_x2dfax_Step3.bat v1_3
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
set step2-xlsx=%PROCESS_DIR%\order_x2dfax_Step2_field_list.xlsx
set dfax-definition-xlsx=%DICTIONARY_DIR%\BigBoss_DFAX_変換定義.xlsx
set out=%PROCESS_DIR%\order_x2dfax_Step3_mapping_list.xlsx

REM Python
set MAIN_PY=%TOOLS_DIR%\domains\order\order_x2dfax_Step3.py

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
echo Order-X2DFAX Step3 
echo ============================================
echo [step2-xlsx] %step2-xlsx%
echo [dfax-definition-xlsx] %dfax-definition-xlsx%
echo [out] %out%
echo.

%PYTHON_CMD% "%MAIN_PY%" ^
  --step2-xlsx "%step2-xlsx%" ^
  --dfax-definition-xlsx "%dfax-definition-xlsx%" ^
  --out "%out%"

if errorlevel 1 (
    echo [ERROR] Step3 %out%の作成に失敗しました。
pause
    exit /b 1
)

echo.
echo [INFO] Step3 %out%を作成しました。
echo [INFO] %OUTPUT%
echo.

pause
endlocal
exit /b 0