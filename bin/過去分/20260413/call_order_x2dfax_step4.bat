@echo off
setlocal

REM =========================================================
REM Order-X2DFAX step4
REM 使用例:
REM   call call_order_x2dfax_step4.bat v1_3
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
set param_i=%TEMPLATE_DIR%\Setting_流通BMS基本形Ver1_3\Layout\基本形1_3：発注JCA128.xml
set param_d=%DICTIONARY_DIR%\BigBoss_DFAX_変換定義.xlsx
set param_o=%PROCESS_DIR%\bbord_dfax_fax.xml

REM Python
set MAIN_PY=%TOOLS_DIR%\domains\order\order_x2dfax_step4.py

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
echo Order-X2DFAX step4 
echo ============================================
echo [param_i] %param_i%
echo [param_d] %param_d%
echo [param_o] %param_o%
echo.

%PYTHON_CMD% "%MAIN_PY%" ^
  -i "%param_i%" ^
  -d "%param_d%" ^
  -o "%param_o%"

if errorlevel 1 (
    echo [ERROR] step4 %param_o%の作成に失敗しました。
pause
    exit /b 1
)

echo.
echo [INFO] step4 %param_o%を作成しました。
echo [INFO] %OUTPUT%
echo.

pause
endlocal
exit /b 0