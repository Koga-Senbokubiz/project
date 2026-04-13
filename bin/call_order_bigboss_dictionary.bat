@echo off
setlocal

REM =========================================================
REM 顧客Order項目辞書の作成（BigBoss）
REM =========================================================
REM 使用方法:
REM   call call_order_bigboss_dictionary.bat v1_3
REM   call call_order_bigboss_dictionary.bat v2_0
REM =========================================================

REM ----- VERSION引数受け取り（未指定ならデフォルト） -----
set DEF_VERSION=v1_3
set VERSION=%~1
if "%VERSION%"=="" set VERSION=%DEF_VERSION%

REM ----- VERSION正規化（v1.3 -> v1_3） -----
set VERSION=%VERSION:.=_%

REM ----- 共通環境変数読込 -----
rem call common_env.bat %VERSION%
call common_env.bat v1_3
if errorlevel 1 (
    echo [ERROR] common_env.bat の呼び出しに失敗しました。
    exit /b 1
)

REM =========================================================
REM パス定義
REM =========================================================
set INPUT_XML=%WORK_DIR%\order_x2f\order_x2f_step1.xml
set COMMON_DICT=D:\github\project\dictionary\order_%VERSION%_field_dictionary.xlsx
set OUTPUT_DICT=D:\github\project\dictionary\order_bigboss_dictionary.xlsx

set MAIN_PY=%TOOLS_DIR%\common\create_order_bigboss_dictionary.py

REM =========================================================
REM 事前チェック
REM =========================================================
if not exist "%INPUT_XML%" (
    echo [ERROR] 入力XMLが存在しません: %INPUT_XML%
    exit /b 1
)

if not exist "%COMMON_DICT%" (
    echo [ERROR] 共通辞書が存在しません: %COMMON_DICT%
    exit /b 1
)

if not exist "%MAIN_PY%" (
    echo [ERROR] Pythonファイルが存在しません: %MAIN_PY%
    exit /b 1
)

rem if not exist "%CUSTOMER_DICT_DIR%" (
rem     mkdir "%CUSTOMER_DICT_DIR%"
rem )

REM =========================================================
REM 実行
REM =========================================================
echo.
echo ============================================
echo Create BigBoss Order Customer Dictionary
echo ============================================
echo [INPUT XML ] %INPUT_XML%
echo [COMMON    ] %COMMON_DICT%
echo [OUTPUT    ] %OUTPUT_DICT%
echo.

%PYTHON_CMD% "%MAIN_PY%" ^
    "%INPUT_XML%" ^
    "%COMMON_DICT%" ^
    -o "%OUTPUT_DICT%"

if errorlevel 1 (
    echo [ERROR] 顧客Order項目辞書の作成に失敗しました。
pause
    exit /b 1
)

rem echo.
rem echo [INFO] 顧客Order項目辞書を作成しました。
rem echo [INFO] %OUTPUT_DICT%
rem echo.
pause
endlocal
exit /b 0