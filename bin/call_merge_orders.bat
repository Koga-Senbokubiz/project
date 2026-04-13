@echo off
setlocal

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
set MAIN_PY=%TOOLS_DIR%\common\merge_orders_for_ee.py

if not exist "%MAIN_PY%" (
    echo [ERROR] Pythonファイルが存在しません: %MAIN_PY%
    exit /b 1
)

%PYTHON_CMD% "%MAIN_PY%" ^
C:\Intercom\EasyExchange\Setting\InputFiles\J9682128\FAX ^
C:\Intercom\EasyExchange\Setting\InputFiles\J9682128_marged.xml
if errorlevel 1 (
    echo [ERROR] Step2 %out%の作成に失敗しました。
pause
    exit /b 1
)

echo.
echo [INFO] marge fileを作成しました。
echo.

pause
endlocal
exit /b 0