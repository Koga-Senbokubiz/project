REM ============================================
REM 共通変数設定バッチ
REM 使い方:
REM   call common_env.bat v1_3
REM   call common_env.bat v2_0
REM ============================================
set DEF_VERSION=v1_3

REM ---- VERSION 引数受け取り（未指定ならデフォルト） ----
set VERSION=%~1
if "%VERSION%"=="" set VERSION=%DEF_VERSION%

REM ---- VERSION 正規化（v1.3 -> v1_3） ----
set VERSION=%VERSION:.=_%

REM ----- Ver形式作成 -----
set VERSION_DIR=Ver%VERSION:~1%

REM ---- Python I/O 文字コード（ログ文字化け対策） ----
set PYTHONUTF8=1

REM ---- 環境変数定義 ----
set PROJECT_ROOT=D:\github\project
set DICTIONARY_DIR=%PROJECT_ROOT%\dictionary
set SCHEMA_DIR=%PROJECT_ROOT%\schema
set TOOLS_DIR=%PROJECT_ROOT%\tools
set GENERATED_DIR=%PROJECT_ROOT%\generated
set TEMPLATE_DIR=%PROJECT_ROOT%\template
set LOG_DIR=%PROJECT_ROOT%\logs
set DATA_DIR=%PROJECT_ROOT%\data
set WORK_DIR=%PROJECT_ROOT%\work
set PYTHON_CMD=python
echo.
echo Job Started At %date% %time%
echo.
