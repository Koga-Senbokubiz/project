REM ============================================
REM 共通ログ出力バッチ
REM 使い方:
REM call common_log.bat 処理名 メッセージ
REM [python]subprocess.call(["common_log.bat", "処理名", "メッセージ"])
REM ============================================

REM 引数チェック
if "%~1"=="" exit /b 1
if "%~2"=="" exit /b 1

REM ログディレクトリがない場合作成する。
if not exist "%LOG_DIR%" (
    mkdir "%LOG_DIR%"
)

REM 日付・時刻取得（YYYYMMDD / HH:MM:SS）
set DATE_STR_F=%DATE:~0,4%%DATE:~5,2%%DATE:~8,2%
set DATE_STR=%DATE:~0,4%/%DATE:~5,2%/%DATE:~8,2%
set TIME_STR=%TIME:~0,2%:%TIME:~3,2%:%TIME:~6,2%

REM 先頭の空白対策
set TIME_STR=%TIME_STR: =0%

REM ログファイル名（日別）
set LOG_FILE=%LOG_DIR%\common_%DATE_STR_F%.log

REM ログ出力
echo [%DATE_STR% %TIME_STR%] [%~1] %~2>>"%LOG_FILE%"

