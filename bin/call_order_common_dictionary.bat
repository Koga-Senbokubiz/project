@echo off

REM ============================================
REM 共通Order項目辞書作成
REM ============================================

call common_env.bat v1_3

set INPUT_FILE=D:\github\project\template\基本メッセージパック_V1_3_20180223\Order_V1_3\Order\Documentations\Schema_Order_20090901.xlsx
rem set OUTPUT_FILE=%DICTIONARY_DIR%\common\bms\order_v1_3_field_dictionary.xlsx
set OUTPUT_FILE=D:\github\project\dictionary\order_v1_3_field_dictionary.xlsx

if not exist "%DICTIONARY_DIR%\common\bms" (
    mkdir "%DICTIONARY_DIR%\common\bms"
)

echo.
echo ============================================
echo Create Order Common Dictionary
echo ============================================

%PYTHON_CMD% %TOOLS_DIR%\common\create_order_common_dictionary.py ^
    "%INPUT_FILE%" ^
    -o "%OUTPUT_FILE%"

if %errorlevel% neq 0 (
    echo ERROR: Failed to create dictionary
    exit /b 1
)

echo SUCCESS: %OUTPUT_FILE%
pause
