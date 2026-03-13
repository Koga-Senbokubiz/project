@echo off
setlocal EnableExtensions

REM ==================================================
REM Step4 : match_xml_xsd
REM ==================================================

call "%~dp0common_env.bat" v1_3

set "MESSAGE_TYPE=order"
set "PROCESS=%MESSAGE_TYPE%-x"
set "CALLEE=match_xml_xsd"
set "STEP_NAME=Step4_match_xml_xsd"

set "IN_FILE=%GENERATED_DIR%\%MESSAGE_TYPE%_%VERSION%\%PROCESS%\Step2_xml_path_list.xlsx"
set "XSD_INDEX_FILE=%GENERATED_DIR%\%MESSAGE_TYPE%_%VERSION%\%PROCESS%\Step3_xsd_index.xlsx"
set "OUT_FILE=%GENERATED_DIR%\%MESSAGE_TYPE%_%VERSION%\%PROCESS%\Step4_match_result.xlsx"

call "%~dp0common_log.bat" call_%CALLEE% "[%STEP_NAME%] [Job Started.]"

REM ---- optional: input existence check ----
if not exist "%IN_FILE%" (
    call "%~dp0common_log.bat" call_%CALLEE% "[%STEP_NAME%] [Job Terminated Abnormally.] IN_FILE not found: %IN_FILE%"
    echo.
    echo ERROR: IN_FILE not found
    echo.
    exit /b 1
)
if not exist "%XSD_INDEX_FILE%" (
    call "%~dp0common_log.bat" call_%CALLEE% "[%STEP_NAME%] [Job Terminated Abnormally.] XSD_INDEX_FILE not found: %XSD_INDEX_FILE%"
    echo.
    echo ERROR: XSD_INDEX_FILE not found
    echo.
    exit /b 1
)

%PYTHON_CMD% "%TOOLS_DIR%\%CALLEE%.py" ^
  -i "%IN_FILE%" ^
  -x "%XSD_INDEX_FILE%" ^
  -o "%OUT_FILE%"

if errorlevel 1 (
    call "%~dp0common_log.bat" call_%CALLEE% "[%STEP_NAME%] [Job Terminated Abnormally.]"
    echo.
    echo ERROR: python failed
    echo.
    exit /b 1
)

call "%~dp0common_log.bat" call_%CALLEE% "[%STEP_NAME%] [Job Terminated Successfully.] ===== %OUT_FILE% is Created. ====="

pause
exit /b 0