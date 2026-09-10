@echo off
python -m venv venv
venv\Scripts\pip install -r requirements.txt > setup_log.txt 2>&1

if %errorlevel% neq 0 (
    echo Something went wrong during pip install. Check setup_log.txt for details.
) else (
    echo Installation of pip packages into venv completed successfully.
)

set "PROJDIR=%~dp0"
if "%PROJDIR:~-1%"=="\" set "PROJDIR=%PROJDIR:~0,-1%"

(
    echo "INTERPRETER_WIN","%PROJDIR%\venv\Scripts\python.exe"
    echo "PYTHONPATH","%PROJDIR%"
) > xlwings.conf

REM credentials.txt: only create if missing, so real values already entered aren't overwritten
if not exist credentials.txt (
    (
        echo GEONAMES_USERNAME="your geonames username goes here"
        echo RESTCOUNTRIES_API_KEY="your restcountries api key goes here"
    ) > credentials.txt
    echo Created credentials.txt with placeholder values - please enter your real credentials!
) else (
    echo credentials.txt already exists - leaving it untouched!
)

pause