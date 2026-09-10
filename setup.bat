python -m venv venv
venv\pip install -r requirements.txt

set "PROJDIR=%~dp0"
if "%PROJDIR:~-1%"=="\" set "PROJDIR=%PROJDIR:~0,-1%"

(
    echo "INTERPRETER_WIN","%PROJDIR%\venv\Scripts\python.exe"
    echo "PYTHONPATH","%PROJDIR%"
) > xlwings.conf

(
    echo RESTCOUNTRIES_API_KEY="your restcountries api key goes here"
    echo GEONAMES_USERNAME="your geonames username goes here"
) > credentials.txt
