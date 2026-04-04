@echo off
REM setup.bat — Run once to set up the project on Windows
REM Usage: double-click or run from terminal: setup.bat
 
echo.
echo  Setting up mercedes-benz-uk-scraper...
echo.
 
REM Create virtual environment
python -m venv venv
echo  Virtual environment created (venv/)
 
REM Activate and install dependencies
call venv\Scripts\activate.bat
pip install --upgrade pip --quiet
pip install -r requirements.txt --quiet
echo  Dependencies installed
 
REM Create folder structure
if not exist "data\used\chunks" mkdir data\used\chunks
if not exist "data\new\chunks"  mkdir data\new\chunks
if not exist "output"           mkdir output
echo  Folder structure created
 
REM Create .env from example if it doesn't exist
if not exist ".env" (
    copy .env.example .env
    echo  .env file created — open it and paste your TOKEN and COOKIE
) else (
    echo  .env already exists — skipping
)
 
echo.
echo  -----------------------------------------
echo  Setup complete!
echo.
echo  Next steps:
echo    1. Open .env and paste your MB_TOKEN and MB_COOKIE
echo    2. Activate the venv:  venv\Scripts\activate
echo    3. Run:                python scrape_used.py
echo  -----------------------------------------
echo.
pause