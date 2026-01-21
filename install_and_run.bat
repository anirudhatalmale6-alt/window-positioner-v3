@echo off
title Window Positioner Installer
echo ==========================================
echo   Window Positioner v4 - Installer
echo ==========================================
echo.

:: Check if Python is installed
python --version >nul 2>&1
if %errorlevel% == 0 (
    echo [OK] Python is already installed
    goto :install_deps
)

echo [!] Python not found. Downloading Python installer...
echo.

:: Download Python installer
curl -L -o python_installer.exe https://www.python.org/ftp/python/3.12.0/python-3.12.0-amd64.exe
if %errorlevel% neq 0 (
    echo [ERROR] Failed to download Python. Please install manually from python.org
    pause
    exit /b 1
)

echo [!] Installing Python... (This may take a minute)
echo [!] Please wait and DO NOT close this window...
python_installer.exe /quiet InstallAllUsers=0 PrependPath=1 Include_test=0

:: Wait for installation
timeout /t 10 /nobreak >nul

:: Refresh environment
set "PATH=%LOCALAPPDATA%\Programs\Python\Python312;%LOCALAPPDATA%\Programs\Python\Python312\Scripts;%PATH%"

:: Verify installation
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python installation failed. Please install manually from python.org
    pause
    exit /b 1
)

echo [OK] Python installed successfully!
del python_installer.exe

:install_deps
echo.
echo [!] Installing required packages...
pip install pystray pillow keyboard --quiet
if %errorlevel% neq 0 (
    echo [!] Retrying package installation...
    python -m pip install pystray pillow keyboard --quiet
)
echo [OK] Packages installed!

echo.
echo ==========================================
echo   Starting Window Positioner...
echo ==========================================
echo.

:: Run the script
pythonw window_positioner_v3.py
if %errorlevel% neq 0 (
    python window_positioner_v3.py
)

exit
