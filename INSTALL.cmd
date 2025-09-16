@echo off
chcp 65001 >nul
title Office Installer by pollybich
color 0A

:: ─── ASCII-АРТ: p0ly ───
echo.
echo "             _ _       _     _      _     "
echo " _ __   ___ | | |_   _| |__ (_) ___| |__  "
echo "| '_ \ / _ \| | | | | | '_ \| |/ __| '_ \ "
echo "| |_) | (_) | | | |_| | |_) | | (__| | | |"
echo "| .__/ \___/|_|_|\__, |_.__/|_|\___|_| |_|"
echo "|_|              |___/                    "
echo.

echo   Installer Microsoft Office VL
echo   Press [SPACE] to continue...
pause >nul

echo.
echo Installation started...
echo.

cd /d "%~dp0"
setup.exe /configure config.xml

echo.
echo Setting up Office activation via KMS...
echo.

:: Determine the path to ospp.vbs depending on the Office bitness
if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" (
    set "OSPP=%ProgramFiles%\Microsoft Office\Office16\ospp.vbs"
) else (
    set "OSPP=%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs"
)

:: Specify the KMS server
cscript //nologo "%OSPP%" /sethst:172.31.2.34

:: Requesting activation
cscript //nologo "%OSPP%" /act

:: Checking the activation status
cscript //nologo "%OSPP%" /dstatus

echo.
echo Installation and activation complete. Press any key to exit...
pause >nul
exit