setlocal enabledelayedexpansion
@echo off

%~d0
cd "%~dp0"

IF "%1%"=="" GOTO ERROR

echo.
echo ========================================
echo Updating HOSTS file...
echo ========================================
echo.

for /f "useback tokens=*" %%a in ('%1') do set STR=%%~a
for /l %%a in (1,1,5) do if "!STR:~-1!"=="/" set STR=!STR:~0,-1!

copy %systemroot%\System32\drivers\etc\hosts %systemroot%\System32\drivers\etc\hosts.bak

echo. >> %systemroot%\System32\drivers\etc\hosts
echo 127.0.0.1	%STR% >> %systemroot%\System32\drivers\etc\hosts

echo.
echo ========================================
echo Updating HOSTS file done.
echo ========================================
echo.

GOTO FINISH

:ERROR
echo.
echo ========================================
echo Domain parameter missing. 
echo (required for updating the hosts file).
echo ========================================
echo.

:FINISH

