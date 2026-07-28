@echo off
rem ============================================================
rem  Open all work pages at once (see ????_??.md for why)
rem  ASCII comments only: cmd mis-parses CP949 Hangul in .bat
rem  (some trail bytes are 0x5C) and reports bogus errors.
rem  1) wait for the app server on 8899 - startup order is not
rem     guaranteed, so Chrome must not open before it is up
rem  2) open app + fixed address + ECOUNT + two Band pages
rem  PIN is stored per origin, so localhost:8899 stays logged in.
rem  No password handling here - use Chrome's saved passwords.
rem ============================================================
setlocal

set "APP=http://localhost:8899"
set "FIXED=https://mulder4780.github.io/coupang-ecount-reconcile/app.html"
set "ERP=https://login.ecount.com/Login/"
set "BAND1=https://band.us/band/90610953"
set "BAND2=https://band.us/band/84789192"

set /a tries=0
:wait
powershell -NoProfile -Command "try{ (New-Object Net.Sockets.TcpClient('127.0.0.1',8899)).Close(); exit 0 }catch{ exit 1 }" >nul 2>&1
if not errorlevel 1 goto ready
set /a tries+=1
if %tries% GEQ 30 goto ready
timeout /t 3 /nobreak >nul
goto wait

:ready
set "CHROME=%ProgramFiles%\Google\Chrome\Application\chrome.exe"
if not exist "%CHROME%" set "CHROME=%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"
if not exist "%CHROME%" set "CHROME=%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"
if not exist "%CHROME%" goto fallback

start "" "%CHROME%" --new-window "%APP%" "%FIXED%" "%ERP%" "%BAND1%" "%BAND2%"
goto done

:fallback
start "" "%APP%"
start "" "%FIXED%"
start "" "%ERP%"
start "" "%BAND1%"
start "" "%BAND2%"

:done
endlocal
