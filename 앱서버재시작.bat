@echo off
rem 앱 서버를 끄고 다시 띄운다. 코드를 고친 뒤에는 이것을 눌러야 화면이 바뀐다.
rem (앱서버실행.bat 은 띄우기만 해서, 이미 떠 있으면 옛 서버가 계속 응답한다)
set PYTHONIOENCODING=utf-8
"%LOCALAPPDATA%\Programs\Python\Python312\python.exe" "%~dp0webapp\restart_server.py"
pause
