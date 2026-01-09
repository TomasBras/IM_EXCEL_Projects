@echo off
setlocal
title Web App Assistant
@echo Starting Web App Assistant (Python HTTPS)

rem Prefer bundled Python; fall back to system if missing
set "PYTHON_EXE=%~dp0..\python\python.exe"
if not exist "%PYTHON_EXE%" set "PYTHON_EXE=python"

"%PYTHON_EXE%" server.py

endlocal
