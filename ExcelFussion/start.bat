@echo off
setlocal
set "ROOT=%~dp0"
set "JAVA_BIN=%ROOT%Java\bin"
set "JAVA_EXE=%JAVA_BIN%\java.exe"

rem Keep a general PowerShell console in the project root with bundled Java on PATH
start "" powershell -NoExit -Command "cd -LiteralPath '%ROOT%'; $env:JAVA_HOME='%ROOT%Java'; $env:PATH='%JAVA_BIN%;'+$env:PATH"

if not exist "%JAVA_EXE%" (
	echo Java runtime not found at %JAVA_EXE%
	exit /b 1
)

rem Start FusionEngine (prefers its start.bat for quickEdit/title setup)
if exist "%ROOT%FusionEngine" (
	start "" powershell -NoExit -Command "cd -LiteralPath '%ROOT%FusionEngine'; $env:JAVA_HOME='%ROOT%Java'; $env:PATH='%JAVA_BIN%;'+$env:PATH; if (Test-Path './start.bat') { ./start.bat } else { & '%JAVA_EXE%' -jar FusionEngine.jar }"
) else (
	echo Skipping: FusionEngine folder not found.
)

rem Start mmiframeworkV2 (prefers its start.bat for quickEdit/title setup)
if exist "%ROOT%IM" (
	start "" powershell -NoExit -Command "cd -LiteralPath '%ROOT%IM'; $env:JAVA_HOME='%ROOT%Java'; $env:PATH='%JAVA_BIN%;'+$env:PATH; if (Test-Path './start.bat') { ./start.bat } else { & '%JAVA_EXE%' -jar mmiframeworkV2.jar }"
) else (
	echo Skipping: IM folder not found.
)

rem Start the web app assistant
if exist "%ROOT%WebAppAssistantV2" (
	start "" powershell -NoExit -Command "cd -LiteralPath '%ROOT%WebAppAssistantV2'; if (Test-Path './start_web_app.bat') { ./start_web_app.bat } else { Write-Host 'start_web_app.bat missing' }"
) else (
	echo Skipping: WebAppAssistantV2 folder not found.
)

rem Start Rasa server (Anaconda Prompt style)
if exist "%ROOT%rasaDemo" (
	start "" cmd /k ""%USERPROFILE%\anaconda3\Scripts\activate.bat" rasa-env && cd /d "%ROOT%rasaDemo" && rasa run --enable-api -m ./models --cors "*""
) else (
	echo Skipping: rasaDemo folder not found.
)

if exist "%ROOT%openpage.bat" (
	call "%ROOT%openpage.bat"
) else (
	echo Skipping: openpage.bat not found.
)

endlocal
