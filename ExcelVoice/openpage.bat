@echo off
setlocal enabledelayedexpansion

for /f "tokens=2 delims=:" %%A in ('ipconfig ^| findstr /c:"IPv4 Address"') do (
    set ip=%%A
    set ip=!ip:~1!
)

echo Seu IP é: %ip%
start "" "https://%ip%:8082"
start "" "https://%ip%:8082/appGui.htm"