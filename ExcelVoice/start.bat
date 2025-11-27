wt --window 0 --title MMI --suppressApplicationTitle -d . -p "Windows Powershell" cmd /k "cd .\mmiframeworkV2\ && ..\Java\bin\java.exe -jar mmiframeworkV2.jar"

wt --window 0 --title FUSION --suppressApplicationTitle -d . -p "Windows Powershell" cmd /k "cd .\FusionEngine\ && ..\Java\bin\java.exe -jar FusionEngine.jar"

wt --window 0 --title RASA --suppressApplicationTitle -d . -p "Windows Powershell" cmd /k ".\Python\python.exe -m rasa run --enable-api -m .\rasaDemo\models\ --cors “*”"

wt --window 0 --title SERVER --suppressApplicationTitle -d . -p "Windows Powershell" cmd /k "cd .\WebAppAssistantV2\ && ..\Python\python.exe server.py"

openpage.bat