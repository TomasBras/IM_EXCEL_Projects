cd python
.\python.exe .\get-pip.py --no-warn-script-location
.\python.exe -m pip install --target .\Lib\site-packages pywin32 docopt-0.6.2-py2.py3-none-any.whl
.\python.exe -m pip install rasa --no-warn-script-location