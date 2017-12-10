python replace_ver.py "2017.01"
if exist build rd /s /q build
if exist dist rd /s /q dist
call complie-ui.bat
python setup.py py2exe
rd /s /q build
pause