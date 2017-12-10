::忽略了自动更新，replace_ver.py文件作废
::使用pyinstaller代替原本的py2exe，setup.py文件也用不到了
::实际只处理main.ui生成ui.py，然后打包
if exist build rd /s /q build
if exist dist rd /s /q dist
call complie-ui.bat
pyinstaller -F --clean -i .\favicon.ico main.py
pause