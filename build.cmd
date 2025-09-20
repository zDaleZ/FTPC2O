@echo off
cd /d %~dp0
rmdir /s /q %cd%\build
rmdir /s /q %cd%\dist
pyinstaller --name ftpc2o --icon=ftp.ico --hidden-import=sv_ttk --add-data "sv_ttk;sv_ttk" --add-data "easy-taskbar-progress.dll;." --add-data "ftp.ico;." -p "%cd%" test.py
pause