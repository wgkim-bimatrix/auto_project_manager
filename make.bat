@echo off

pyinstaller -F "������ǥ�ٿ�δ�.py" ^
--add-data "nos_setup.exe;." ^
--add-data "����.json;."

copy nos_setup.exe dist\nos_setup.exe
copy ����.json dist\����.json