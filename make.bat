@echo off

pyinstaller -F "매출전표다운로더.py" ^
--add-data "nos_setup.exe;." ^
--add-data "설정.json;."

copy nos_setup.exe dist\nos_setup.exe
copy 설정.json dist\설정.json