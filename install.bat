@echo off
pushd "%~dp0"
py -3 -m pip install --user --quiet -r requirements.txt
py -3 install.py
pause
