@echo off

chcp 65001
set PYTHONIOENCODING=utf-8

python.exe -m pip install --upgrade pip

start /wait cmd /c npm i --location=global appium

start /wait cmd /c appium driver install uiautomator2

python.exe -m pip install --upgrade pip
pip show Appium-Python-Client >nul 2>&1 || pip install Appium-Python-Client
pip show pandas >nul 2>&1 || pip install pandas
pip show openpyxl >nul 2>&1 || pip install openpyxl
pip show psutil >nul 2>&1 || pip install psutil
pip show hid >nul 2>&1 || pip install hid
pip show pyserial >nul 2>&1 || pip install pyserial

start cmd /k "appium"

python AutoTest_AllMakes.py

pause
