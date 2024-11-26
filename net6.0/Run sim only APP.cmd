@echo off

REM Thiết lập mã ký tự cho Command Prompt sang UTF-8
chcp 65001
set PYTHONIOENCODING=utf-8

python Run_Sim_Only.py

pause
