@echo off
cd /d "C:\Users\danaya.th\OneDrive - Gulf\Documents\GitHub\gulf-dashboard\automation"
python run.py >> "%~dp0run_log.txt" 2>&1
