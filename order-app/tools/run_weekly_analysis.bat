@echo off
rem Beaufield weekly demand analysis (analyze_demand.py)
rem Task Scheduler: py -3.12 recommended. Uses %~dp0 to avoid hardcoded paths.
py -3.12 "%~dp0analyze_demand.py"
