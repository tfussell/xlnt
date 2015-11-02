@echo off
setlocal EnableDelayedExpansion
for /f %%i in ('where python') DO (set PYTHON=%%i) & goto :done1
:done1
@where python3 > nul 2>&1 && for /f %%i in ('@where python3') DO (@set PYTHON=%%i) & goto :done2
:done2
!PYTHON! configure