@echo off
chcp 65001 >nul
cd /d "%~dp0"
where python >nul 2>nul
if %errorlevel%==0 (set "PYCMD=python") else (set "PYCMD=py")
echo.
echo   CONTROLLO DEL COMPUTER
echo.
%PYCMD% controlla_pc.py
