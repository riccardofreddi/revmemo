@echo off
chcp 65001 >nul
cd /d "%~dp0"
set "PYCMD="
if exist "runtime\python.exe" set "PYCMD=runtime\python.exe"
if not defined PYCMD where python >nul 2>nul && set "PYCMD=python"
if not defined PYCMD where py >nul 2>nul && set "PYCMD=py"
if not defined PYCMD echo Python non trovato: vedi LEGGIMI.txt & pause & exit /b 1
echo.
echo   CONTROLLO DEL PACCHETTO
echo.
%PYCMD% controlla_pc.py
