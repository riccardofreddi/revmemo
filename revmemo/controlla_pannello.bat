@echo off
title RevMemo - Controllo del pannello
cd /d "%~dp0"

if not exist "runtime\python.exe" (
  echo.
  echo   PROBLEMA: non trovo la cartella "runtime".
  echo   Estrai TUTTO lo zip - tasto destro sullo zip, "Estrai tutto..."
  echo   - e riprova da dentro la cartella estratta.
  echo.
  pause
  exit /b 1
)

"runtime\python.exe" controlla_pannello.py

echo.
pause
