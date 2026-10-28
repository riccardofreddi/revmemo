@echo off
title RevMemo - Pannello di controllo
cd /d "%~dp0"

if not exist "runtime\python.exe" (
  echo.
  echo   PROBLEMA: non trovo la cartella "runtime".
  echo   Il pacchetto non e' completo: probabilmente non hai estratto
  echo   tutto lo zip. Fai cosi':
  echo     1. tasto destro sul file RevMemo_v4.zip
  echo     2. "Estrai tutto..."
  echo     3. apri la cartella estratta e riprova da li'
  echo.
  pause
  exit /b 1
)

rem File arrivati da internet: Windows a volte li blocca.
rem Se servisse, li sblocco in silenzio - serve solo la prima volta.
dir /r "runtime\python.exe" 2>nul | find /i "Zone.Identifier" >nul
if not errorlevel 1 (
  powershell -NoProfile -ExecutionPolicy Bypass -Command "Get-ChildItem -LiteralPath '%~dp0' -Recurse -File | Unblock-File" >nul 2>&1
)

echo.
echo  ==============================================================
echo   RevMemo - PANNELLO DI CONTROLLO
echo  ==============================================================
echo   Tra un istante si apre la pagina nel tuo browser.
echo   LASCIA APERTA questa finestra mentre usi il pannello:
echo   puoi ridurla a icona, ma NON chiuderla, perche' chiudendola
echo   si spegne il pannello.
echo   Se la pagina NON si apre, qui sotto compare il motivo:
echo   fai una foto di questa finestra e inviala a chi ti assiste.
echo  ==============================================================
echo.

"runtime\python.exe" pannello.py

if errorlevel 1 (
  echo.
  echo   ERRORE: il pannello si e' fermato.
  echo   Fai una FOTO di questa finestra e inviala a chi ti assiste.
  echo   Aiuta anche il file controlla_pannello.bat: doppio click.
  echo.
  pause
) else (
  echo.
  echo   Pannello chiuso. A presto!
  ping -n 3 127.0.0.1 >nul
)
exit /b 0
