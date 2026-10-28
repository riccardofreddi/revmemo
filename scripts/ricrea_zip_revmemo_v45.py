#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Ricrea RevMemo_v4.zip (v4.5 con PANNELLO). Esclude __pycache__ e file
del cliente: niente config.ini, niente clienti_revisioni.xlsx, niente
reminders_log.txt (si ricreano da soli al primo avvio)."""
import os
import zipfile

BASE = "/home/z/my-project/download/RevMemo"
DEST = "/home/z/my-project/download/RevMemo_v4.zip"

if os.path.exists(DEST):
    os.remove(DEST)

n = 0
ESCLUDI_FILE = {"config.ini", "clienti_revisioni.xlsx",
                "reminders_log.txt", "pannello_log.txt",
                "registro_invii.xlsx",
                "assistenza_pannello.txt", "template_email_vecchio.html"}
with zipfile.ZipFile(DEST, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as z:
    for root, dirs, files in os.walk(BASE):
        dirs[:] = [d for d in dirs if d != "__pycache__"
                   and d.lower() != "backup"]
        dirs.sort()
        for f in sorted(files):
            if f.endswith(".pyc") or f in ESCLUDI_FILE:
                continue
            full = os.path.join(root, f)
            rel = os.path.join("RevMemo", os.path.relpath(full, BASE))
            z.write(full, rel)
            n += 1

size = os.path.getsize(DEST) / (1024 * 1024)
print(f"ZIP ricreato: {DEST}")
print(f"File inclusi: {n}  |  Dimensione: {size:.1f} MB")

errori = []
with zipfile.ZipFile(DEST) as z:
    names = z.namelist()
    print("-- file chiave:")
    for check in ("RevMemo/revmemo.py", "RevMemo/pannello.py",
                  "RevMemo/pannello.html", "RevMemo/apri_pannello.bat",
                  "RevMemo/controlla_pc.py", "RevMemo/controlla_pannello.py",
                  "RevMemo/controlla_pannello.bat",
                  "RevMemo/template_email.html",
                  "RevMemo/logo.png", "RevMemo/LEGGIMI.txt",
                  "RevMemo/runtime/python.exe",
                  "RevMemo/runtime/Lib/site-packages/openpyxl/__init__.py"):
        ok = check in names
        print(("OK   " if ok else "MANCA ") + check)
        if not ok:
            errori.append(check)
    print("-- file del cliente (devono MANCARE):")
    for vietato in ("RevMemo/config.ini", "RevMemo/clienti_revisioni.xlsx",
                    "RevMemo/reminders_log.txt", "RevMemo/__pycache__"):
        assente = vietato not in names
        print(("OK   " if assente else "ERRORE! ") + vietato + " assente")
        if not assente:
            errori.append(vietato)
    print("-- integrita':", "OK" if z.testzip() is None else "CORROTTO")

if errori:
    raise SystemExit("PROBLEMI NELLO ZIP: " + str(errori))
print("ZIP OK")
