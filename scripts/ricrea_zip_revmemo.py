#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Ricrea RevMemo_v3.zip completo (pacchetto Pro aggiornato v3.1).
Esclude __pycache__ e file di test: il pacchetto NON contiene ne'
config.ini ne' clienti_revisioni.xlsx (si ricreano da soli al primo
avvio, cosi' aggiornare non cancella mai le impostazioni o i clienti)."""
import os
import zipfile

BASE = "/home/z/my-project/download/RevMemo"
DEST = "/home/z/my-project/download/RevMemo_v3.zip"

if os.path.exists(DEST):
    os.remove(DEST)

n = 0
with zipfile.ZipFile(DEST, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as z:
    for root, dirs, files in os.walk(BASE):
        dirs[:] = [d for d in dirs if d != "__pycache__"]
        dirs.sort()
        for f in sorted(files):
            if f.endswith(".pyc"):
                continue
            full = os.path.join(root, f)
            rel = os.path.join("RevMemo", os.path.relpath(full, BASE))
            z.write(full, rel)
            n += 1

size = os.path.getsize(DEST) / (1024 * 1024)
print(f"ZIP ricreato: {DEST}")
print(f"File inclusi: {n}  |  Dimensione: {size:.1f} MB")

# verifica rapida
errori = []
with zipfile.ZipFile(DEST) as z:
    names = z.namelist()
    for check in ("RevMemo/config_ESEMPIO.ini", "RevMemo/clienti_ESEMPIO.xlsx",
                  "RevMemo/revmemo.py", "RevMemo/controlla_pc.py",
                  "RevMemo/template_email.html", "RevMemo/logo.png",
                  "RevMemo/runtime/python.exe",
                  "RevMemo/runtime/Lib/site-packages/openpyxl/__init__.py"):
        ok = check in names
        print(("OK  " if ok else "MANCA ") + check)
        if not ok:
            errori.append(check)
    # questi NON devono esserci: sono del cliente, si ricreano da soli
    for vietato in ("RevMemo/config.ini", "RevMemo/clienti_revisioni.xlsx",
                    "RevMemo/reminders_log.txt"):
        assente = vietato not in names
        print(("OK  " if assente else "ERRORE! ") + vietato + " assente dal pacchetto")
        if not assente:
            errori.append(vietato)

if errori:
    raise SystemExit("VERIFICA ZIP FALLITA: " + ", ".join(errori))
print("Verifica zip: TUTTO OK")
