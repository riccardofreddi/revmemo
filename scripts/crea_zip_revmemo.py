# -*- coding: utf-8 -*-
"""Crea RevMemo_v4.zip dal pacchetto pulito e ne verifica il contenuto."""
import os
import zipfile

SRC = "/home/z/my-project/download/RevMemo"
OUT = "/home/z/my-project/download/RevMemo_v4.zip"

if os.path.exists(OUT):
    os.remove(OUT)

count = 0
with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
    for root, dirs, files in os.walk(SRC):
        dirs[:] = [d for d in dirs if d != "__pycache__"]
        for f in sorted(files):
            p = os.path.join(root, f)
            arc = os.path.relpath(p, os.path.dirname(SRC))
            z.write(p, arc)
            count += 1

with zipfile.ZipFile(OUT) as z:
    nomi = z.namelist()
    richiesti = [
        "RevMemo/revmemo.py", "RevMemo/controlla_pc.py", "RevMemo/LEGGIMI.txt",
        "RevMemo/template_email.html", "RevMemo/logo.png",
        "RevMemo/avvia_invio.bat", "RevMemo/avvia_prova.bat", "RevMemo/controlla_pc.bat",
        "RevMemo/colora_excel.bat",
        "RevMemo/runtime/python.exe", "RevMemo/runtime/python312._pth",
        "RevMemo/runtime/Lib/site-packages/openpyxl/__init__.py",
    ]
    for r in richiesti:
        assert r in nomi, f"MANCA NELLO ZIP: {r}"
    vietati = [n for n in nomi
               if "ESEMPIO" in n or "GUIDA_INSTALLATORE" in n
               or "ISTRUZIONI_CENTRO" in n or n.endswith("config.ini")
               or "clienti_revisioni" in n or "reminders_log" in n
               or "__pycache__" in n]
    assert not vietati, f"FILE VIETATI NELLO ZIP: {vietati}"

print(f"OK: {OUT}")
print(f"File: {count}   Dimensione: {os.path.getsize(OUT) / 1024 / 1024:.1f} MB")
print("Verifica: file chiave presenti, nessun file ESEMPIO/config/log nel pacchetto")
