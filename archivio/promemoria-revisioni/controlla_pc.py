#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Controllo del computer per il sistema "Promemoria Revisioni".
Verifica: Python, libreria openpyxl, file Excel, configurazione Gmail.
Avvialo con doppio click su controlla_pc.bat
"""

import importlib
import os
import sys

CARTELLA = os.path.dirname(os.path.abspath(__file__))
OK, MALE = "[OK]  ", "[!!]  "


def riga(sim, testo, dettaglio=""):
    print(f"  {sim}{testo}" + (f"\n        {dettaglio}" if dettaglio else ""))


print("=" * 62)
print("  CONTROLLO DEL COMPUTER – Promemoria Revisioni")
print("=" * 62)

# 1) versione Python
riga(OK, f"Python installato – versione {sys.version.split()[0]}")

# 2) openpyxl
try:
    import openpyxl  # noqa: F401
    riga(OK, "Libreria openpyxl installata (lettura file Excel)")
except ImportError:
    riga(MALE, "Manca la libreria openpyxl",
         "Apri il Prompt dei comandi e digita:  pip install openpyxl")

# 3) file Excel
percorso_excel = os.path.join(CARTELLA, "clienti_revisioni.xlsx")
if os.path.exists(percorso_excel):
    riga(OK, f"File clienti trovato: clienti_revisioni.xlsx")
else:
    riga(MALE, "Non trovo clienti_revisioni.xlsx in questa cartella",
         f"Metti il file in: {CARTELLA}")

# 4) configurazione Gmail (letta dallo script principale)
try:
    pr = importlib.import_module("promemoria_revisioni")
    if pr.MITTENTE_PASSWORD.strip():
        riga(OK, f"Gmail configurato: {pr.MITTENTE_EMAIL}")
        riga(OK, "Password per app inserita")
        passo = "Tutto pronto! Fai doppio click su test_promemoria.bat per la prova."
    else:
        riga(MALE, f"Email impostata: {pr.MITTENTE_EMAIL}")
        riga(MALE, "Manca la Password per app di Google",
             "Vedi ISTRUZIONI.md punto 3: 5 minuti per crearla, poi incollala\n"
             "nello script promemoria_revisioni.py alla riga MITTENTE_PASSWORD")
        passo = "Prossimo passo: crea la Password per app (ISTRUZIONI.md punto 3)."
except Exception as e:
    riga(MALE, f"Non riesco a leggere la configurazione dallo script: {e}")
    passo = "Controlla che promemoria_revisioni.py sia in questa cartella."

print()
print(f"  >> {passo}")
input("\n  Premi INVIO per chiudere...")
