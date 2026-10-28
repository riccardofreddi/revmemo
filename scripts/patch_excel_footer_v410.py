#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""v4.10: aggiorna la nota di istruzioni dentro i file Excel di esempio:
'Promemoria a 1 mese' -> 'Promemoria a 30 giorni' (allineato al pannello)."""
import openpyxl

FILE = ["/home/z/my-project/download/RevMemo/clienti_revisioni.xlsx",
        "/home/z/my-project/download/RevMemo/clienti_ESEMPIO.xlsx"]

for percorso in FILE:
    wb = openpyxl.load_workbook(percorso)
    ws = wb.active
    trovato = False
    for riga in ws.iter_rows():
        for cella in riga:
            if isinstance(cella.value, str) and "Promemoria a 1 mese" in cella.value:
                cella.value = cella.value.replace("Promemoria a 1 mese",
                                                  "Promemoria a 30 giorni")
                trovato = True
    if trovato:
        wb.save(percorso)
        print(f"{percorso.split('/')[-1]}: nota aggiornata")
    else:
        print(f"{percorso.split('/')[-1]}: 'Promemoria a 1 mese' non trovato "
              "(da controllare)")
    del wb
