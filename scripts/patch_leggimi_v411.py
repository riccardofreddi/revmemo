# -*- coding: utf-8 -*-
"""v4.11: LEGGIMI.txt — menu righe ora 20/30/50/100 (Clienti + Registro)."""

PERCORSO = "/home/z/my-project/download/RevMemo/LEGGIMI.txt"

VECCHIO_A = (
    "                C'e' pure il CERCA per trovare subito un cliente e il\r\n"
    "                menu VEDI 20/30/50 RIGHE: scegli quante righe vedere\r\n"
    "                prima di scorrere dentro la tabella (l'intestazione\r\n"
    "                resta fissa e la scelta resta memorizzata).\r\n"
)
NUOVO_A = (
    "                C'e' pure il CERCA per trovare subito un cliente e il\r\n"
    "                menu VEDI 20/30/50/100 RIGHE: scegli quante righe\r\n"
    "                vedere prima di scorrere dentro la tabella\r\n"
    "                (l'intestazione resta fissa e la scelta resta\r\n"
    "                memorizzata).\r\n"
)
VECCHIO_B = (
    "                apre), con lo stesso menu VEDI 20/30/50 RIGHE.\r\n"
)
NUOVO_B = (
    "                apre), con lo stesso menu VEDI 20/30/50/100 RIGHE.\r\n"
)

with open(PERCORSO, "rb") as f:
    testo = f.read().decode("utf-8")

fatti = 0
if VECCHIO_A in testo:
    testo = testo.replace(VECCHIO_A, NUOVO_A)
    fatti += 1
if VECCHIO_B in testo:
    testo = testo.replace(VECCHIO_B, NUOVO_B)
    fatti += 1

with open(PERCORSO, "wb") as f:
    f.write(testo.encode("utf-8"))

print("sostituzioni fatte:", fatti)
resto = testo.count("20/30/50 ")
print("residui '20/30/50 ' (attesi 0):", resto)
