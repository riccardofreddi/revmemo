#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Patch documenti v4.10: '1 mese' -> '30 giorni' (allineato alle altre
finestre), scroll da 10 righe in Home (email di oggi + prossime scadenze),
menu VEDI 20/30/50 RIGHE in Clienti e Registro. LEGGIMI resta CRLF."""
import os

BASE = "/home/z/my-project/download/RevMemo"

# ---------------------------------------------------------------- LEGGIMI
percorso = os.path.join(BASE, "LEGGIMI.txt")
with open(percorso, "r", encoding="utf-8", newline="") as f:
    txt = f.read()
originali = txt

sostituzioni = [
    # 1) COSA FA: "a 1 mese" -> "a 30 giorni"
    ("a 1 mese, a 10 giorni e a",
     "a 30 giorni, a 10 giorni e a"),
    # 1-bis HOME: prossime scadenze con scroll 10 (niente piu' "6 fisse")
    ("EMAIL DA INVIARE OGGI e le prossime scadenze (sempre\r\n"
     "                le 6 piu' vicine: NIENTE SCROLL, quando un cliente fa\r\n"
     "                la revisione entra il successivo; se ce ne sono altre\r\n"
     "                compare \"e altre N\"). Poi i bottoni: PROVA (nessun",
     "EMAIL DA INVIARE OGGI (sono TUTTE: la lista resta alta 10 righe\r\n"
     "                e il resto si raggiunge con la barra a destra) e le\r\n"
     "                prossime scadenze (TUTTE in ordine, anche qui la\r\n"
     "                lista resta alta 10 righe e si scorre con la barra;\r\n"
     "                con poche righe la barra non compare). Poi i\r\n"
     "                bottoni: PROVA (nessun"),
    # 1-bis CLIENTI: menu 20/30/50
    ("                C'e' pure il CERCA per trovare subito un cliente.\r\n",
     "                C'e' pure il CERCA per trovare subito un cliente e il\r\n"
     "                menu VEDI 20/30/50 RIGHE: scegli quante righe vedere\r\n"
     "                prima di scorrere dentro la tabella (l'intestazione\r\n"
     "                resta fissa e la scelta resta memorizzata).\r\n"),
    # 1-bis REGISTRO: menu 20/30/50
    ("                (salva TUTTI i promemoria in registro_invii.xlsx e lo\r\n"
     "                apre). Le statistiche adesso stanno in HOME: sono la\r\n"
     "                prima cosa che si vede quando apri il pannello",
     "                (salva TUTTI i promemoria in registro_invii.xlsx e lo\r\n"
     "                apre), con lo stesso menu VEDI 20/30/50 RIGHE.\r\n"
     "                Le statistiche adesso stanno in HOME: sono la prima\r\n"
     "                cosa che si vede quando apri il pannello"),
]
fatte = 0
for vecchio, nuovo in sostituzioni:
    if vecchio in txt:
        txt = txt.replace(vecchio, nuovo, 1)
        fatte += 1
    else:
        print("ATTENZIONE: blocco non trovato in LEGGIMI.txt:")
        print("   " + vecchio[:70].replace("\r\n", "\\r\\n") + "...")

if txt != originali:
    with open(percorso, "w", encoding="utf-8", newline="") as f:
        f.write(txt)
print(f"LEGGIMI.txt: {fatte}/{len(sostituzioni)} blocchi sostituiti "
      f"({'CRLF conservato' if '\r\n' in txt else 'CRLF PERSO!'})")

# ------------------------------------------------- ISTRUZIONI_CENTRO.md
percorso = os.path.join(BASE, "ISTRUZIONI_CENTRO.md")
with open(percorso, "r", encoding="utf-8", newline="") as f:
    txt = f.read()
if "1 mese, 10 giorni e 1 giorno" in txt:
    txt = txt.replace("1 mese, 10 giorni e 1 giorno",
                      "30 giorni, 10 giorni e 1 giorno", 1)
    with open(percorso, "w", encoding="utf-8", newline="") as f:
        f.write(txt)
    print("ISTRUZIONI_CENTRO.md: aggiornato")
else:
    print("ISTRUZIONI_CENTRO.md: blocco non trovato (da controllare)")

# ------------------------------------------------ GUIDA_INSTALLATORE.md
percorso = os.path.join(BASE, "GUIDA_INSTALLATORE.md")
with open(percorso, "r", encoding="utf-8", newline="") as f:
    txt = f.read()
if '26\u201330 gg ("1 mese")' in txt:
    txt = txt.replace('26\u201330 gg ("1 mese")', '26\u201330 gg ("30 giorni")', 1)
    with open(percorso, "w", encoding="utf-8", newline="") as f:
        f.write(txt)
    print("GUIDA_INSTALLATORE.md: aggiornato")
else:
    print("GUIDA_INSTALLATORE.md: blocco non trovato (da controllare)")

# ----------------------------------------------------------- verifica
with open(os.path.join(BASE, "LEGGIMI.txt"), "r", encoding="utf-8",
          newline="") as f:
    txt = f.read()
problemi = []
if "1 mese" in txt:
    problemi.append("resta '1 mese' in LEGGIMI.txt")
if "\r\n" not in txt:
    problemi.append("CRLF perso in LEGGIMI.txt")
for segno in ("VEDI 20/30/50 RIGHE", "resta alta 10 righe",
              "a 30 giorni, a 10 giorni"):
    if segno not in txt:
        problemi.append("manca: " + segno)
if problemi:
    raise SystemExit("PROBLEMI: " + "; ".join(problemi))
print("VERIFICA DOCUMENTI OK")
