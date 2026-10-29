#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Inserisce la sezione NOVITA' v4.14 in LEGGIMI.txt (mantenendo i CRLF),
sopra la sezione della v4.13. Idempotente: se la sezione c'e' gia', non tocca nulla."""
import sys

PERCORSO = sys.argv[1] if len(sys.argv) > 1 else \
    "/home/z/my-project/revmemo-github/revmemo/LEGGIMI.txt"

SEZIONE = """NOVITA' v4.14 (ottobre 2026)
----------------------------------------------------------------------
- NIENTE FILE CLIENTI ROVINATI: se un invio automatico parte mentre
  stai salvando un cliente dal pannello (o viceversa), le due scritte
  si mettono in fila da sole: una alla volta, con una piccola attesa
  (qualche secondo). Prima potevano sovrapporsi e rovinare il file
  Excel; adesso non puo' piu' succedere.
- Il salvataggio del file clienti e' "tutto o niente": chi legge nel
  frattempo vede il file vecchio o quello nuovo, mai uno scritto a
  meta'. Se il file e' aperto in Excel, l'errore che vedi e' quello
  di sempre ("chiudi Excel e riprova").
- ANCORA MENO RISCHIO DI EMAIL DOPPIE: il blocco anti-doppio-invio ora
  controlla SE l'altro invio sta davvero girando prima di toccare
  qualsiasi cosa. Un invio lento non viene mai interrotto; se invece
  un invio e' morto per strada, il blocco si libera da solo SUBITO
  (prima si aspettavano fino a 15 minuti).
- Se il file clienti risulta occupato, il messaggio lo dice chiaramente
  ("riprova tra qualche secondo") invece di un errore generico.
- La versione del programma si scrive in UN solo posto: quello che
  vedi all'avvio (RevMemo v4.14) e' sempre quello vero.
- Nota: i piccoli file tecnici revmemo.lock ed excel.lock servono
  proprio a queste protezioni: si possono cancellare a programma
  chiuso, il programma se li ricrea da solo.

"""

with open(PERCORSO, "rb") as f:
    grezzo = f.read()

cr_lf = b"\r\n" in grezzo
testo = grezzo.decode("utf-8")
if "NOVITA' v4.14" in testo:
    print("Sezione v4.14 gia' presente: niente da fare.")
    sys.exit(0)

a_capo = "\r\n" if cr_lf else "\n"
sezione = SEZIONE.replace("\n", a_capo)

jalon = "NOVITA' v4.13"
pos = testo.find(jalon)
if pos < 0:
    print("ERRORE: non trovo il punto di inserimento (NOVITA' v4.13).")
    sys.exit(1)

nuovo = testo[:pos] + sezione + testo[pos:]
with open(PERCORSO, "wb") as f:
    f.write(nuovo.encode("utf-8"))
print(f"Sezione NOVITA' v4.14 inserita in {PERCORSO} "
      f"(CRLF={'si' if cr_lf else 'no'}).")
