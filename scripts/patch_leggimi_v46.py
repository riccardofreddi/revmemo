#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Aggiorna LEGGIMI.txt per il pannello v4.6: finestra visibile,
controlla_pannello.bat, blocco Windows (file da internet)."""
import os

PERCORSO = "/home/z/my-project/download/RevMemo/LEGGIMI.txt"
with open(PERCORSO, "r", encoding="utf-8", newline="") as f:
    testo = f.read()

nl = "\r\n" if "\r\n" in testo else "\n"

SOSTITUIRE = [
    # 1-bis: nota sulla finestra nera
    ("Nota: parte una finestra nera RIDOTTA A ICONA (e' il motore del\n"
     "pannello): non chiuderla finche' usi il pannello.",
     "Nota: resta aperta una finestra nera (e' il motore del pannello):\n"
     "puoi ridurla a icona, ma NON chiuderla finche' usi il pannello.\n"
     "Se la pagina non si apre, quella finestra ti mostra il motivo:\n"
     "fai una foto e inviala a chi ti assiste."),
    # elenco file: aggiunge controlla_pannello.bat
    ("- controlla_pc.bat ........ controllo quando qualcosa non va "
     "(te lo dice)",
     "- controlla_pc.bat ........ controllo quando qualcosa non va "
     "(te lo dice)\n"
     "- controlla_pannello.bat .. controllo quando il PANNELLO non si "
     "apre:\n"
     "                            salva il risultato in "
     "assistenza_pannello.txt"),
    # sezione 9: nuovi casi
    ("- La finestra si chiude subito\n"
     "    Doppio click su controlla_pc.bat e leggi cosa dice",
     "- Il PANNELLO non si apre (o la pagina non arriva)\n"
     "    Doppio click su apri_pannello.bat: ora la finestra nera resta\n"
     "    aperta e ti dice il motivo: fai una foto. Se non basta, doppio\n"
     "    click su controlla_pannello.bat: controlla tutto e salva il\n"
     "    risultato in assistenza_pannello.txt (manda quel file o la foto)\n"
     "- \"Windows ha protetto il tuo PC\" o \"Avviso di sicurezza\"\n"
     "    E' la protezione per i file arrivati da internet: clicca\n"
     "    \"Ulteriori informazioni\" e poi \"Esegui comunque\". Il pacchetto\n"
     "    ora sblocca i file da solo alla prima apertura: di solito non\n"
     "    serve fare nulla\n"
     "- La finestra si chiude subito\n"
     "    Doppio click su controlla_pc.bat e leggi cosa dice"),
]

fatte = 0
for vecchio, nuovo in SOSTITUIRE:
    chiave = vecchio.replace("\n", nl)
    valore = nuovo.replace("\n", nl)
    if chiave not in testo:
        raise SystemExit("NON TROVATO: " + vecchio[:60] + "...")
    testo = testo.replace(chiave, valore, 1)
    fatte += 1

with open(PERCORSO, "w", encoding="utf-8", newline="") as f:
    f.write(testo)
print(f"LEGGIMI.txt aggiornato: {fatte} sostituzioni (fine riga "
      f"{'CRLF' if nl == chr(13) + chr(10) else 'LF'})")
