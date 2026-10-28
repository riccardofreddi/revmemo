#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Patch v4.9 a LEGGIMI.txt (CRLF): statistiche in Home con mesi dal
primo utilizzo, niente Cerca in Home, niente scroll nelle prossime
scadenze ("e altre N"), menu provider email in Impostazioni, campo
"Risposte verso" tolto (STOP sempre letto), statistiche fuori dal
Registro."""
import sys

PERCORSO = "/home/z/my-project/download/RevMemo/LEGGIMI.txt"
with open(PERCORSO, encoding="utf-8", newline="") as f:
    testo = f.read()

NL = "\r\n" if "\r\n" in testo else "\n"

def blocco(*righe):
    return NL.join(righe)

VECCHIO_HOME = blocco(
    " - HOME         la situazione: clienti, le EMAIL DA INVIARE OGGI, le",
    "                prossime scadenze (sempre le 6 piu' vicine, non si",
    "                allungano da sole) e il CERCA CLIENTE rapido (scrivi",
    "                cognome o targa). Poi i bottoni: PROVA (nessun invio) e",
    "                INVIA ORA. Mentre un'azione lavora puoi FERMARLA col",
    "                pulsante FERMA o nasconderla con CHIUDI: il risultato",
    "                resta visibile finche' non lo chiudi tu, anche dopo che",
    "                e' finito")

NUOVO_HOME = blocco(
    " - HOME         la situazione: clienti, le STATISTICHE DEGLI INVII",
    "                (grafico mese per mese: parte dal mese del PRIMO",
    "                INVIO, senza colonne vuote dei mesi mai usati), le",
    "                EMAIL DA INVIARE OGGI e le prossime scadenze (sempre",
    "                le 6 piu' vicine: NIENTE SCROLL, quando un cliente fa",
    "                la revisione entra il successivo; se ce ne sono altre",
    "                compare \"e altre N\"). Poi i bottoni: PROVA (nessun",
    "                invio) e INVIA ORA. Mentre un'azione lavora puoi",
    "                FERMARLA col pulsante FERMA o nasconderla con CHIUDI:",
    "                il risultato resta visibile finche' non lo chiudi tu,",
    "                anche dopo che e' finito")

VECCHIO_IMP_REG = blocco(
    " - IMPOSTAZIONI la casella email (GMAIL o un'altra: Outlook, Libero,",
    "                Virgilio, Aruba... con Server e Porta), la Password per",
    "                app, i dati del centro e pure il CONTROLLO PC (con il",
    "                suo bottone CHIUDI per togliere il risultato)",
    " - REGISTRO     chi ha ricevuto cosa (ogni promemoria parte 1 sola volta)",
    "                con il CERCA per email/targa/data, le STATISTICHE",
    "                mensili degli invii (grafico degli ultimi 6 mesi) e",
    "                ESPORTA IN EXCEL (salva TUTTI i promemoria in",
    "                registro_invii.xlsx e lo apre)")

NUOVO_IMP_REG = blocco(
    " - IMPOSTAZIONI la casella email (GMAIL o un'altra: scegli il nome",
    "                dal MENU - Gmail, Outlook/Hotmail, Libero, Virgilio,",
    "                Aruba - e Server e Porta si scrivono DA SOLI; col",
    "                tuo fornitore non c'e'? Scegli \"Altro\" e scrivili",
    "                tu), la Password per app, i dati del centro e pure il",
    "                CONTROLLO PC (con il suo bottone CHIUDI). Le risposte",
    "                dei clienti (anche lo STOP) arrivano sempre nella",
    "                casella del centro: il programma le legge da solo",
    " - REGISTRO     chi ha ricevuto cosa (ogni promemoria parte 1 sola volta)",
    "                con il CERCA per email/targa/data ed ESPORTA IN EXCEL",
    "                (salva TUTTI i promemoria in registro_invii.xlsx e lo",
    "                apre). Le statistiche adesso stanno in HOME: sono la",
    "                prima cosa che si vede quando apri il pannello")

fatte = 0
for vecchio, nuovo in ((VECCHIO_HOME, NUOVO_HOME),
                       (VECCHIO_IMP_REG, NUOVO_IMP_REG)):
    if vecchio in testo:
        testo = testo.replace(vecchio, nuovo, 1)
        fatte += 1
    else:
        print("BLOCCO NON TROVATO:")
        print(vecchio[:120])
        sys.exit(1)

with open(PERCORSO, "w", encoding="utf-8", newline="") as f:
    f.write(testo)
print(f"[OK] LEGGIMI.txt patchato v4.9 ({fatte} blocchi), CRLF={NL == chr(13) + chr(10)}")
