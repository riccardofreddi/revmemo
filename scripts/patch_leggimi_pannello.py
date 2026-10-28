#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Aggiorna LEGGIMI.txt con la sezione del PANNELLO (CRLF preservati,
sostituzioni univoche verificate una per una)."""
import io

PERCORSO = "/home/z/my-project/download/RevMemo/LEGGIMI.txt"

with io.open(PERCORSO, "r", encoding="utf-8", newline="") as f:
    t = f.read()

def sostituisci(testo, vecchio, nuovo, nome):
    n = testo.count(vecchio)
    assert n == 1, f"'{nome}': trovate {n} occorrenze (ne serve 1)"
    return testo.replace(vecchio, nuovo)

# --- 1) COSA FA: aggiunta riga pannello
t = sostituisci(t,
    "- Non serve installare niente: tutto e' dentro la cartella. Serve solo Internet.\r\n"
    "\r\n"
    "2) I FILE DI QUESTA CARTELLA",
    "- Non serve installare niente: tutto e' dentro la cartella. Serve solo Internet.\r\n"
    "- NUOVO: il PANNELLO - doppio click su apri_pannello.bat e fai tutto con\r\n"
    "  i pulsanti, senza toccare config.ini (vedi 1-bis).\r\n"
    "\r\n"
    "1-bis) IL PANNELLO: TUTTO CON I PULSANTI (il modo piu' facile)\r\n"
    "----------------------------------------------------------------------\r\n"
    "Doppio click su apri_pannello.bat: si apre una pagina con i pulsanti,\r\n"
    "in italiano, nel tuo browser. Tutto in un posto solo:\r\n"
    " - HOME         la situazione: clienti, chi tocca inviare oggi, prossime\r\n"
    "                scadenze. Poi i bottoni: PROVA (nessun invio), INVIA ORA,\r\n"
    "                Colora Excel e Controllo PC\r\n"
    " - MESSAGGIO    le PAROLE delle email senza Blocco note, con ANTEPRIMA\r\n"
    "                vera dell'email con un cliente di esempio\r\n"
    " - CLIENTI      apre l'Excel con un click (l'Excel resta il file vero)\r\n"
    " - IMPOSTAZIONI email Gmail, Password per app, dati del centro\r\n"
    " - REGISTRO     chi ha ricevuto cosa (ogni promemoria parte 1 sola volta)\r\n"
    "Il pannello gira SOLO su questo computer: nessun dato va online.\r\n"
    "Scrive nello stesso config.ini di sempre: se preferisci il Blocco\r\n"
    "note, config.ini funziona ancora uguale.\r\n"
    "Nota: parte una finestra nera RIDOTTA A ICONA (e' il motore del\r\n"
    "pannello): non chiuderla finche' usi il pannello.\r\n"
    "\r\n"
    "2) I FILE DI QUESTA CARTELLA",
    "sezione 1 + 1-bis")

# --- 2) elenco file
t = sostituisci(t,
    "- avvia_invio.bat ......... INVIO VERO dei promemoria (l'uso di tutti i giorni)\r\n",
    "- avvia_invio.bat ......... INVIO VERO dei promemoria (l'uso di tutti i giorni)\r\n"
    "- apri_pannello.bat ....... il PANNELLO: tutto con i pulsanti (vedi 1-bis)\r\n",
    "elenco file: apri_pannello.bat")

t = sostituisci(t,
    "- revmemo.py .............. il programma (non si tocca mai)\r\n",
    "- revmemo.py .............. il programma (non si tocca mai)\r\n"
    "- pannello.py / pannello.html  il pannello (non si toccano)\r\n",
    "elenco file: pannello.py/html")

# --- 3) LA PRIMA VOLTA: consiglio pannello
t = sostituisci(t,
    "5. Tutto ok? Doppio click su avvia_invio.bat\r\n",
    "5. Tutto ok? Doppio click su avvia_invio.bat\r\n"
    "\r\n"
    "MODO PIU' FACILE: doppio click su apri_pannello.bat -> compili\r\n"
    "Impostazioni e Messaggio con i pulsanti e guardi l'ANTEPRIMA\r\n"
    "dell'email vera prima di inviare (vedi 1-bis).\r\n",
    "sezione 3: modo facile")

# --- 8) USO DI TUTTI I GIORNI
t = sostituisci(t,
    "1. Doppio click su avvia_invio.bat (tutti i giorni o quando vuoi)\r\n"
    "2. Fine: il titolare riceve il riepilogo di cio' che e' partito\r\n",
    "1. Doppio click su apri_pannello.bat -> bottone INVIA ORA\r\n"
    "   (in alternativa: avvia_invio.bat come una volta)\r\n"
    "2. Fine: il titolare riceve il riepilogo di cio' che e' partito\r\n",
    "sezione 8 passo 1")

# --- 10) NUOVO CENTRO
t = sostituisci(t,
    "2. Apri config.ini: casella Gmail del centro, password per app, dati\r\n"
    "   del centro (nome, citta', indirizzo, telefono, sito web, orari,\r\n"
    "   colore) e il MESSAGGIO delle email nella sezione [messaggio]\r\n",
    "2. Casella Gmail del centro, password per app, dati del centro e\r\n"
    "   MESSAGGIO: il modo piu' facile e' apri_pannello.bat -> schede\r\n"
    "   IMPOSTAZIONI e MESSAGGIO (oppure config.ini come una volta)\r\n",
    "sezione 10 passo 2")

t = sostituisci(t,
    "5. Doppio click su avvia_prova.bat -> poi avvia_invio.bat\r\n",
    "5. Doppio click su apri_pannello.bat -> PROVA -> INVIA ORA\r\n"
    "   (oppure avvia_prova.bat e poi avvia_invio.bat)\r\n",
    "sezione 10 passo 5")

with io.open(PERCORSO, "w", encoding="utf-8", newline="") as f:
    f.write(t)
print("LEGGIMI.txt aggiornato:", len(t.split('\r\n')), "righe")
