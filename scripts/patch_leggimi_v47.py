#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Patch LEGGIMI.txt per RevMemo v4.7: pannello rivisto con il cliente
(FERMA/Chiudi, niente Colora Excel, Controllo PC in Impostazioni,
guida al mouse nel Messaggio, SALVA e ANTEPRIMA separati, cerca in
Home/Clienti/Registro, scadenza come da libretto). Mantiene CRLF."""

PERCORSO = "/home/z/my-project/download/RevMemo/LEGGIMI.txt"

VECCHIA_HOME = """ - HOME         la situazione: clienti, chi tocca inviare oggi, prossime
                scadenze. Poi i bottoni: PROVA (nessun invio), INVIA ORA,
                Colora Excel e Controllo PC
 - MESSAGGIO    le PAROLE delle email senza Blocco note, con ANTEPRIMA
                vera dell'email con un cliente di esempio
 - CLIENTI      i clienti si aggiungono e si modificano DA QUI: bottone
                NUOVO CLIENTE e si scrivono nel solito Excel, con
                copia di sicurezza automatica prima di ogni cambiamento.
                Tocchi SÌ/NO per cambiare Revisione Fatta o Attivo (il
                colore si mette da solo), RIATTIVA fa ripartire chi aveva
                risposto STOP, MODIFICA corregge i dati del cliente.
                Se preferisci Excel: bottone APRI EXCEL e lavori da li',
                i due modi restano sempre sincronizzati
 - IMPOSTAZIONI email Gmail, Password per app, dati del centro
 - REGISTRO     chi ha ricevuto cosa (ogni promemoria parte 1 sola volta)"""

NUOVA_HOME = """ - HOME         la situazione: clienti, chi tocca inviare oggi, prossime
                scadenze e il CERCA CLIENTE rapido (scrivi cognome o
                targa). Poi i bottoni: PROVA (nessun invio) e INVIA ORA.
                Mentre un'azione lavora puoi CHIUDERE la finestra e
                continuare a usare il pannello (finisce da sola), oppure
                FERMARLA col pulsante FERMA
 - MESSAGGIO    le PAROLE delle email senza Blocco note, con SALVA e
                ANTEPRIMA separati (l'anteprima NON salva). Passa il
                mouse sopra ogni voce o segnaposto: compare la
                spiegazione di cosa fa
 - CLIENTI      i clienti si aggiungono e si modificano DA QUI: bottone
                NUOVO CLIENTE e si scrivono nel solito Excel, con
                copia di sicurezza automatica prima di ogni cambiamento.
                La scadenza la scrivi come sta sul libretto (es.
                09/2026) e nell'elenco resta esattamente cosi'.
                Tocchi SI'/NO per cambiare Revisione Fatta o Attivo (il
                colore si mette da solo), RIATTIVA fa ripartire chi aveva
                risposto STOP, MODIFICA corregge i dati del cliente.
                C'e' pure il CERCA per trovare subito un cliente.
                Se preferisci Excel: bottone APRI EXCEL e lavori da li',
                i due modi restano sempre sincronizzati
 - IMPOSTAZIONI email Gmail, Password per app, dati del centro e pure
                il CONTROLLO PC (controlla che tutto sia a posto)
 - REGISTRO     chi ha ricevuto cosa (ogni promemoria parte 1 sola volta)
                con il CERCA per email/targa/data (mostra i 250 piu'
                recenti)"""

VECCHIA_COLORA = ("- colora_excel.bat ........ mette i colori nell'Excel (verde/rosso), si usa una volta")
NUOVA_COLORA = ("- colora_excel.bat ........ mette i colori nell'Excel (verde/rosso):\n"
                "                            puoi usarlo tu, dal pannello e' stato tolto")

with open(PERCORSO, "r", encoding="utf-8", newline="") as f:
    testo = f.read()
crlf = "\r\n" in testo
pulito = testo.replace("\r\n", "\n")

fatte = 0
for vecchio, nuovo in ((VECCHIA_HOME, NUOVA_HOME), (VECCHIA_COLORA, NUOVA_COLORA)):
    if vecchio in pulito:
        pulito = pulito.replace(vecchio, nuovo)
        fatte += 1
    elif nuovo in pulito:
        fatte += 1          # gia' applicato
    else:
        print("ATTENZIONE: blocco non trovato:", vecchio.splitlines()[0][:60])

if crlf:
    pulito = pulito.replace("\n", "\r\n")
with open(PERCORSO, "w", encoding="utf-8", newline="") as f:
    f.write(pulito)
print(f"patch applicata: {fatte}/2 blocchi ({'CRLF' if crlf else 'LF'})")
