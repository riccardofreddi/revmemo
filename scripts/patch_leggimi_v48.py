#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Patch LEGGIMI.txt per RevMemo v4.8: Chiudi anche a fine azione e su
Controllo PC, card 'Email da inviare oggi', statistiche mensili ed
export Excel del Registro, caselle email qualsiasi (Server/Porta).
Mantiene CRLF."""

PERCORSO = "/home/z/my-project/download/RevMemo/LEGGIMI.txt"

VECCHIA_HOME = """ - HOME         la situazione: clienti, chi tocca inviare oggi, prossime
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

NUOVA_HOME = """ - HOME         la situazione: clienti, le EMAIL DA INVIARE OGGI, le
                prossime scadenze (sempre le 6 piu' vicine, non si
                allungano da sole) e il CERCA CLIENTE rapido (scrivi
                cognome o targa). Poi i bottoni: PROVA (nessun invio) e
                INVIA ORA. Mentre un'azione lavora puoi FERMARLA col
                pulsante FERMA o nasconderla con CHIUDI: il risultato
                resta visibile finche' non lo chiudi tu, anche dopo che
                e' finito
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
                colore si mette da solo, centrati sotto la loro
                colonna), RIATTIVA fa ripartire chi aveva risposto STOP,
                MODIFICA corregge i dati del cliente.
                C'e' pure il CERCA per trovare subito un cliente.
                Se preferisci Excel: bottone APRI EXCEL e lavori da li',
                i due modi restano sempre sincronizzati
 - IMPOSTAZIONI la casella email (GMAIL o un'altra: Outlook, Libero,
                Virgilio, Aruba... con Server e Porta), la Password per
                app, i dati del centro e pure il CONTROLLO PC (con il
                suo bottone CHIUDI per togliere il risultato)
 - REGISTRO     chi ha ricevuto cosa (ogni promemoria parte 1 sola volta)
                con il CERCA per email/targa/data, le STATISTICHE
                mensili degli invii (grafico degli ultimi 6 mesi) e
                ESPORTA IN EXCEL (salva TUTTI i promemoria in
                registro_invii.xlsx e lo apre)"""

VECCHIA_PWD = """Questa stessa password serve anche per leggere le risposte STOP
(sezione 6). La puoi cancellare in qualsiasi momento da quella pagina:
e' un accesso dedicato al programma, non la tua password generale."""

NUOVA_PWD = """Questa stessa password serve anche per leggere le risposte STOP
(sezione 6). La puoi cancellare in qualsiasi momento da quella pagina:
e' un accesso dedicato al programma, non la tua password generale.

NON USI GMAIL? Nessun problema: RevMemo funziona con QUALSIASI casella
email. Dal PANNELLO -> Impostazioni -> carta "Casella email diversa da
Gmail" compila SERVER e PORTA (esempi: Outlook smtp.office365.com 587,
Libero smtp.libero.it 465, Virgilio smtp.virgilio.it 465, Aruba
smtp.aruba.it 465). Da Blocco note e' la stessa cosa: sezione [smtp]
righe host e porta. Se sbagli server o porta, il programma te lo dice
al primo invio: basta correggere e riprovare."""

with open(PERCORSO, "r", encoding="utf-8", newline="") as f:
    testo = f.read()
crlf = "\r\n" in testo
pulito = testo.replace("\r\n", "\n")

fatte = 0
for vecchio, nuovo in ((VECCHIA_HOME, NUOVA_HOME), (VECCHIA_PWD, NUOVA_PWD)):
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
print(f"patch v4.8 applicata: {fatte}/2 blocchi ({'CRLF' if crlf else 'LF'})")
