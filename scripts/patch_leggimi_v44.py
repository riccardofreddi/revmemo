# -*- coding: utf-8 -*-
"""Patch LEGGIMI.txt per la v4.4 (personalizzazione totale del messaggio).
Preserva i CRLF; ogni sostituzione deve comparire UNA volta sola."""
import io
import sys

PERCORSO = "/home/z/my-project/download/RevMemo/LEGGIMI.txt"

with io.open(PERCORSO, "r", encoding="utf-8", newline="") as f:
    testo = f.read()

crlf = "\r\n" in testo
print("CRLF:", crlf)


def sostituisci(vecchio, nuovo):
    global testo
    n = testo.count(vecchio)
    if n != 1:
        print(f"ERRORE: il frammento compare {n} volte (attesa 1):")
        print(repr(vecchio[:90]))
        sys.exit(1)
    testo = testo.replace(vecchio, nuovo)


# --- A) sezione 2: riga template_email.html -------------------------
sostituisci(
    "- template_email.html ..... il disegno dell'email (non toccare a mano)",
    "- template_email.html ..... il disegno delle email (grafica); le PAROLE\r\n"
    "                            si scrivono in config.ini sezione [messaggio]",
)

# --- B) sezione 3 passo 2: citare anche il messaggio -----------------
sostituisci(
    "   e compila: mittente, password_app (vedi sezione 5) e i dati del centro\r\n"
    "   nella sezione [centro] (nome, citta', telefono, email_contatti, colore_tema)",
    "   e compila: mittente, password_app (vedi sezione 5) e i dati del centro\r\n"
    "   nella sezione [centro] (nome, citta', telefono, email_contatti, colore_tema).\r\n"
    "   Se vuoi, cambia anche le PAROLE delle email: sezione [messaggio] (vedi 7)",
)

# --- C) sezione 7 riscritta ------------------------------------------
sostituisci(
    "7) IL LOOK DELLE EMAIL\r\n"
    "----------------------------------------------------------------------\r\n"
    "- LOGO: sostituisci il file logo.png con quello del tuo centro\r\n"
    "  (stesso nome esatto, formato PNG). Se Excel/OneDrive non ha finito di\r\n"
    "  sincronizzare, il programma te lo avvisa prima di mandare email mute\r\n"
    "- COLORE: in config.ini alla riga colore_tema (es. #1F3B66 blu,\r\n"
    "  #C8102E rosso, #1B7D46 verde)\r\n"
    "- TESTO: non modificare template_email.html a mano (dentro ci sono\r\n"
    "  segnaposto tipo {TARGA} che il programma riempie): se vuoi cambiare\r\n"
    "  una frase, chiedi a chi ti ha installato il programma",

    "7) IL MESSAGGIO E IL LOOK DELLE EMAIL\r\n"
    "----------------------------------------------------------------------\r\n"
    "LE PAROLE (config.ini, sezione [messaggio]): ogni centro scrive il\r\n"
    "messaggio COME VUOLE, senza toccare il codice. Si possono cambiare:\r\n"
    "- oggetto .............. l'oggetto dell'email\r\n"
    "- oggetto_1mese / oggetto_10giorni / oggetto_1giorno\r\n"
    "                         oggetti speciali per i 3 promemoria\r\n"
    "                         (vuoto = si usa l'oggetto generale)\r\n"
    "- saluto ................ l'apertura (es. \"Gentile {NOME},\")\r\n"
    "- testo ................. il corpo del messaggio (una riga vuota\r\n"
    "                          separa i paragrafi)\r\n"
    "- firma ................. la chiusura\r\n"
    "- riga_extra_piede ...... una riga in piu' nel piede (es. \"Parcheggio\r\n"
    "                          gratuito per i clienti\")\r\n"
    "- testo_disiscrizione ... le istruzioni di cancellazione nel piede\r\n"
    "                          (IMPORTANTE: lascia dentro la parola STOP)\r\n"
    "\r\n"
    "Dentro queste scritte i SEGNAPOSTO si riempiono da soli con i dati\r\n"
    "veri: {NOME} cliente, {TARGA}, {SCADENZA}, {GIORNI}, {TRAFRA}\r\n"
    "(\"tra 29 giorni\"), {NOME_CENTRO}, {CITTA}, {INDIRIZZO}, {TELEFONO},\r\n"
    "{EMAIL_CENTRO}, {SITO_WEB}, {ORARI}, {PRODOTTO}.\r\n"
    "Nel testo: **parola** = grassetto. Es.: scade **entro il {SCADENZA}**.\r\n"
    "ATTENZIONE: le righe sotto \"testo =\" devono iniziare con qualche\r\n"
    "spazio (come gia' sono): non togliere lo spazio iniziale, altrimenti\r\n"
    "config.ini non si legge piu' (il programma te lo spiega se succede).\r\n"
    "\r\n"
    "NUOVI DATI DEL CENTRO (config.ini, sezione [centro]): oltre a nome,\r\n"
    "citta', telefono ed email ora ci sono anche indirizzo, sito_web e\r\n"
    "orari: se li compili appaiono nelle email, se restano vuoti no.\r\n"
    "\r\n"
    "IL LOOK (il disegno, non le parole):\r\n"
    "- LOGO: sostituisci il file logo.png con quello del tuo centro\r\n"
    "  (stesso nome esatto, formato PNG). Se Excel/OneDrive non ha finito di\r\n"
    "  sincronizzare, il programma te lo avvisa prima di mandare email mute\r\n"
    "- COLORE: in config.ini alla riga colore_tema (es. #1F3B66 blu,\r\n"
    "  #C8102E rosso, #1B7D46 verde)\r\n"
    "- GRAFICA: template_email.html e' il disegno (colori, misure, box\r\n"
    "  \"GIORNI RIMANENTI\"). Le PAROLE arrivano da config.ini, quindi NON\r\n"
    "  serve piu' toccarlo per cambiare una frase: si fa tutto in config.ini\r\n"
    "- AGGIORNAMENTO: se il tuo template_email.html e' di una versione\r\n"
    "  precedente, il programma lo porta da solo alla versione nuova al\r\n"
    "  primo avvio (la copia vecchia resta in template_email_vecchio.html)",
)

# --- D) sezione 10 passo 2 -------------------------------------------
sostituisci(
    "2. Apri config.ini: casella Gmail del centro, password per app, dati\r\n"
    "   del centro (nome, citta', telefono, colore)",
    "2. Apri config.ini: casella Gmail del centro, password per app, dati\r\n"
    "   del centro (nome, citta', indirizzo, telefono, sito web, orari,\r\n"
    "   colore) e il MESSAGGIO delle email nella sezione [messaggio]",
)

with io.open(PERCORSO, "w", encoding="utf-8", newline="") as f:
    f.write(testo)

print("Patch LEGGIMI v4.4 applicata: 4 blocchi sostituiti.")
