# -*- coding: utf-8 -*-
"""Aggiorna LEGGIMI.txt (CRLF) per RevMemo v4.1:
- regola d'oro: non usare l'email del centro come email di un cliente
- sezione 6: protezioni STOP (email del centro ignorate, citazioni ignorate)
- sezione 9: voce problemi "Ho messo la MIA email su una riga cliente"
"""
import io

PERCORSO = "/home/z/my-project/download/RevMemo/LEGGIMI.txt"

with io.open(PERCORSO, "r", encoding="utf-8", newline="") as f:
    testo = f.read()

# normalizza a \n per lavorare, poi riconverti in \r\n alla fine
unix = testo.replace("\r\n", "\n")

def sostituisci(vecchio, nuovo, etichetta):
    global unix
    if vecchio not in unix:
        raise SystemExit(f"BLOCCO NON TROVATO: {etichetta}")
    if unix.count(vecchio) != 1:
        raise SystemExit(f"BLOCCO NON UNICO: {etichetta}")
    unix = unix.replace(vecchio, nuovo)
    print(f"  [OK] {etichetta}")

# --- 1) regola d'oro nella sezione 4 (I CLIENTI) -----------------------
sostituisci(
    """- Quando il cliente rinnova la revisione: cambia il mese e rimetti
  Revisione Fatta = NO (i promemoria ripartono da soli per il nuovo ciclo)
- Quando hai finito: salva (Ctrl+S) e CHIUDI Excel prima di lanciare
  il programma (se resta aperto, il programma si ferma e te lo dice)""",
    """- Quando il cliente rinnova la revisione: cambia il mese e rimetti
  Revisione Fatta = NO (i promemoria ripartono da soli per il nuovo ciclo)
- NON usare la tua email (quella del centro) come email di un cliente:
  per le prove usa un'altra casella, es. quella di un familiare (vedi
  sezione 6)
- Quando hai finito: salva (Ctrl+S) e CHIUDI Excel prima di lanciare
  il programma (se resta aperto, il programma si ferma e te lo dice)""",
    "regola d'oro email del centro")

# --- 2) sezione 6: protezioni -------------------------------------------
sostituisci(
    """- Da quel momento quel cliente non riceve piu' niente
- Per riattivarlo: nell'Excel rimetti Attivo = S\u00cc
- Nella stessa casella, un cliente con 2 macchine viene messo in pausa
  su entrambe le righe (stesso indirizzo = stessa persona)""",
    """- Da quel momento quel cliente non riceve piu' niente
- Per riattivarlo: nell'Excel rimetti Attivo = S\u00cc
- Nella stessa casella, un cliente con 2 macchine viene messo in pausa
  su entrambe le righe (stesso indirizzo = stessa persona)

Protezioni (importanti, evitano STOP falsi):
- Le email che PARTONO dalla casella del centro stesso sono SEMPRE
  ignorate: non sono risposte vere. Il promemoria contiene la parola
  STOP nel footer, quindi se ti mandi un'email di prova a te stesso
  NON viene mai contata come uno STOP
- Per questo: NON usare la tua email del centro come email di un
  cliente. Per provare lo STOP serve un'altra casella (es. quella di
  un familiare o di un amico): da li' rispondi STOP e rilancia il
  programma
- Nelle risposte conta solo il testo scritto dal cliente: la parte
  citata sotto ("Il giorno ... ha scritto:") viene ignorata. Se il
  cliente risponde "VA BENE GRAZIE" NON entra in STOP, anche se nel
  messaggio citato compare la parola STOP""",
    "protezioni sezione 6")

# --- 3) sezione 9: voce problemi ----------------------------------------
sostituisci(
    """- Voglio re-inviare le stesse email in prova
    Cancella reminders_log.txt (si ricrea da solo al primo invio):
    il programma "dimentica" cio' che ha gia' mandato e riparte
- Ho sbagliato un invio""",
    """- Voglio re-inviare le stesse email in prova
    Cancella reminders_log.txt (si ricrea da solo al primo invio):
    il programma "dimentica" cio' che ha gia' mandato e riparte
- Ho messo la MIA email su una riga cliente e ora c'e' scritto STOP
    Capito con le prime prove (versioni prima della 4.1): apri l'Excel
    e rimetti S\u00cc nella colonna Attivo. Da ora in poi usa un'altra
    email per le prove (vedi sezione 6)
- Ho sbagliato un invio""",
    "voce problemi STOP su riga cliente")

with io.open(PERCORSO, "w", encoding="utf-8", newline="") as f:
    f.write(unix.replace("\n", "\r\n"))

print("LEGGIMI.txt aggiornato (CRLF mantenuto).")
