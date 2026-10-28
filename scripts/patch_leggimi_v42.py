# -*- coding: utf-8 -*-
"""Aggiorna LEGGIMI.txt (CRLF) per RevMemo v4.2:
- regola d'oro sezione 4: la propria email va bene come cliente di prova,
  ma gli STOP vanno scritti come RISPOSTA a un promemoria
- sezione 6: 2 modi per provare lo STOP (RISPONDI dalla propria casella
  oppure un'altra casella); le copie automatiche restano ignorate
"""
import io

PERCORSO = "/home/z/my-project/download/RevMemo/LEGGIMI.txt"

with io.open(PERCORSO, "r", encoding="utf-8", newline="") as f:
    testo = f.read()

unix = testo.replace("\r\n", "\n")

def sostituisci(vecchio, nuovo, etichetta):
    global unix
    if vecchio not in unix:
        raise SystemExit(f"BLOCCO NON TROVATO: {etichetta}")
    if unix.count(vecchio) != 1:
        raise SystemExit(f"BLOCCO NON UNICO: {etichetta}")
    unix = unix.replace(vecchio, nuovo)
    print(f"  [OK] {etichetta}")

# --- 1) regola d'oro sezione 4 ------------------------------------------
sostituisci(
    """- NON usare la tua email (quella del centro) come email di un cliente:
  per le prove usa un'altra casella, es. quella di un familiare (vedi
  sezione 6)""",
    """- Per fare da cliente di prova va bene anche la TUA email, ma se
  vuoi provare lo STOP devi RISPONDERE a un promemoria ricevuto
  (vedi sezione 6)""",
    "regola d'oro sezione 4")

# --- 2) sezione 6: protezioni aggiornate v4.2 ----------------------------
sostituisci(
    """Protezioni (importanti, evitano STOP falsi):
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
    """Protezioni (importanti, evitano STOP falsi):
- Le COPIE dei promemoria che ti arrivano nella tua casella (quando
  il programma manda un'email a te stesso) vengono ignorate: contengono
  la parola STOP nel footer ma non sono risposte vere
- Per provare lo STOP hai 2 modi:
  a) dalla TUA casella: apri il promemoria ricevuto, premi RISPONDI,
     scrivi STOP e invia. Deve essere una RISPOSTA (Re:) a un
     promemoria, non un'email nuova; ha effetto se nell'Excel c'e'
     una riga con la tua email
  b) da un'altra casella (es. quella di un familiare o un amico):
     esattamente come farebbe un cliente vero
- Nelle risposte conta solo il testo scritto a mano: la parte citata
  sotto ("Il giorno ... ha scritto:") viene ignorata. Se il cliente
  risponde "VA BENE GRAZIE" NON entra in STOP, anche se nel messaggio
  citato compare la parola STOP""",
    "protezioni sezione 6 v4.2")

with io.open(PERCORSO, "w", encoding="utf-8", newline="") as f:
    f.write(unix.replace("\n", "\r\n"))

print("LEGGIMI.txt aggiornato a v4.2 (CRLF mantenuto).")
