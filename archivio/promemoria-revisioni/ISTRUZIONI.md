# Promemoria Revisioni Automatici — ISTRUZIONI (versione 2)

Sistema per ricordare ai clienti la scadenza della revisione della macchina:
file **Excel** con i clienti + script **Python** che manda le email di promemoria
**1 mese**, **10 giorni** e **1 giorno** prima della scadenza.

---

## 1. Cosa c'è nel pacchetto

| File | A cosa serve |
|------|--------------|
| `clienti_revisioni.xlsx` | L'elenco dei clienti: nome, email, targa, mese di scadenza, revisione fatta, attivo |
| `controlla_pc.bat` | **Doppio click**: controlla se Python, openpyxl e Gmail sono a posto e ti dice il prossimo passo |
| `test_promemoria.bat` | Doppio click = **prova**: mostra cosa verrebbe inviato, SENZA inviare nulla |
| `avvia_promemoria.bat` | Doppio click = **invio reale** dei promemoria del giorno |
| `promemoria_revisioni.py` | Lo script che controlla le scadenze e invia le email |
| `ISTRUZIONI.md` | Questo file |
| `reminders_log.txt` | Si crea da solo al primo invio: registro anti-doppioni |

> I clienti nel file Excel sono **dati di esempio**: sostituiscili con i tuoi.

---

## 2. Controllare il computer (dove sta Python)

Non serve cercare nulla a mano: fai doppio click su **`controlla_pc.bat`**.
Ti dice se Python e openpyxl sono a posto e qual è il prossimo passo.

Se vuoi vedere a mano dove sta Python:
1. Premi Start, scrivi **cmd**, apri il Prompt dei comandi
2. Scrivi: `where python` e INVIO → ti mostra il percorso (es. `C:\Users\riccardo\AppData\Local\Programs\Python\Python313\python.exe`)
3. Se dicessi "non è riconosciuto come comando", prova `py --version` (i file .bat provano da soli sia `python` sia `py`)

Se mancasse openpyxl (lo script te lo dirà), nel Prompt dei comandi scrivi:
```
pip install openpyxl
```

---

## 3. Configurare Gmail — passo passo (5 minuti)

Le email partono dalla tua casella **riccardofreddi6@gmail.com** (già inserita nello script).
Google **non permette** di usare la password normale per inviare email da programmi:
serve una **Password per app** di 16 caratteri. Si crea una volta sola.

### Passo 1 — Attiva la verifica in 2 passaggi
1. Vai su: **https://myaccount.google.com/security**
2. Accedi con riccardofreddi6@gmail.com
3. Cerca la voce **"Verifica in 2 passaggi"** → cliccala → **Attiva**
   (ti chiede il numero di telefono: serve per la sicurezza dell'account)

### Passo 2 — Crea la Password per app
1. Vai su: **https://myaccount.google.com/apppasswords**
   (se ti dice che non è disponibile, ripeti il Passo 1: è obbligatorio)
2. Dove chiede il nome dell'app scrivi: `Promemoria revisioni` → clicca **Crea**
3. Google ti mostra una **password di 16 caratteri**, tipo: `abcd efgh ijkl mnop`
4. **Copiala** (cliccaci sopra, si copia)

### Passo 3 — Mettila nello script
1. Apri la cartella dove hai messo i file
2. Clicca con il tasto destro su `promemoria_revisioni.py` → **Apri con → Blocco note**
3. In cima trovi la sezione CONFIGURAZIONE: alla riga `MITTENTE_PASSWORD` incolla la password tra le virgolette:

```python
MITTENTE_EMAIL    = "riccardofreddi6@gmail.com"   # (già impostata)
MITTENTE_PASSWORD = "abcd efgh ijkl mnop"         # <- incolla qui la tua
NOME_OFFICINA     = "Officina Rossi"              # <- metti il nome del tuo negozio/officina
```

4. **Salva** (File → Salva) e chiudi
5. Rifai doppio click su `controlla_pc.bat`: deve dirti "Gmail configurato"

> Puoi incollare la password con o senza gli spazi: Google li accetta entrambi.
> La password per app NON si vede da nessun'altra parte: se la perdi, ricreala
> dallo stesso sito (puoi crearne una nuova e cancellare la vecchia).

---

## 4. Compilare il file Excel

Apri `clienti_revisioni.xlsx`, un cliente per riga (da riga 5 in giù):

| Colonna | Cosa scrivere |
|---------|---------------|
| **Nome Cliente** | Nome e cognome |
| **Email** | Indirizzo del cliente |
| **Targa** | Es. AB123CD |
| **Scadenza Revisione** | **Come sul libretto**: solo mese e anno, es. `09/2026` |
| **Revisione Fatta** | `SÌ` quando il cliente ha fatto la revisione → **nessuna email più** |
| **Attivo** | `SÌ` = riceve i promemoria, `NO` = escluso (non cancellare la riga) |

### Come funziona la scadenza per mese
Sul libretto c'è solo `09/2026`: per legge la revisione va fatta **entro l'ultimo
giorno del mese**. Quindi tu scrivi solo `09/2026` e lo script calcola da solo
che la scadenza vera è il **30/09/2026**, e da lì calcola i 30/10/1 giorni.

> Se per sbaglio Excel trasforma `09/2026` in una data tipo `01/09/2026`,
> nessun problema: lo script lo riconosce e usa comunque il mese giusto.
> Per sicurezza la colonna è già impostata in formato testo.

### Il ciclo del cliente (importante!)
1. Cliente nuovo → scrivi i dati, Revisione Fatta = **NO**
2. Partono i promemoria (1 mese, 10 giorni, 1 giorno prima)
3. Il cliente fa la revisione → metti Revisione Fatta = **SÌ** → **si ferma tutto**
4. Quando arriva il libretto rinnovato → aggiorna il mese di scadenza e rimetti
   Revisione Fatta = **NO** → i promemoria ripartono per il nuovo ciclo

---

## 5. Provarlo e farlo partire

1. **Doppio click su `controlla_pc.bat`** → tutto deve dire [OK]
2. **Doppio click su `test_promemoria.bat`** → mostra le email che partirebbero
   (nessuna email vera). Se non c'è niente da inviare oggi, te lo dice cliente per cliente.
3. **Doppio click su `avvia_promemoria.bat`** → invio reale

---

## 6. Automatizzare: farlo partire da solo ogni giorno

1. Premi Start e cerca **"Utilità di pianificazione"** (Task Scheduler) → apri
2. A destra: **"Crea attività di base..."**
3. Nome: `Promemoria revisioni` → Avanti
4. **Ogni giorno**, ore **09:00** → Avanti
5. **Avvia un programma** → Sfoglia → scegli `avvia_promemoria.bat` → Avanti → Fine

Consiglio: nei primi giorni schedula `test_promemoria.bat`, così vedi a schermo
cosa farebbe senza mandare email vere. Quando ti fidi, sostituisci con `avvia_promemoria.bat`.
Il computer deve essere acceso all'ora scelta: metti un'orario in cui l'officina è aperta.

---

## 7. Come funziona in breve

- **Scadenza** = ultimo giorno del mese scritto sul libretto (es. `09/2026` → 30/09/2026)
- **Finestre di invio** (per non perdere promemoria se il PC è spento il giorno esatto):
  - "1 mese prima" → quando restano da 26 a 30 giorni
  - "10 giorni prima" → quando restano da 8 a 10 giorni
  - "1 giorno prima" (ultimo avviso) → quando restano 1 o 2 giorni
- Ogni tipo di promemoria parte **una volta sola**: ogni invio viene scritto in
  `reminders_log.txt`, così anche lanciando lo script più volte non arrivano doppioni
- **Revisione Fatta = SÌ** → zero email, anche se la scadenza è vicinissima
- **Attivo = NO** → zero email (cliente disiscritto)
- Revisioni già scadute o che scadono proprio oggi → nessun promemoria (da gestire a mano)
- Se l'invio a un cliente fallisce, lo script continua con gli altri e segnala l'errore

L'email che riceve il cliente (essenziale, 3 righe):

> Da: Officina Rossi \<riccardofreddi6@gmail.com\>
> Oggetto: [Officina Rossi] Promemoria revisione AB123CD – scade tra 30 giorni
>
> La revisione del veicolo targato AB123CD scade entro il 30/11/2026.
> Giorni rimanenti: 30.
> La preghiamo di provvedere per tempo.

---

## 8. Problemi comuni

| Problema | Soluzione |
|----------|-----------|
| Il .bat si chiude subito | Fai doppio click su `controlla_pc.bat` e leggi cosa dice |
| `pip install openpyxl` non funziona | Prova: `py -m pip install openpyxl` |
| Errore `Username and Password not accepted` | Non hai messo la **Password per app** (16 caratteri): rifai il punto 3. Non è la password che usi per leggere la posta! |
| Le email finiscono in spam | Chiedi ai clienti di aggiungere riccardofreddi6@gmail.com ai contatti |
| `non trovo il file Excel` | `clienti_revisioni.xlsx` deve stare nella stessa cartella dello script |
| Voglio cambiare i giorni di invio | Nello script, sezione `FINESTRE_PROMEMORIA`: es. `15: (13, 15)` aggiunge un promemoria a 15 giorni |
| Voglio cambiare il testo dell'email | Nello script, funzione `componi_email` |
| Ho sbagliato e voglio cancellare un invio registrato | Apri `reminders_log.txt` col Blocco note e cancella la riga del cliente |
