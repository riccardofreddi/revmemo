# RevMemo — ISTRUZIONI per il Centro Revisioni

Promemoria automatici di scadenza revisione via email, con il logo e i colori
del centro. Nessuna installazione necessaria: serve solo Internet.

---

## 1. I file in breve

| File | A cosa serve |
|------|--------------|
| `avvia_invio.bat` | L'esecuzione quotidiana: manda i promemoria del giorno |
| `avvia_prova.bat` | Prova a vuoto: mostra cosa farebbe, senza inviare nulla |
| `controlla_pc.bat` | Controllo automatico: tutto a posto? te lo dice |
| `clienti_revisioni.xlsx` | I tuoi clienti: l'unico file che modifichi di continuo. **Se manca, il programma lo crea da solo** copiando l'esempio |
| `clienti_ESEMPIO.xlsx` | Il modello con i clienti di prova: non lo toccare |
| `config.ini` | La configurazione (si compila una volta). **Se manca, si crea da solo** al primo avvio di controlla_pc.bat o avvia_prova.bat |
| `config_ESEMPIO.ini` | Solo un modello di config.ini: non lo usa il programma. Se vuoi, copialo e rinominalo `config.ini` |
| `template_email.html` + `logo.png` | Il look delle email (si personalizza una volta) |
| `reminders_log.txt` | Registro di ogni invio (anti-doppioni): si crea da solo. Cancellalo SOLO se vuoi re-inviare le stesse email |
| `backup/` | Copie di sicurezza automatiche del file clienti |

---

## 2. Compilare la configurazione (una sola volta)

Apri `config.ini` con Blocco note e compila la sezione `[smtp]`
(non c'è ancora? Lancia una volta `controlla_pc.bat` e lo crea da solo,
oppure copia `config_ESEMPIO.ini` e rinomina la copia in `config.ini`):

```ini
[smtp]
host = smtp.gmail.com
porta = 465
mittente = nomecentro@gmail.com          <- la casella del centro
password_app = aaaa bbbb cccc dddd       <- la Password per app (vedi punto 3)
```

E la sezione `[centro]` — **questi dati appaiono nelle email dei clienti**:

```ini
[centro]
nome = Centro Revisioni Rossi
citta = Milano
telefono = 02 1234567
email_contatti = info@centrorossi.it
colore_tema = #C8102E                    <- il colore del vostro marchio
```

Salva e chiudi. Fatto.

---

## 3. Creare la Password per app di Google (5 minuti, una sola volta)

Google non accetta la password normale per inviare email da programmi.

1. Su **https://myaccount.google.com/security** attiva la **Verifica in 2 passaggi**
2. Su **https://myaccount.google.com/apppasswords** crea una password per app
   (nome: `RevMemo`) → Google mostra 16 caratteri
3. Incollali in `config.ini` alla riga `password_app =` (con o senza spazi: è uguale, il programma li toglie da solo)
4. Salva

> La password per app si può revocare in qualsiasi momento dalla stessa pagina:
> è un accesso dedicato solo a questo programma, non è la vostra password generale.

---

## 4. Compilare l'elenco clienti (`clienti_revisioni.xlsx`)

Un cliente per riga, **una riga per veicolo** (stesso cliente con 2 macchine = 2 righe):

| Colonna | Cosa scrivere |
|---------|---------------|
| Nome Cliente | Nome e cognome (o ragione sociale) |
| Email | Dove arriva il promemoria |
| Telefono | Per gli SMS futuri (facoltativo) |
| Targa | Es. AB123CD |
| Scadenza Revisione | **Come sul libretto**: mese/anno, es. `09/2026` |
| Revisione Fatta | `SÌ` quando il cliente ha fatto la revisione → si fermano le email |
| Attivo | `SÌ` = riceve promemoria; `NO` = escluso |

**Regole d'oro:**
- Non rinominare le intestazioni della riga 4
- Quando il cliente rinnova: aggiorna il mese di scadenza e rimetti `Revisione Fatta = NO`
- Quando un cliente risponde **STOP** a un'email: metti `Attivo = NO`
- Scrivi la scadenza solo mese e anno: il programma calcola da solo che `09/2026`
  vale fino al 30/09/2026 (ultimo giorno del mese, come da legge)

---

## 5. Il ciclo quotidiano

1. **Tutti i giorni** (o quando vuoi): doppio click su `avvia_invio.bat`
   — oppure in automatico ogni mattina con **Utilità di pianificazione di Windows**:
   crea un'attività che alle 9:00 avvia `avvia_invio.bat`
2. Il titolare riceve il **riepilogo email** di ciò che è partito (si può disattivare in config)
3. Una volta a settimana: doppio click su `avvia_prova.bat` per un controllo a vuoto

### Cosa succede automaticamente
- Email professionali in HTML col vostro logo a **30 giorni, 10 giorni e 1 giorno**
  dalla scadenza (con finestre di tolleranza: 26–30 gg, 8–10 gg, 1–2 gg — se il PC
  è spento un giorno, il promemoria non si perde)
- **Zero doppioni**: ogni promemoria parte una volta sola per ogni scadenza
  (quando il cliente rinnova la revisione e aggiorni la scadenza nell'Excel,
  i promemoria ripartono da soli per il nuovo ciclo)
- **Risposte dei clienti**: quando un cliente risponde a un promemoria, la
  risposta arriva nella casella da cui parte la posta (sezione `[smtp]` di
  config.ini). Se preferisci riceverle su un'altra casella, compila
  `risposta_a = ...` sempre nella sezione `[smtp]`
- Revisioni già fatte, clienti disiscritti e revisioni scadute non generano email
- Backup automatico del file clienti prima di ogni esecuzione (ultimi 14 salvati)

---

## 6. Personalizzare il look (opzionale, una volta)

- **Logo**: sostituisci il file `logo.png` col logo del centro (stesso nome file)
- **Colore**: in `config.ini`, `colore_tema` (es. `#1F3B66` blu, `#C8102E` rosso, `#1B7D46` verde)
- **Testo dell'email**: modificando `template_email.html` (o chiedere all'installatore)

---

## 7. Problemi comuni

| Problema | Soluzione |
|----------|-----------|
| La finestra si chiude subito | Doppio click su `controlla_pc.bat` e leggi cosa segnala |
| `Username and Password not accepted` | In config.ini non c'è la **Password per app** (16 caratteri): rifai il punto 3 |
| Le email finiscono in spam | Fatevi aggiungere come contatto dai clienti; succede poco perché le email sono normali e sobrie |
| "Mittente non impostato" | Compila `mittente` in config.ini |
| Voglio cambiare i giorni dei promemoria | In config.ini sezione `[promemoria]` (es. `finestra_1mese = 24-30`) |
| Ho sbagliato un invio | Apri `reminders_log.txt`, cancella la riga del cliente, al prossimo avvio riparte il promemoria |
| Voglio re-inviare tutto in prova | Cancella `reminders_log.txt` (si ricrea da solo): i promemoria in finestra ripartono |
| Non vedo il logo nelle email | Controlla che `logo.png` non sia vuoto (tasto destro → Proprietà): se è danneggiato, ricopialo dal pacchetto |
| Email di un cliente sbagliata | Correggi l'email nell'Excel; i vecchi invii del giorno restano validi (nessun doppio invio) |
