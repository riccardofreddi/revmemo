# GUIDA INSTALLATORE / RIVENDITORE (file personale — NON consegnare al cliente)

Questa guida è per chi installa RevMemo presso i centri revisioni.
Al termine dell'installazione **elimina questo file** dal pacchetto del cliente.

---

## 1. Modello commerciale

| Voce | Prezzo consigliato |
|------|--------------------|
| Installazione + licenza **Pro** (una tantum) | **€ 399** |
| Installazione + licenza **Base** (senza HTML/riepilogo, testo semplice) | € 199 |
| Manutenzione annuale (assistenza + modifiche + aggiornamenti) | **€ 99/anno** |
| Add-on SMS (collegamento al gateway a scelta del centro) | + € 100 una tantum + crediti SMS a carico del centro (es. Aruba 100 SMS = € 6,70) |

Argomenti di vendita pronti:
- "Quante revisioni perdete ogni mese per clienti che dimenticano la scadenza?
  Una revisione recuperata (€ 60–90) ripaga il sistema."
- "Nessun canone: si paga una volta. I gestionali chiedono 30–100 €/mese."
- "Zero installazioni: copiamo una cartella, i clienti ricevono le email col VOSTRO logo."
- "Riepilogo giornaliero al titolare: vedete ogni giorno cosa è stato inviato."

---

## 2. Preparare il pacchetto per un nuovo cliente (10 minuti a casa)

1. Copia l'intera cartella `RevMemo` rinominandola es. `RevMemo - Centro Rossi`
2. Apri `config.ini` (se non c'è: lancia una volta `controlla_pc.bat` e lo crea da solo,
   oppure copia `config_ESEMPIO.ini` rinominando la copia in `config.ini`):
   - `[smtp] mittente` → casella Gmail **del cliente** (si crea sul posto, vedi §3)
   - `[centro]` → nome, città, telefono, email, colore del marchio del cliente
   - `[prodotto] nome` → qui decidi TU il nome con cui spacci il prodotto
     (es. `RevMemo`, o un nome tuo — il white label è libero)
3. Sostituisci `logo.png` con quello del cliente (stesso nome file) — se non ce l'ha,
   cancella il file: l'email si adatta da sola
4. Svuota `clienti_revisioni.xlsx` dai dati di esempio (lascia solo la riga 4 delle
   intestazioni e la nota: verranno riconosciute comunque)
5. **Cancella `GUIDA_INSTALLATORE.md`** e la cartella `backup/` se presente
6. Prova in locale: `python revmemo.py --dry-run` (o avvia_prova.bat)

---

## 3. Installazione presso il centro (40–60 minuti)

| # | Passo | Tempo |
|---|-------|-------|
| 1 | Copia la cartella sul PC del centro (desktop o C:\RevMemo) | 2 min |
| 2 | Doppio click su `controlla_pc.bat` → verifica runtime incluso funzioni | 1 min |
| 3 | Crea/usa una casella Gmail del centro (consiglia una casella dedicata es. promemoria.centrorossi@gmail.com); attiva 2FA + Password per app (ISTRUZIONI_CENTRO punto 3) | 10 min |
| 4 | Compila `config.ini` insieme al cliente (dati centro + password app) | 5 min |
| 5 | Import clienti: prendi l'elenco che usano (Excel/cartaceo/gestionale) e trascrivilo in `clienti_revisioni.xlsx` — le colonne importanti sono Email, Targa, Scadenza (mese/anno dal libretto) | 15–20 min |
| 6 | Sostituisci logo.png (foto al logo se serve, ritaglia in PNG ~260x64) | 5 min |
| 7 | `avvia_prova.bat` → mostra a schermo i promemoria che partirebbero; **fai vedere al titolare** il nome del centro nell'oggetto e i casi esclusi | 5 min |
| 8 | Invio di prova a un indirizzo del titolare (metti la sua email in una riga fittizia con scadenza a 28 giorni, invia, poi rimuovi la riga o registra il log) | 5 min |
| 9 | Configura **Utilità di pianificazione Windows**: attività giornaliera 9:00 → `avvia_invio.bat` (chiamalo anche "Promemoria revisioni") | 5 min |
| 10 | Passa `ISTRUZIONI_CENTRO.md`, spiega le colonne Excel e il ciclo "fatta SÌ → rinnovo → rimetti NO" | 10 min |

Fatto questo: fatti firmare/accettare e programma la chiamata di follow-up a 7 giorni.

---

## 4. Note tecniche che ti devono essere chiare

- **Runtime incluso**: la cartella `runtime/` contiene Python 3.12 Windows + openpyxl.
  Sul PC del centro non serve installare nulla. Se per qualche motivo il runtime non
  parte, i file .bat usano automaticamente il Python di sistema (se presente).
- **Limiti Gmail**: 500 email/24h (account normale), 2.000/24h (Workspace). I
  promemoria tipici sono 15–25/giorno: nessun rischio. Consiglia al centro di usare
  la casella anche normalmente (reputazione migliore).
- **Finestre di invio**: 26–30 gg ("30 giorni"), 8–10 gg ("10 giorni"), 1–2 gg ("ultimo
  avviso") — modificabili in config.ini. Il log blocca i doppioni **per tipo e per
  scadenza** (chiave email;targa;tipo;scadenza): se il PC resta spento 3 giorni, al
  rientro parte solo l'ultimo tipo dovuto, e mai due volte lo stesso promemoria.
  Quando il cliente rinnova (nuova scadenza nell'Excel), i promemoria ripartono da soli.
- **Risposte dei clienti**: arrivano nella casella `mittente`; per ridirigerle compila
  `risposta_a` in `[smtp]` (header Reply-To).
- **Backup**: cartella `backup/`, ultimi 14 file clienti. In caso di disastro,
  ripristinare l'ultimo backup e cancellare `reminders_log.txt` (riparte tutto pulito).
- **STOP dei clienti**: quando rispondono STOP, si mette Attivo = NO. Non cancellare
  mai la riga: resta lo storico.
- **SMS add-on**: sezione `[sms]` in config già predisposta (attivo = NO). Quando il
  centro sceglie il gateway (Aruba/Skebby consigliati, crediti a loro nome), si collega
  l'invio — vendilo come extra a € 100.
- **Aggiornamenti del prodotto**: il pacchetto scaricato NON contiene `config.ini`
  (solo `config_ESEMPIO.ini` come modello) NE' `clienti_revisioni.xlsx` (solo
  `clienti_ESEMPIO.xlsx`): puoi ricoprire tutta la cartella sopra quella del cliente
  SENZA perdere né impostazioni né anagrafiche — i file mancanti si ricreano da soli
  al primo avvio. Non toccare comunque mai `clienti_revisioni.xlsx`, `config.ini` e
  `reminders_log.txt` che sono del cliente.

---

## 5. Checklist "prima di andare via dal cliente"

- [ ] `controlla_pc.bat` tutto [OK]
- [ ] Prova a vuoto mostrata e capita dal titolare
- [ ] Email di prova reale ricevuta nella casella del titolare (controlla logo, colore, telefono)
- [ ] Attività pianificata alle 9:00 attiva (test: eseguila manualmente da Task Scheduler)
- [ ] File clienti compilato con i clienti veri
- [ ] `GUIDA_INSTALLATORE.md` eliminata dal pacchetto
- [ ] Follow-up a 7 giorni segnato in agenda
