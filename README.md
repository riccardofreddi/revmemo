# RevMemo — archivio completo del progetto

Copia di sicurezza di **RevMemo**: il programma per la gestione dei
promemoria di revisione per i centri revisioni. Pannello locale nel
browser, motore di invio email, file Excel dei clienti, tutto offline
sul computer del centro.

## Che cosa c'è dentro

| Cartella | Contenuto |
|---|---|
| `revmemo/` | Il pacchetto completo pronto da installare: pannello (`pannello.py` + `pannello.html`), motore (`revmemo.py`), runtime Windows Python 3.12.8 incorporato con openpyxl 3.1.5, guide (`LEGGIMI.txt`, `ISTRUZIONI_CENTRO.md`, `GUIDA_INSTALLATORE.md`) |
| `dist/` | I file ZIP consegnati: `RevMemo_v4.zip` = versione attuale (v4.15); `archivio/` = versioni precedenti |
| `scripts/` | Strumenti di sviluppo: suite di test `test_pannello.py` (151 controlli) e `test_correzioni_v413.py` (55 controlli), script per ricostruire lo ZIP (`ricrea_zip_revmemo_v45.py`), verifiche HTML, harness fotografici |
| `scripts/revmemo_build/` | Asset per ricostruire il runtime incorporato (python_embed.zip + wheel openpyxl) |
| `foto/` | Schermate di verifica del pannello, versione per versione |
| `archivio/` | Progetti precedenti (`promemoria-revisioni`, la prima versione senza pannello) |
| `worklog.md` | Diario di sviluppo: ogni versione con richieste, scelte, test e verifiche |

## Versione attuale: v4.15 (2026-10-31)

- **Nessuna richiesta STOP si perde più**: la ricerca delle risposte
  guarda gli ultimi **35 giorni** della casella email (prima 7) — anche
  a computer fermo per settimane, al rientro le richieste vengono
  trovate e registrate PRIMA di inviare qualsiasi email.
- **Nuovo pulsante "CONTROLLA LE RISPOSTE STOP"** nella Home del
  pannello: legge la casella e registra gli STOP nel file clienti
  SENZA mandare niente (motore: `--solo-risposte`). Rispetta il blocco
  anti-doppio-invio: se un altro invio sta girando, aspetta e lo dice.
- `controlla_risposte(in_invio=...)`: il giro di invio si comporta
  esattamente come prima; gli avvisi del pulsante dedicato parlano di
  "riprova" invece che di invii.
- Test: **test_v415.py 28/28 (nuovo)** + 151/151 pannello + 39/39 v4.14
  + 55/55 correzioni + 9/9 orfano + v43 42/42 + v44 52/52 + verifica
  HTML; v4/v41/v42 con le solite 2 verifiche pre-esistenti.

### v4.14 (2026-10-30)

- **Serratura condivisa del file clienti**: pannello e motore (processi
  separati) salvano il file Excel UNO ALLA VOLTA (`excel.lock`, PID del
  proprietario; a un processo vivo non si ruba mai, uno morto si riprende
  subito). Prima il pannello teneva in fila solo i suoi thread: un invio
  automatico poteva sovrapporsi a un salvataggio del pannello.
- **Salvataggio atomico** (`salva_excel_sicuro`): prima il file nuovo su
  un provvisorio, poi lo scambio — chi legge nel frattempo vede il file
  vecchio o quello nuovo, mai uno scritto a metà. PermissionError (Excel
  aperto) invariato, così i messaggi e i test restano gli stessi.
- **Blocco anti-doppio-invio con controllo del vivo**: prima di riprendersi
  un blocco vecchio, il motore legge il PID e verifica che il processo sia
  morto; un invio lento ma vivo non viene MAI rubato (anche oltre i 15
  minuti), un blocco orfano si libera SUBITO (prima: attesa fissa 15 min).
- **Versione unica**: costante `VERSIONE = "4.14"` in `revmemo.py`, usata
  nell'intestazione e nel banner di avvio (l'intestazione diceva ancora
  "versione 4.4 (PRO)"); il pannello la legge dal motore.
- **Messaggi più chiari**: file clienti occupato → "riprova tra qualche
  secondo" (pannello) / avvisi dedicati per STOP e colora Excel.
- Test: **test_v414.py 39/39 (nuovo)** + **151/151 pannello + 55/55
  correzioni + 9/9 blocco orfano + v43 42/42 + v44 52/52**; v4/v41/v42 con
  le solite 2 verifiche pre-esistenti (pacchetto di sviluppo con config.ini
  dentro, non è quello dello ZIP)

### v4.13 (2026-10-30)

- **Pannello più sicuro**: chiave di sessione + controllo Host/Origin — una
  pagina web qualsiasi aperta nel browser non può più comandare il pannello
  (cambiare server di posta, leggere la password, cancellare clienti, avviare
  invii). La password non viene più inviata al browser.
- **Niente più doppie email**: blocco anti-doppio-invio con battito di 15
  minuti tra pannello e attività pianificata (`revmemo.lock`); liberato solo
  se orfano (controllo del PID)
- **STOP più solidi**: memoria degli STOP già gestiti (`stop_gestiti.txt`,
  una riattivazione manuale non viene più annullata), citazioni di Outlook
  riconosciute, riga del piede sempre scartata, letture IMAP con BODY.PEEK
  (le risposte dei clienti non risultano più "già lette")
- **Invii più robusti**: password sbagliata = fermata immediata con
  spiegazione; oggetto con a-capo salta solo quel cliente; pausa "1,5";
  finestra "30-26"; cella Attivo VUOTA = cliente attivo
- **Backup**: ultimi 20 + ultimo di ogni giorno degli ultimi 30
- **Email**: intestazioni Date/Message-ID/List-Unsubscribe (bottone
  "Annulla iscrizione" in Gmail → STOP); riepilogo titolare con escape HTML
- **Messaggi chiari** per config.ini rotto (prima appariva solo "1")
- `controlla_pannello.py` aggiornato a `/api/ping` (il controllo "pannello
  già aperto" funziona con la nuova sicurezza)
- Test: **151/151 pannello + 55/55 correzioni + v43 42/42 + v44 52/52**;
  v4/v41/v42 con le solite 2 verifiche pre-esistenti (pacchetto di sviluppo
  con config.ini dentro, non è quello dello ZIP)

### v4.12 (2026-10-29)

- Salvando una **scadenza nuova**, la voce "Revisione fatta" torna **NO
  da sola** (avviso nel pannello) — niente più promemoria bloccati per errore
- Pulsante **🗑 Elimina** cliente accanto a Modifica: conferma interna
  nel pannello (nessun browser), backup automatico prima di cancellare,
  le righe sotto si riaggregano da sole
- Aiuto password in Impostazioni: **Libero/Virgilio = password normale
  della casella**, Gmail = Password per app di 16 caratteri

Nota: le correzioni v4.13 nascono da una revisione esterna (patch
`archivio/patch/revmemo_v4.13.patch`), verificata riga per riga, integrata
e completata con correzioni proprie prima della consegna.

Storico completo delle versioni in `worklog.md`.

## Come si ricostruisce il pacchetto ZIP

```
python3 scripts/ricrea_zip_revmemo_v45.py
```

Lo script produce lo ZIP con 258 file (runtime incluso, dati clienti
esclusi). Prima di ricompilare, verificare che non ci siano cartelle
`__pycache__` nel pacchetto.

## Dati dei clienti

I file Excel con i clienti VERI **non sono in questo archivio**: stanno
solo sul computer del centro, con i backup automatici nella cartella
`backup`. Qui ci sono solo i dati d'esempio (Mario Rossi, Luigi Verdi e
compagnia). Anche la configurazione (`config.ini`) è il modello vuoto,
senza password.
