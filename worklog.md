# Worklog

---
Task ID: 1
Agent: main (Z.ai)
Task: Pacchetto "Promemoria revisioni automatici" — Excel clienti + script Python invio email (SMTP Gmail) a 30/10/1 giorni dalla scadenza, con log anti-doppioni e dry-run.

Work Log:
- Caricato skill xlsx; letti SKILL.md, scenes/create.md, engines/design.md, templates/base.py, quality/pipeline.md
- Chiarimenti utente via AskUserQuestion: Gmail SMTP, colonne base (Nome/Email/Targa/Scadenza/Attivo), email minimale 3 righe, log su file reminders_log.txt, dry-run, modalità esecuzione da decidere (fornite entrambe: manuale .bat + Task Scheduler)
- Creato /home/z/my-project/scripts/genera_excel_clienti.py → genera clienti_revisioni.xlsx con design system skill (origine B2, header primario, righe alternate, dropdown SÌ/NO su F5:F300, formattazione condizionale verde/rosso su Attivo, date DD/MM/YYYY, freeze panes B5); 8 clienti di esempio con date relative a oggi (30/10/1 gg, non attivi, scaduta)
- QA Excel: recalc/audit/scan/validate → tutti 0 errori, exit 0
- Creato promemoria_revisioni.py: lettura Excel per nome-intestazione (colonne riordinabili), parsing date robusto (datetime/date/3 formati stringa), trigger 30/10/1 giorni, log anti-doppioni (data;email;targa;tipo), SMTP_SSL Gmail 465 con EmailMessage UTF-8, --dry-run, --file, percorsi assoluti da __file__ (compatibile Task Scheduler), riepilogo finale con contatori
- Creati test_promemoria.bat (dry-run) e avvia_promemoria.bat (invio reale), convertiti a CRLF
- Creato ISTRUZIONI.md: requisiti, Password per app Google, compilazione Excel, dry-run, schedulazione Task Scheduler passo-passo, regole funzionamento, tabella problemi comuni
- Test dry-run: rilevata riga-nota interpretata come cliente → fix skip (richiesta almeno una tra email/targa/scadenza); re-test OK (8 clienti, 3 invii previsti, 5 esclusi)
- Test log: rilevato bug in gia_inviato (match esatto vs prefisso) → fix startswith; re-test OK (stesso tipo=True, tipo diverso=False, altra email=False); log di test rimosso

Stage Summary:
- Deliverable in /home/z/my-project/download/promemoria-revisioni/: clienti_revisioni.xlsx, promemoria_revisioni.py, test_promemoria.bat, avvia_promemoria.bat, ISTRUZIONI.md
- Script Excel generatore persistito in /home/z/my-project/scripts/genera_excel_clienti.py per rigenerazioni future
- Nessun artefatto di test residuo nel pacchetto
- Utente deve solo: inserire email + Password per app Google nello script, sostituire clienti di esempio, scegliere esecuzione manuale (.bat) o Task Scheduler

---
Task ID: 2
Agent: main (Z.ai)
Task: Versione 2 su richiesta utente — colonna "Revisione Fatta" (SÌ = stop totale email), scadenza mese/anno tipo libretto (09/2026 = ultimo giorno del mese), supporto configurazione Gmail per riccardofreddi6@gmail.com.

Work Log:
- Rigenerato clienti_revisioni.xlsx: nuova colonna "Revisione Fatta" (dropdown SÌ/NO, CF verde/rosso), colonna scadenza in formato testo MM/YYYY con esempi mese-corrente/-1/+1/+2/+11, nota d'uso aggiornata; QA recalc/audit/scan/validate → 0 errori
- Riscritto promemoria_revisioni.py (v2): parsing scadenza MM/YYYY → ultimo giorno del mese (calendar.monthrange), gestione auto-conversione Excel (datetime con giorno=1 → mese libretto), finestre di tolleranza {30:[26-30], 10:[8-10], 1:[1-2]} per non perdere invii se PC spento, blocco totale se Revisione Fatta=SÌ, log anti-doppioni per TIPO (30/10/1) con chiave data;email;targa;tipo, MITTENTE_EMAIL=riccardofreddi6@gmail.com, From con nome officina, guard password vuota (spazi ammessi nella app password)
- Creato controlla_pc.py + controlla_pc.bat: diagnostica Python/openpyxl/Excel/config Gmail con prossimo passo suggerito; .bat aggiornati con rilevamento automatico python/py (where python)
- Test: dry-run file reale (mesi libretto, 8 clienti: scaduta, fatta-bloccata, non attivi) OK; file test temporaneo _test_trigger.xlsx con date dentro/fuori finestre → trigger 30/9/2 gg corretti, D(28gg, fatta=SÌ) bloccato, E(25gg)/F(7gg)/G(0gg) esclusi; test gia_inviato per tipo OK; file test cancellato
- Riscritto ISTRUZIONI.md: sezione Gmail passo-passo (2FA + app password con URL esatti), ciclo cliente (fatta SÌ → aggiorna mese → rimetti NO), dove sta Python (where python), finestre di invio, nuovi problemi comuni

Stage Summary:
- Pacchetto v2 completo in /home/z/my-project/download/promemoria-revisioni/ (7 file)
- Regola utente implementata e verificata: Revisione Fatta=SÌ → nessuna email nemmeno a 28 giorni dalla scadenza
- Scadenze gestite a mesi come sul libretto; finestre di tolleranza documentate
- Utente deve ancora: creare Password per app Google e incollarla nello script (ISTRUZIONI.md punto 3)

---
Task ID: 3
Agent: main (Z.ai)
Task: RevMemo v3 PRO — pacchetto vendibile white-label a centri revisioni: runtime zero-installazione, email HTML con logo, config.ini esterno, riepilogo titolare, retry, backup, guide rivenditore.

Work Log:
- Verificati limiti Gmail via web (500 email/24h consumer, 2000 Workspace, account nuovi ~100/giorno, 0.3% soglia reclami spam) e prezzi SMS Italia (Aruba 100 SMS 6,70 euro, BulkGate ~0,017 euro/SMS)
- Deciso modello di distribuzione: cartella portatile con Python 3.12.8 embeddable Windows + openpyxl 3.1.5/et_xmlfile 2.0.0 estratti in Lib/site-packages, python312._pth configurato (python312.zip, ., Lib/site-packages); SSL DLLs presenti (libssl-3, _ssl.pyd, libcrypto-3); fallback nei .bat su python/py di sistema
- Scritto revmemo.py v3 (~500 righe): config.ini esterno (configparser utf-8, finestre configurabili "26-30", pausa invii, notifica titolare), Excel con colonna Telefono facoltativa, email multipart testo+HTML con template_email.html (segnaposto {NOME_CENTRO}/{COLORE}/{TARGA}/{GIORNI} ecc. + logo via add_related cid), classe Sender (1 connessione SMTP persistente, retry 3 tentativi con backoff, riconnessione), backup automatico Excel (rotazione 14), riepilogo giornaliero HTML al titolare (solo se invii o errori), console ASCII-safe (reconfigure utf-8) per Task Scheduler, sezione [sms] predisposta
- Creati: config.ini compilabile, template_email.html (table-based compatibile client email, footer STOP GDPR-friendly), logo.png placeholder PIL (RevMemo bianco su trasparente), 3 .bat con rilevamento runtime->python->py, LEGGIMI.txt, controlla_pc.py v3 (verifica anche intestazioni Excel riga 4 e config)
- Excel rigenerato con colonna Telefono (7 colonne B-H), dropdown/CF spostati su G/H, nota multi-veicolo; QA validate → 0 problemi
- Test: controlla_pc OK (2 sole cose attese da compilare); dry-run 8 clienti corretto; rendering HTML senza segnaposto residui e con logo; test trigger con file temporaneo (30/9/2 gg corretti, fatta=SÌ bloccata, digest titolare stampato); test retry SMTP su host rifiutato → 3 tentativi 9s, errore gestito, sessione pulita
- Scritti ISTRUZIONI_CENTRO.md (fine user) e GUIDA_INSTALLATORE.md (white-label, listino 399/199/99+SMS 100, checklist installazione 10 passi, note tecniche)
- .bat/LEGGIMI convertiti CRLF; creato RevMemo_v3.zip (11 MB) in download/

Stage Summary:
- /home/z/my-project/download/RevMemo/ = prodotto vendibile completo; RevMemo_v3.zip pronto da trasferire su Windows
- White label totale: nessun nome azienda; nome prodotto in config.ini [prodotto]; branding email = del centro cliente
- Da verificare su Windows reale (non testabile da sandbox): avvio runtime\python.exe e invio SMTP reale; fallback a Python di sistema previsto
- Prossimi passi possibili: presentazione commerciale PDF, collegamento SMS gateway, versione SaaS su VPS quando >5 centri

---
Task ID: 4
Agent: main (Z.ai)
Task: Diagnosi e correzione dei 3 problemi segnalati dal controllo sul PC utente (runtime assente, config.ini mancante, openpyxl rotta su Python 3.14 di sistema con circular import "decimal").

Work Log:
- Verificato RevMemo_v3.zip: COMPLETO (config.ini, runtime/, openpyxl inclusi) -> sul PC utente la copia era parziale (mancavano cartella runtime e config.ini)
- Analizzato controlla_pc.py originale: test openpyxl catturava solo ImportError -> l'errore "cannot import name 'Decimal' from partially initialized module 'decimal'" (openpyxl presente ma rotta su Python 3.14.2 di sistema, nuovo install manager) veniva mostrato come "Manca openpyxl" -> diagnosi fuorviante
- Confermata catena causa: openpyxl/compat/__init__.py -> .numbers -> from decimal import Decimal; openpyxl 3.1.5 non affidabile su 3.14 -> la strada giusta e' il runtime incluso (Python 3.12 + openpyxl testato)
- Riscritto controlla_pc.py v2 auto-riparante: (a) relaunch automatico su runtime\python.exe via os.execv con guard env REVMEMO_NEL_RUNTIME, (b) diagnosi openpyxl differenziata (manca vs presente-ma-rotta), (c) auto-creazione config.ini default se assente, (d) rilevamento file .py con nomi riservati (decimal.py, numbers.py...) causa circular import, (e) blocco finale "COSA FARE ADESSO"
- Patch revmemo.py: relaunch runtime automatico PRIMA dell'import openpyxl (rimosso vecchio import pre-relaunch con MultiEdit), messaggi errore openpyxl differenziati, carica_config auto-crea config.ini da CONFIG_DEFAULT invece di uscire
- Verificato python312._pth: include Lib/site-packages -> openpyxl nel runtime importabile
- LEGGIMI.txt: aggiunte "REGOLA D'ORO" (copiare cartella intera) e "SE QUALCOSA NON VA" (runtime mancante, openpyxl 3.14, config auto-creato, Decimal)
- Test in sandbox: (1) controllo su copia senza runtime/config -> config auto-creato, runtime segnalato con istruzioni; (2) revmemo.py --dry-run senza config -> auto-creazione + dry-run 8 clienti corretto; (3) diff config auto-creato vs originale -> IDENTICI; py_compile OK
- Ricreato RevMemo_v3.zip (254 file, 10.9 MB) con zip verificato per contenuto chiave

Stage Summary:
- Il pacchetto e' ora auto-riparante: config.ini si ricrea da solo; se c'e' la cartella runtime, TUTTO gira sul Python 3.12 incluso (il Python 3.14 rotto del PC non conta piu')
- Per l'utente: basta copiare/estrarre la cartella runtime dallo zip (o ricopiare il pacchetto intero), rifare il controllo, compilare config.ini
- Rimane da fare lato utente: App Password Google in config.ini [smtp]

---
Task ID: 5
Agent: main (Z.ai)
Task: Errore utente in dry-run: PermissionError [Errno 13] su clienti_revisioni.xlsx (file aperto in Excel durante la lettura).

Work Log:
- Diagnosi: PermissionError su load_workbook = file Excel aperto in Excel / bloccato da OneDrive (cartella utente in OneDrive\Desktop); NON un problema di contenuto modificato
- Patch revmemo.py leggi_clienti: try/except PermissionError -> messaggio amichevole in italiano (chiudi Excel e riprova; se continua riavvia il PC) + except zipfile.BadZipFile -> messaggio file non valido (aprilo con Excel e risalva); aggiunto import zipfile
- Patch controlla_pc.py: dettaglio amichevole su "Permission denied / Errno 13" nel test lettura Excel
- py_compile OK; ricreato RevMemo_v3.zip (254 file, 10.9 MB, verificato)

Stage Summary:
- Il programma ora spiega l'errore "file aperto" in italiano invece del traceback
- Utente guidato: salvare/chiudere Excel e rilanciare avvia_prova.bat

---
Task ID: 6
Agent: main (Z.ai)
Task: Utente ha perso le chiavi ricopiando lo zip sopra la cartella (config.ini default ha sovrascritto quello compilato). Rendere il pacchetto a prova di perdita impostazioni.

Work Log:
- Causa: lo zip conteneva config.ini con valori di partenza -> copiato sopra l'installazione, ha azzerato mittente/password_app
- Rinominato config.ini -> config_ESEMPIO.ini nel pacchetto (modello non usato dal programma); zip ricreato SENZA config.ini (verifica automatica nello script)
- Conseguenza: aggiornamenti/riscoperta del pacchetto non cancellano piu' le impostazioni; il config si ricrea da solo al primo avvio (logica gia' presente dal Task 4)
- Patch revmemo.py: password_app ripulita da TUTTI gli spazi ("abcd efgh ijkl mnop" -> "abcdefghijklmnop...") -> la guida "con o senza spazi e' uguale" ora e' vera
- Aggiornate ISTRUZIONI_CENTRO.md (tabella file + sezione 2 + punto 3), GUIDA_INSTALLATORE.md (par.2 e nota aggiornamenti), LEGGIMI.txt
- py_compile OK; zip 254 file 10.9 MB verificato (config_ESEMPIO presente, config.ini assente)

Stage Summary:
- Pacchetto a prova di perdita: config.ini mai distribuito, si auto-crea al primo avvio
- Utente guidato a ricompilare mittente+password_app; se la chiave non se la ricorda, Google non la ridisplay: crearne una nuova su myaccount.google.com/apppasswords

---
Task ID: 7
Agent: main (Z.ai)
Task: Utente conferma invio funzionante ma segnala 3 problemi: (1) vuole anche la risposta del cliente, (2) logo non visibile nell'email, (3) re-invio bloccato + domanda "posso cancellare reminders? si ricrea da solo?".

Work Log:
- TROVATO BUG REALE DEL LOGO: email.contentmanager scrive Content-ID SENZA parentesi angolari ("Content-ID: logo") quando si passa cid="logo"; per RFC 2392 src="cid:logo" richiede "Content-ID: <logo>" -> molti client (Gmail incluso) non mostrano l'immagine. Fix: add_related(..., cid="<logo>") in Sender.invia; verificato su raw bytes
- Logo segnaposto era bianco su trasparente (invisibile su fondo chiaro): rigenerato badge bianco arrotondato con scritta blu tema + bordo, visibile su qualsiasi sfondo (scripts/nuovo_logo_revmemo.py, anteprime su blu/bianco verificate); template: background-color esplicito anche sulle celle interne dell'intestazione; altezza logo 34->38 e alt dinamico col nome centro
- leggi_logo riscritta con cache: avvisa in console se logo.png e' vuoto/illeggibile (es. OneDrive) invece di mandare email mute
- LOG ANTI-DOPPIONI: scoperto che la chiave includeva la data del giorno (bloccava solo nello stesso giorno -> email ripetute nei giorni successivi con PC schedulato). Nuova semantica per-ciclo: chiave email;targa;tipo;scadenza (riga log: data;email;targa;tipo;scadenza;descrizione); compatibilita' con vecchio formato (righe senza scadenza valgono come gia' inviato, zero doppioni); al rinnovo della revisione (nuova scadenza) i promemoria ripartono da soli
- RISPOSTE CLIENTI: arrivano per natura nella casella mittente; aggiunta opzione risposta_a (Reply-To) in [smtp] (CONFIG_DEFAULT revmemo.py + controlla_pc.py + config_ESEMPIO.ini sincronizzati)
- PACCHETTO A PROVA DI AGGIORNAMENTO COMPLETO: rinominato clienti_revisioni.xlsx -> clienti_ESEMPIO.xlsx nel pacchetto; nuova funzione crea_excel_se_manca in revmemo.py e auto-copia in controlla_pc.py: il file clienti mancante si ricrea dall'esempio -> lo zip estratto sopra l'installazione non tocca NE' config NE' anagrafiche
- controlla_pc.py: nuova riga di stato del registro anti-doppioni (presente con N invii / si crea da solo; suggerisce di cancellarlo per re-inviare in prova) e controllo logo.png vuoto/danneggiato
- Guide aggiornate: LEGGIMI.txt (file principali, re-invio, sezione "LE RISPOSTE DEI CLIENTI", aggiornamenti), ISTRUZIONI_CENTRO.md (tabella file, zero doppioni per scadenza, risposte, 2 nuove righe problemi comuni), GUIDA_INSTALLATORE.md (finestre/chiave log, risposta_a, aggiornamenti sicuri)
- Test scripts/test_revmemo_v31.py: 27/27 OK (py_compile, struttura MIME multipart/related con image/png Content-ID <logo>, dedupe per scadenza + legacy + case-insensitive + registra_invio, auto-copia Excel da controlla_pc, dry-run con cliente in finestra 8-10 -> PROVA invio con logo SI, risposta_a default)
- Ricreato RevMemo_v3.zip (252 file, 10.9 MB) con verifica: config.ini, clienti_revisioni.xlsx, reminders_log.txt ASSENTI dal pacchetto; __pycache__ escluso

Stage Summary:
- Logo risolto alla radice (bug RFC 2392 + placeholder invisibile): nella nuova estrazione l'immagine arriva allegata correttamente
- Anti-doppioni ora davvero "una volta sola per scadenza"; per re-inviare in prova: cancellare reminders_log.txt (si ricrea da solo)
- Il pacchetto e' diventato interamente update-proof: estrazione sopra l'installazione non puo' piu' distruggere ne' chiavi ne' clienti
- Prossimo possibile add-on commerciale: lettura automatica risposte/STOP via IMAP (Attivo=NO automatico)

---
Task ID: 9
Agent: main (Z.ai)
Task: Ok dell'utente al piano v4: "cancella tutto e ridammi tutto pulito, verifica che tutto funzioni, sii critico". Implementare STOP automatico, pulizia file, test completi, zip nuovo.

Work Log:
- revmemo.py v4 (sostituzioni chirurgiche, py_compile OK): (1) nuova sezione 3-bis "RISPOSTE DEI CLIENTI: STOP AUTOMATICO" — IMAP SSL porta 993 sulla stessa casella mittente con la stessa password_app, ricerca email degli ultimi 7 giorni (mesi in INGLESE per evitare il bug di strftime con locale italiano), max 60 email, pre-filtro mittente vs email presenti nell'Excel, parola STOP con regex \bstop\b case-insensitive (stop/Stop/STOP, anche in frase; "stoppare" non conta), corpo text/plain o HTML privo di tag; al match: Attivo=STOP su TUTTE le righe dello stesso indirizzo, salvataggio, email marcata \Seen SOLO dopo salvataggio riuscito (retry al prossimo avvio), evento nel log reminders_log.txt; (2) refactor individua_colonne condiviso; (3) crea_excel_se_manca riscritta: Excel esempio generato DA CODICE (titolo, nota, 7 colonne, 8 esempi con mese-libretto ~±6 gg, dropdown SÌ/NO G5:H300, CF verde/rosso, freeze A5) — niente piu' clienti_ESEMPIO.xlsx; (4) attivo_testo -> messaggio "ha chiesto STOP -> nessuna email"; (5) riepilogo titolare: contatore STOP, tabella "Clienti che hanno risposto STOP", inviato anche se SOLO STOP; (6) riferimento guida -> LEGGIMI.txt sezione 5; (7) commento config su risposta_a che uccide lo STOP automatico
- BUG REALE TROVATO DAI TEST: _primo_blocco ritornava il prefisso imaplib (b'1 (RFC822 {123}') invece del contenuto del messaggio -> con Gmail vera NESSUN STOP sarebbe stato letto. Fix: scartare i byte che matchano prefisso \{\d+\}$ / ^\d+ \(, prendere il payload
- controlla_pc.py v3: niente piu' copia da clienti_ESEMPIO (Excel mancante = [i] "si crea da solo al primo avvio"), CONFIG_DEFAULT sincronizzato, riferimenti LEGGIMI.txt sezione 5, zip generico
- LEGGIMI.txt riscritto come UNICA guida (10 sezioni: cosa fa, file, prima volta, Excel, password per app, STOP automatico, look, uso quotidiano, problemi comuni, installare per altro centro) in CRLF
- Eliminati dal pacchetto: clienti_ESEMPIO.xlsx, config_ESEMPIO.ini, GUIDA_INSTALLATORE.md, ISTRUZIONI_CENTRO.md; eliminati anche il vecchio pacchetto v1 download/promemoria-revisioni/ e RevMemo_v3.zip
- GUIDA_INSTALLATORE.md ricreata FUORI dal pacchetto (download/GUIDA_INSTALLATORE.md, personale) aggiornata: STOP automatico, test STOP davanti al cliente, risposta_a NON compilare, aggiornamenti via revmemo.py, checklist senza voce file da eliminare
- Test scripts/test_revmemo_v4.py: 53/53 OK — parola STOP (9 casi), host imap, Excel da codice (8 clienti, SÌ/NO/Fatta), segna_stop (riga 5 = STOP, log, idempotente, sconosciuto ignorato), estrai testo/HTML, controlla_risposte con FAKE IMAP (dry-run non tocca nulla; reale scrive STOP riga 6, marca \Seen solo lo STOP, ignora esterni; rilancio = zero eventi), main --dry-run completo, controlla_pc subprocess OK, pacchetto reale resta senza file generati
- scripts/crea_zip_revmemo.py: RevMemo_v4.zip 248 file 10.9 MB, verificato (file chiave presenti; zero ESEMPIO/config.ini/clienti/log/__pycache__)

Stage Summary:
- Pacchetto finale pulito: RevMemo/ con 9 elementi (LEGGIMI.txt, 3 .bat, controlla_pc.py, logo.png, revmemo.py v4, runtime/, template_email.html) + RevMemo_v4.zip in download/
- La cartella dell'utente dopo estrazione mostra SOLO file che servono davvero; config.ini, clienti_revisioni.xlsx e reminders_log.txt si creano da soli al primo avvio
- Utente deve: cancellare vecchia cartella, estrarre RevMemo_v4.zip, controlla_pc.bat, compilare config.ini (mittente+password_app+centro, risposta_a VUOTA), avvia_prova.bat
- Da chiedere all'utente (critico): avvio automatico giornaliero (Task Scheduler), conferma automatica allo STOP, avvisi miglioramenti da segnalare

Work Log:
- Riletto revmemo.py completo: confermato che Attivo!=SÌ blocca gli invii (quindi "STOP" nell'Excel funziona subito come blocco); log = reminders_log.txt (riga data;email;targa;tipo;scadenza); logo allegato con cid "<logo>" (fix v3.1 gia' presente); clienti_ESEMPIO.xlsx copiato se manca il file clienti; config.ini auto-creato da CONFIG_DEFAULT
- Letti template_email.html, LEGGIMI.txt, ISTRUZIONI_CENTRO.md per la guida "dove cambiare i dati" (segnaposto: NOME_CENTRO/COLORE/TARGA/SCADENZA/GIORNI/TRAFRA/PRODOTTO/TELEFONO/EMAIL_CENTRO/CITTA/ANNO/LOGO_IMG)
- Piano formulato e spiegato all'utente: STOP via IMAP Gmail (imap.gmail.com 993, stessa casella mittente + stessa password_app, niente di nuovo da configurare) con regex parola intera case-insensitive, match email mittente <-> colonna Email Excel, scrittura "STOP" in Attivo, email marcata come letta, avviso se mittente non trovato; controllo eseguito ad ogni avvio PRIMA degli invii (in prova solo simulazione)
- Pulizia proposta: fuori dal pacchetto clienti_ESEMPIO.xlsx (il programma creera' l'esempio da solo via codice), config_ESEMPIO.ini (ridondante), GUIDA_INSTALLATORE.md (solo per installatore), ISTRUZIONI_CENTRO.md (fusa nel nuovo LEGGIMI.txt unico); risultato: cartella con 8 elementi + runtime
- In attesa di OK utente per implementare

Stage Summary:
- Nessun file ancora modificato: solo piano spiegato all'utente
- Al varo: revmemo.py v4 (controlla_risposte IMAP + crea Excel esempio da codice), nuovo LEGGIMI.txt unico, rimozione 4 file, zip ricreato, istruzioni di aggiornamento PC utente (copiare revmemo.py + LEGGIMI.txt, cancellare i 4 file)

---
Task ID: 10
Agent: main (Z.ai)
Task: Utente non riesce piu' a ricevere email di prova: "se metto la mia email ad un cliente mi fa vedere STOP". Diagnosi + fix v4.1.

Work Log:
- CAUSA RADICE individuata: (1) il footer del promemoria contiene la parola "STOP" ("Rispondi a questa email con oggetto STOP"); (2) l'utente ha messo la PROPRIA email (riccardofreddi6@gmail.com) su una riga cliente; (3) Gmail consegna nella inbox anche le email mandate a se stessi -> il controllo IMAP vedeva un'email "da un cliente" contenente STOP -> scriveva Attivo=STOP nell'Excel (righe 5 e 6: Mario Rossi, Luigi Verdi)
- Fix 1 revmemo.py v4.1: caselle_del_centro() = mittente + email_titolare + email_contatti; in controlla_risposte le email che partono dal centro stesso vengono SEMPRE ignorate (check SPOSTATO PRIMA del match clienti, cosi' conta anche se l'email del centro e' su una riga cliente), con contatore a schermo "(ignorate N email inviate dal centro stesso)"
- Fix 2 revmemo.py v4.1: rimuovi_citazioni() = nelle risposte viene letta solo la parte scritta dal cliente; taglio su ">" iniziale, "ha scritto:", "wrote:", "messaggio originale/inoltrato", "original/forwarded message"; una riga che contiene SOLO "STOP" (anche sotto la citazione, bottom-post) resta valida; il footer citato nel reply "VA BENE" non genera piu' STOP falsi
- Header docstring + commento sezione 3-bis aggiornati a v4.1 con spiegazione delle 2 protezioni
- LEGGIMI.txt (CRLF preservato, patch via scripts/patch_leggimi_v41.py): regola d'oro "NON usare la tua email come email di un cliente", sezione 6 "Protezioni (importanti, evitano STOP falsi)", sezione 9 voce "Ho messo la MIA email su una riga cliente"
- Test scripts/test_revmemo_v41.py: 33/33 OK (riproduce ESATTAMENTE lo scenario utente con Gmail finta: self-email con footer STOP -> ignorata e contata; "VA BENE" citando il footer -> NON stop; "STOP" in cima e bottom-post -> stop; Excel intatto sulle righe non-STOP; rilancio senza doppioni; main dry-run completo)
- Regression test v4: 53/53 OK; py_compile OK; __pycache__ rimosso dal pacchetto
- Ricreato RevMemo_v4.zip (248 file, 10.9 MB) con verifica automatica

Stage Summary:
- v4.1 consegnata: le email del centro stesso non diventano mai STOP; le risposte citate non generano STOP falsi; schermo piu' trasparente
- Per l'utente: (1) rimettere SÌ nella colonna Attivo alle righe di test (ora c'e' STOP); (2) sostituire revmemo.py sul PC; (3) per re-inviare in prova cancellare reminders_log.txt; (4) per provare lo STOP usare un'ALTRA casella email (mai quella del centro)

---
Task ID: 11
Agent: main (Z.ai)
Task: "ho risposto stop da riccardofreddi6@gmail.com e riavviato il programma ma non legge stop" — la v4.1 ignorava TUTTE le email del centro stesso, anche le risposte vere del titolare. Fix v4.2.

Work Log:
- Analisi: comportamento atteso della v4.1 (protezione anti-falso-STOP) ma UX sbagliata: il titolare che prova lo STOP dalla propria casella non riceve alcun effetto ne' feedback
- Fix revmemo.py v4.2: (1) le email del centro stesso NON sono piu' scartate a prescindere: se sono una RISPOSTA VERA (header In-Reply-To/References oppure oggetto che inizia per Re:/R:) o hanno la parola STOP nell'OGGETTO, si prosegue e il corpo decide; (2) le COPIE automatiche (promemoria/riepilogo mandati a se stessi: niente Re:, oggetto "Promemoria..."/"riepilogo") restano ignorate con contatore a schermo; (3) nuovo caso [INFO]: risposta STOP da una casella del centro NON presente nell'Excel -> messaggio chiaro invece del silenzio
- Il check del corpo usa sempre rimuovi_citazioni: lo STOP citato nel footer non conta, solo lo STOP scritto a mano
- patch_leggimi_v42.py (CRLF): sezione 4 regola d'oro aggiornata ("per fare da cliente di prova va bene la TUA email, ma lo STOP va scritto RISPONDENDO a un promemoria"), sezione 6 con i 2 modi per provare lo STOP (a: RISPONDI dalla propria casella; b: altra casella)
- Test scripts/test_revmemo_v42.py: 22/22 OK — scenario esatto utente (reply STOP da mittente con In-Reply-To -> riga 8 STOP), reply VA BENE ignorato, oggetto STOP da secondo indirizzo del centro -> riga 12, copia promemoria e copia riepilogo ignorate (contatore=2), [INFO] per info@test.it non nell'Excel, cliente vero riga 5, rilancio senza doppioni, main dry-run OK
- Regressioni: v4.1 33/33 OK, v4 53/53 OK (totale 108 controlli verdi); py_compile OK; __pycache__ rimosso
- Ricreato RevMemo_v4.zip (248 file, 10.9 MB) con verifica automatica

Stage Summary:
- v4.2: lo STOP funziona anche dalla casella del centro SE scritto come risposta (Re:) o con oggetto STOP; le copie automatiche restano ignorate; [INFO] esplicito quando l'email non e' nell'Excel
- Per l'utente: sostituire revmemo.py, poi RISPONDERE (Re:) al promemoria ricevuto scrivendo STOP e rilanciare avvia_invio.bat

---
Task ID: 12
Agent: main (Z.ai)
Task: "si può fare un excel migliore. tipo si verde e no e stop rosso. in modi che si vede subito la differenza" — Excel con colori immediati (SÌ verde, NO e STOP rosso).

Work Log:
- Analisi: v4.2 colorava gia' SÌ/NO (CellIsRule esatta, con accento) nelle colonne G/H ma: STOP senza colore, "SI" senza accento non riconosciuto, nessuna visibilita' a livello riga, tendina solo SÌ/NO
- revmemo.py v4.3: COLORI_EXCEL + applica_colori(ws, colonna_di, ultima) con 7 regole di FORMATTAZIONE CONDIZIONALE FormulaRule: G/H cella SÌ verde (C6EFCE/006100), NO rosso (FFC7CE/9C0006), STOP rosso scuro C00000 testo bianco grassetto; 2 tinte riga intera B:H secondo $Attivo (verde E2EFDA / rosso FCE4E4); formule OR(UPPER(...)) per SI/SÌ/minuscole; PatternFill start_color+end_color (dxf robusto in Excel)
- crea_excel_se_manca: tendina "SÌ,NO,STOP", colonna Scadenza F5:F300 formato testo @ (01/2026 non diventa piu' gen-26), nota a video sui colori
- Nuova migliora_excel(): per file gia' esistenti senza toccare i dati — riscrive le scadenze trasformate in data da Excel (giorno 1 -> "MM/YYYY", contate e riferite a video), aggiorna/crea tendina con STOP (showErrorMessage=False per permettere STOP manuale), azzera CF vecchie (ConditionalFormattingList) e rimette le 7 regole (rilanciabile senza doppioni), freeze A5 se assente, larghezze solo se mai impostate (lettera non in column_dimensions), errori friendly per PermissionError/file strani
- CLI --colora-excel (sys.exit 0/1) + colora_excel.bat (CRLF, modello avvia_invio.bat)
- prepara_v43_extra.py (CRLF): LEGGIMI sezione 2 (+colora_excel.bat), sezione 4 (paragrafo colori + scadenza testo), sezione 9 FAQ "gen-26"
- Test scripts/test_revmemo_v43.py: 42/42 OK (converti_scadenza con datetime, 7 regole, STOP C00000, tinte, testo @, migliora su file vecchio: dati intatti tranne scadenze convertite, 2 rilanci senza doppioni, leggi_clienti coerente, CLI exit 0)
- Fix durante i test: riga residua fuori dal blocco except di crea (inserimento migliora); larghezze: openpyxl restituisce width=13 di default -> check "lettera not in column_dimensions"
- Regressioni: v4 53/53, v4.1 33/33, v4.2 22/22 (totale 150 verdi); XML xlsx ispezionato (priorita' 1-7, formule con virgole, dxfs 5 con fg+bg); __pycache__ rimosso; RevMemo_v4.zip ricreato (249 file, 10.9 MB, colora_excel.bat incluso)

Stage Summary:
- v4.3: Excel self-explanatory — riga verde = attivo, rossa = NO/STOP, STOP rosso scuro; i colori sono nel file e reagiscono alle modifiche in Excel senza rilanciare il programma; colora_excel.bat porta i colori nel file esistente dell'utente e sistema le scadenze "gen-26"
- Per l'utente: sostituire revmemo.py + copiare colora_excel.bat nella cartella del programma, chiudere Excel e fare doppio click su colora_excel.bat una volta sola

---
Task ID: 13
Agent: main (Z.ai)
Task: correzione v4.3 su richiesta utente: "solo le celle devono essere colorate... non ci deve essere si no e stop su revisione fatta. stop non ci deve essere come opzione da nessuna parte".

Work Log:
- revmemo.py v4.3.1: eliminate le 2 tinte di riga intera (B5:H300) -> SOLO celle G/H colorate (5 regole: G Sì/NO, H Sì/NO/STOP); COLORI_EXCEL senza tinta_*; menu a tendina torna '"SÌ,NO"' sia in crea sia in migliora (STOP mai come opzione, resta solo il COLORE rosso scuro quando lo scrive il programma); migliora_excel ora aggiorna le tendine che coprono fatta O attivo (sistema anche i file toccati dalla v4.3 sbagliata con '"SÌ,NO,STOP"'); showErrorMessage=False (menu solo suggerimento); nota nel foglio e messaggi aggiornati
- patch_leggimi_v431.py (CRLF): paragrafo sezione 4 corretto ("Colori automatici nelle CELLE... STOP non e' nel menu a tendina: lo scrive solo il programma")
- test_revmemo_v43.py aggiornato: 42/42 OK (tendina SOLO SÌ,NO, 5 regole, niente tinte: intervalli solo G5:/H5:, vecchio file con menu sbagliato SÌ,NO,STOP riportato a SÌ,NO, STOP rosso scuro resta come colore cella, dati intatti, doppio lancio 5 regole)
- verifica_xml_v431.py: XML ispezionato — intervalli ['G5:G300','H5:H300'], 5 cfRule, niente B5:H300, menu "SÌ,NO" senza STOP, C00000 negli stili (nota: primo tentativo inline fallito per assert scritto al contrario, rifatto come script)
- Regressioni: v4 53/53, v4.1 33/33, v4.2 22/22 (totale 150 verdi); __pycache__ rimosso; RevMemo_v4.zip ricreato (249 file, 10.9 MB)

Stage Summary:
- v4.3.1: solo le celle di Revisione Fatta e Attivo si colorano (SÌ verde, NO rosso, STOP rosso scuro); STOP NON e' presente in nessun menu a tendina, arriva solo dalle risposte dei clienti; colora_excel.bat aggiorna il file esistente e RIPARA anche i menu toccati dalla v4.3
- Per l'utente: sostituire revmemo.py (il bat non cambia), chiudere Excel e rilanciare colora_excel.bat

---
Task ID: 14
Agent: main (Z.ai)
Task: "devo avere piu' margine di modifica per adattarlo ad ogni centro; ogni centro vuole un messaggio da personalizzare; crea la possibilita' di fare tutto; verifica e spiegami poi ti do ok" — v4.4 PERSONALIZZAZIONE TOTALE.

Work Log:
- Analisi: le PAROLE delle email erano scritte nel codice (componi_email) e nel template HTML: per cambiare una frase bisognava toccare il codice -> i centri chiamerebbero l'assistenza. Progettato il modello "config.ini = LE PAROLE, template_email.html = IL DISEGNO"
- revmemo.py v4.4:
  - CONFIG_DEFAULT con nuova sezione [messaggio]: oggetto, oggetto_1mese/oggetto_10giorni/oggetto_1giorno (vuoto = generale), saluto, testo (MULTIRIGA con riga vuota = paragrafo), firma, riga_extra_piede, testo_disiscrizione; commenti con l'elenco completo dei segnaposto
  - [centro] esteso: indirizzo, sito_web, orari (facoltativi; appaiono nelle email solo se compilati)
  - carica_config: lettura [messaggio] con testi predefiniti (identici alle parole di prima + nome cliente nel saluto), helper _multiriga, ParsingError -> messaggio "le righe dopo testo = devono iniziare con QUALCHE SPAZIO" + exit 1
  - aggiorna_config(): config.ini VECCHI arricchiti da soli (append [messaggio] + inserimento indirizzo/sito_web/orari dentro [centro] prima della sezione successiva), CRLF preservato, idempotente, non tocca nessun valore
  - componi_email(cliente, giorni, tipo): segnaposto estesi ({NOME} con fallback "Cliente", {INDIRIZZO}, {SITO_WEB}, {ORARI}, ...), oggetto speciale per tipo, valori HTML escapati (&, <, >), testo_a_html (**grassetto** -> <b>, riga vuota -> <p>, riga -> <br>), testo semplice senza **, footer con contatti/orari/indirizzo/riga extra/disiscrizione personalizzabile (parola STOP consigliata nel testo)
  - TEMPLATE_DEFAULT_HTML integrato nel .py + aggiorna_template(): template mancante -> creato; template vecchio (senza {TESTO_HTML}) -> aggiornato con backup in template_email_vecchio.html; idempotente
  - main(): aggiorna_template() all'avvio + componi_email(c, giorni, tipo) + dry-run mostra "Messaggio scritto in config.ini sezione [messaggio]"
  - Bug corretto in itinere: togli_grassetto applicato prima della conversione HTML -> grassetto perso; risolto con versione RAW (per HTML) + pulita (per testo)
- template_email.html v2: corpo con {SALUTO_HTML}/{TESTO_HTML}/{FIRMA_HTML}, piede con {INDIRIZZO_FOOTER}/{ORARI_FOOTER}/{RIGA_EXTRA_FOOTER}/{TESTO_DISISCRIZIONE}; la carta GIORNI RIMANENTI resta nel design
- controlla_pc.py: nuovi controlli 6/6-ter (template nuovo/vecchio con auto-aggiornamento, sezione [messaggio] presente/da aggiungere)
- patch_leggimi_v44.py (CRLF): sezione 2 (template = disegno), sezione 3 passo 2, sezione 7 riscritta "IL MESSAGGIO E IL LOOK" (tabella opzioni + segnaposto + regola spazi + aggiornamento automatico), sezione 10 passo 2 per nuovi centri
- test_revmemo_v44.py: 52/52 OK (config vecchio CRLF arricchito senza toccare valori, idempotenza, testi predefiniti, fallback "Cliente", escape &/<, grassetto solo in HTML, oggetti per tipo, footer personalizzato, config rotto -> exit 1 con spiegazione, CLI dry-run in cartella vergine con cliente preparato in finestra 26-30)
- Regressioni: v4 53/53, v4.1 33/33, v4.2 22/22, v4.3 42/53->ricontrollato 42/42 (totale 202 controlli verdi); py_compile OK; __pycache__ rimosso; nota: test CLI v43 dipendeva dai clienti di esempio nella finestra (dipendente dalla data), il nuovo test prepara un cliente a mano
- RevMemo_v4.zip ricreato (249 file, 10.9 MB) con verifica automatica; creato smoke test in /tmp/smoke44 per componi_email

Stage Summary:
- v4.4: ogni centro puo' scrivere DA SOLO le parole delle email (oggetto, oggetti speciali per i 3 promemoria, saluto, testo con paragrafi e grassetto, firma, riga extra, disiscrizione) in config.ini sezione [messaggio], con segnaposto {NOME} {TARGA} {SCADENZA} {GIORNI} {TRAFRA} {NOME_CENTRO} {CITTA} {INDIRIZZO} {TELEFONO} {EMAIL_CENTRO} {SITO_WEB} {ORARI} {PRODOTTO}; nuovi dati centro (indirizzo, sito_web, orari); template vecchio aggiornato da solo con backup; config vecchi arricchiti da soli; white label intatto
- Aggiornamento per i centri esistenti: basta sostituire revmemo.py (e facoltativo template_email.html/controlla_pc.py): config.ini e template si aggiornano da soli al primo avvio
- In attesa dell'OK dell'utente (richiesto: "verifica e spiegami poi ti do ok")

---
Task ID: 15
Agent: main (Z.ai)
Task: Consegnare RevMemo_v4.zip all'utente che non vede il pannello file della chat ("non compare nulla, dammelo qui o con altro metodo").

Work Log:
- Verificata integrità ZIP: 249 file, nessun corrotto, revmemo.py = v4.4 (PRO)
- Esplorato canale preview ufficiale: FC_CUSTOM_LISTEN_PORT=81, gateway su 81 inoltra alla porta 3000 (verificato con server di prova)
- Creato scripts/servi_download.py (server HTTP porta 3000, pagina italiana con bottone download, serve /home/z/my-project/download); attivo in background con log
- Preview URL non trovabile: bot-id assente da env/proc/config; candidati preview-d464.../preview-ws-.../preview-web-cf26... tutti 404; fcapp.run 404; niente doc nello skill fullstack-dev
- Upload su file host esterni (sandbox cn-hongkong): catbox TIMEOUT bloccato; filebin 400 (Varnish); uguu 415 filetype; x0.at 429; transfer.sh/bashupload vuoti; pixeldrain richiede API key; tmpfiles.org OK ma 60 min + interstitial; gofile OK ma pagina non caricabile da sandbox browser (agent-browser wedged, poi "Could not load this folder" anche dopo retry) -> non verificabile
- SUCCESSO: temp.sh -> upload curl -F file=@... https://temp.sh/upload = https://temp.sh/ZuPgP/RevMemo_v4.zip (3 giorni); la pagina usa form POST per il download
- Verifica end-to-end: POST sul link = HTTP 200, 11417008 bytes, zip OK 249 file, sha256 identico all'originale

Stage Summary:
- Link consegnato all'utente: https://temp.sh/ZuPgP/RevMemo_v4.zip (valido 3 giorni, pagina nera con bottone "Click here to download")
- Backup non verificato: https://gofile.io/d/eIlkIdfZ
- Server locale porta 3000 (scripts/servi_download.py) lasciato attivo come canale alternativo
- In attesa che l'utente scarichi; se il link scade si puo' ricaricare con curl -F file=@... https://temp.sh/upload

---
Task ID: 16
Agent: main (Z.ai)
Task: "IL CLIENTE HA VISTO TUTTO MA VUOLE UNA INTERFACCIA GRAFICA... VAI CON IL PANNELLO" + domanda "GitHub+Vercel+Supabase o Cloudflare gratis?" — v4.5 PANNELLO DI CONTROLLO.

Work Log:
- Scelta strada 1 (pannello web locale): motore revmemo.py INVARIATO, pannello = faccia nuova che legge/scrive config.ini e lancia le azioni collaudate
- pannello.py (nuovo, stdlib only, porta 127.0.0.1:8768): pagina HTML + API JSON (stato/messaggio/impostazioni/registro/anteprima/azione/azione-stato/apri-excel/logo); azioni prova/invio/colora/controllo in subprocess con python.exe + CREATE_NO_WINDOW e output live; controllo=subprocess.run(input=INVIO) per l'input() finale di controlla_pc.py; seconda istanza su porta occupata -> apre solo il browser ed esce 0; salva_campi aggiorna config.ini AL POSTO GIUSTO (commenti, CRLF, chiavi sconosciute, blocco multiriga con righe vuote interne gestite con lookahead)
- pannello.html (nuovo): 5 schede (Home/Messaggio/Clienti/Impostazioni/Registro), badge SÌ/NO/STOP colorati, anteprima VERA dell'email in modale (componi_email del motore, cid:logo sostituito con /logo.png nel solo preview), segnaposto cliccabili, azioni con sondaggio 900ms e terminale verde live, password nascosta con occhio (vuota = resta quella salvata)
- apri_pannello.bat (nuovo, CRLF): avvia runtime\python.exe pannello.py ridotto a icona
- controlla_pc.py: nuovo controllo 6-quater (pannello.py/pannello.html presenti)
- LEGGIMI.txt: sezione 1-bis IL PANNELLO + elenco file + sezioni 3/8/10 aggiornate (patch CRLF verificata)
- Bug corretti in itinere: collisione chiave "nome" (centro vs prodotto) -> leggi_config_raw per sezione; SystemExit del motore non catturato -> BaseException; righe vecchie del blocco multiriga dopo riga vuota non rimosse -> lookahead; lista "Oggi" mostrava le prossime scadenze -> oggi_elenco dal server; plurale "inviato/i"
- test_pannello.py: 40/40 OK (avvio, HTML, stato+conteggi, salvataggi config al posto giusto con CRLF/commenti/chiavi estranee/password intatte, roundtrip, anteprima vera con segnaposto e riga extra, azione prova codice 0, controllo PC senza blocco input, registro ultime-prima, 404, doppia azione rifiutata, seconda istanza esce 0 e la prima vive, niente __pycache__)
- Regressioni: v4 53/53, v4.1 33/33, v4.2 22/22, v4.3 42/42, v4.4 52/52 -> TOTALE 242 verdi
- Screenshot reali (agent-browser, demo "Centro Revisioni Aurora" con 8 clienti): 6 foto in download/foto_pannello; script scripts/foto_pannello.sh (click via eval, demo su 0.0.0.0 solo con env PANNELLO_DEMO)
- RevMemo_v4.zip ricreato (252 file, 10.9 MB, integrità OK, file cliente assenti) e caricato su temp.sh insieme alle 6 foto

Stage Summary:
- v4.5: il centro fa tutto con apri_pannello.bat (bottone HOME con stato/azioni, MESSAGGIO con anteprima vera, CLIENTI, IMPOSTAZIONI, REGISTRO); config.ini e .bat restano validi come via alternativa
- Risposta hosting (per Riccardo): Vercel Hobby VIETA uso commerciale; Supabase free mette in pausa i progetti; Gmail SMTP/IMAP da server cloud è il vero ostacolo tecnico; ora = pannello locale (zero costi, zero GDPR cloud), GitHub privato sì per backup, SaaS dopo con VPS piccolo quando le entrate lo coprono

---
Task ID: 17
Agent: main (Z.ai)
Task: Utente ha visto le foto del pannello e approva ("OK MI PIACE") -> consegna finale RevMemo v4.5 (pannello grafico).

Work Log:
- Sanity check post-approvazione: test_pannello.py 40/40 OK; ZIP 252 file, integro, nessun file utente dentro, revmemo.py = 4.4 (PRO) invariato, pannello.py 572 righe + pannello.html 651 righe + apri_pannello.bat presenti
- Ricaricato RevMemo_v4.zip su temp.sh (i link della sessione precedente non erano stati registrati): https://temp.sh/SSuYP/RevMemo_v4.zip
- Verifica end-to-end: POST sul link = 11433824 bytes, sha256 75252dfe32f5d924... IDENTICO all'originale, 252 file, zip integro
- Pulizia file temporaneo verifica_download.zip

Stage Summary:
- Link finale consegnato: https://temp.sh/SSuYP/RevMemo_v4.zip (3 giorni, pagina nera con bottone di download POST)
- v4.5 = motore 4.4 invariato + pannello grafico locale (Home/Messaggio/Clienti/Impostazioni/Registro) apribile con apri_pannello.bat; .bat e config.ini restano come via alternativa
- 242 controlli totali verdi (v4 53 + v4.1 33 + v4.2 22 + v4.3 42 + v4.4 52 + pannello 40)

---
Task ID: 18
Agent: main (Z.ai)
Task: Utente chiede "ma si fa tutto da pannello? tipo si aggiornano i clienti con email e tutto il resto?" -> v4.6 GESTIONE CLIENTI COMPLETA NEL PANNELLO.

Work Log:
- pannello.py: nuove API /api/clienti (lista completa con righe), /api/cliente/nuovo, /api/cliente/modifica, /api/cliente/stato (fatta/attivo); SERRATURA_EXCEL; _valida_cliente (email, targa, scadenza via converti_scadenza del motore + anno>=2000 + normalizzazione 3/2026->03/2026); _carica_per_scrittura (backup_excel del motore PRIMA di ogni scrittura, load senza data_only, individua_colonne); ERRORE_EXCEL_APERTO su PermissionError (chmod 444 in test); targa doppia rifiutata; _prima_riga_libera riempie i buchi fino a riga 300; file assente -> crea_excel_se_manca del motore + pulizia degli 8 esempi (file vuoto pronto); scadenza sempre number_format "@"; Attivo=Fatta default NO/SÌ
- REGOLA STOP rispettata: STOP non scrivibile dal pannello (solo il motore via IMAP); RIATTIVA scrive SÌ (come da istruzioni del motore)
- BUG corretto in itinere: _attivo_visuale() - il motore scrive "SÌ" (accentato) e accetta SI/S/SÌ/TRUE/VERO/1/YES/Y: la visualizzazione vecchia mostrava NO per le celle "SÌ" del file creato dal motore; normalizzata in STOP/SI/NO sia in lista_clienti sia in dati_stato.elenco
- pannello.html: scheda CLIENTI rifatta (➕ NUOVO CLIENTE, tabella completa con toggle SÌ/NO cliccabili, ✏️ Modifica, RIATTIVA per STOP), seconda modale (form cliente: nome/targa/email/telefono/scadenza, Enter=salva, Esc=Annulla), CSS (button.badge, .bottone.piccolo, #velo-cliente, .contenuto)
- test_pannello.py: da 40 a 65 controlli (lista, nuovo+normalizzazione+default, targa doppia, email/scadenza/senza nome/anno corto rifiutati, modifica con Attivo-Fatta intatti, fatta SÌ, riattiva STOP, STOP rifiutato, Excel aperto, buco riga 6, file assente->creato senza esempi con tendine+colori, backup automatico)
- Regressioni: v4 53/53, v4.1 33/33, v4.2 22/22, v4.3 42/42, v4.4 52/52 -> TOTALE 267 verdi (202 motore + 65 pannello)
- LEGGIMI.txt: bullet CLIENTI della sezione 1-bis riscritto (patch CRLF)
- Foto aggiornate (7, demo Aurora): 4_clienti.png nuova + 7_nuovo_cliente.png (modale compilata)
- RevMemo_v4.zip ricreato: 252 file, 10.9 MB, integrità OK, file cliente assenti; caricato su temp.sh https://temp.sh/SWQJV/RevMemo_v4.zip e verificato sha256 952b92a2... IDENTICO
- Foto su temp.sh: https://temp.sh/jDpQv/4_clienti.png e https://temp.sh/SDmhk/7_nuovo_cliente.png

Stage Summary:
- v4.6: dal pannello si fa TUTTO: clienti (nuovo/modifica/SÌ-NO/riattiva), messaggio con anteprima, impostazioni, prova/invio/colora/controllo, registro. Excel resta via alternativa sincronizzata. Backup automatico prima di ogni cambiamento.
- Link finale: https://temp.sh/SWQJV/RevMemo_v4.zip (3 giorni, POST download)

## Task 18 — FIX "il pannello non si apre" (v4.6)
Problema: apri_pannello.bat lanciava pannello.py con `start /min` (finestra ridotta a icona): con qualsiasi errore (file bloccati da Windows/antivirus, crash avvio, browser non aperto) la finestra spariva subito = "non si apre nulla", nessun messaggio.
Correzioni:
- apri_pannello.bat riscritto: avvio VISIBILE (niente /min), sblocco automatico Zone.Identifier (MotW) via PowerShell solo se presenti, messaggi chiari, pause se errore / chiusura corta se ok
- pannello.py main() blindato: log avvio in pannello_log.txt a ogni passo, porta di riserva 8768→8772 (se occupata da altro programma), riconoscimento pannello già attivo (GET /api/stato = nome_centro), fallback se il browser non parte (mostra l'indirizzo da copiare), wrapper __main__ che cattura QUALSIASI errore e lascia la finestra aperta in italiano con "fai una foto"
- NUOVO controlla_pannello.py/.bat: diagnostica completa (cartella scrivibile, runtime, openpyxl, revmemo, config.ini, excel, porte 8768/8769, MotW, browser, ultime righe log) -> salva assistenza_pannello.txt
- LEGGIMI aggiornato (finestra visibile, sezione 9 con i nuovi casi, elenco file)
Test: test_pannello.py esteso (71 controlli: +cambio porta, +bat senza /min, +controlla_pannello compila) = 71/71 VERDI. revmemo.py intatto (4.4 PRO).
ZIP: 258 file, 11.457.683 byte, sha256 b00c25d051ad6b8ecbfd2592969faaace60512454b0d12aacec6248e418449a5; esclusi file cliente/test (config.ini, clienti_revisioni.xlsx, log, backup, __pycache__).
Upload: https://temp.sh/OqZKG/RevMemo_v4.zip — verificato via POST: sha256 identico, testzip OK, bat senza /min, controlla_pannello incluso.

## Task 19 — FIX "No module named 'revmemo'" sul runtime incluso (v4.6b)
Diagnosi GRACE ai messaggi del cliente (la finestra di errore ha funzionato):
- runtime portatile (python312._pth) = modalita' ISOLATA: la cartella dello script NON va in sys.path -> "import revmemo" fallisce con ModuleNotFoundError. Il motore lanciato dai .bat funziona (e' lui lo script); fallivano solo pannello.py e controlla_pannello.py. I test sandbox non l'avevano colto perche' usano Python di sistema.
Correzioni:
- pannello.py e controlla_pannello.py: sys.path.insert(0, BASE) prima di import revmemo (BASE = cartella pacchetto)
- runtime/python312._pth: aggiunta riga ".." (root pacchetto sempre visibile al runtime)
- controlla_pc.py NON richiede fix (non importa revmemo)
- test_pannello.py: regressione "runtime isolato" con python -I (simula ._pth) -> 72/72 VERDI
- prova reale: python3 -I controlla_pannello.py -> "revmemo importato [OK]"
ZIP: sha256 fa7065f8462516d15583159bd574c7f9136ab06d4f8114b3ef5db3399b3ce5a1; upload https://temp.sh/jliga/RevMemo_v4.zip verificato via POST (sha256 identico, fix inclusi).
Istruzioni al cliente: estrarre in cartella NUOVA; copiare config.ini / clienti_revisioni.xlsx SOLO se gia' compilati.

## Task 19 — RevMemo v4.7: pannello rivisto sulle richieste del cliente (2026-10-28)
Richieste del cliente (tutte implementate):
1. Home: PROVA/INVIA bloccavano il pannello -> pulsanti FERMA (endpoint
   POST /api/azione-ferma, terminate del processo figlio, flag "fermata")
   e "Chiudi finestra" (nasconde il terminale, l'azione finisce da sola).
2. Pulsante "Colora Excel" rimosso da Home e Clienti (tipo "colora" tolto
   anche da lancia_azione; colora_excel.bat resta nel pacchetto).
3. "Controllo PC" spostato in Impostazioni (carta propria con terminale
   dedicato uscita-controllo); _esegui unificato su Popen (stdin per l'INVIO
   di controlla_pc, timeout 300s mantenuto).
4. Carta "Email Gmail" spostata da Home a Impostazioni (stato-gmail dentro
   la carta Gmail; caricaStato() richiamata all'apertura di Impostazioni).
5. Tab Messaggio: tooltip (title) su OGNI campo, bottoni e 13 segnaposto
   (con spiegazioni); SALVA e ANTEPRIMA SEPARATI: ANTEPRIMA fa GET e non
   salva piu', con guard "MESSAGGIO_SALVATO" (avviso se modifiche non
   salvate); label "Riga extra nel piede" -> "Riga extra".
6. Tab Clienti: scadenza mostrata COME SCRITTA nel file (nuova funzione
   _testo_scadenze in pannello.py: riga->testo grezzo, usata in
   dati_stato e lista_clienti; 09/2026 resta 09/2026), dicitura
   "Fatta" -> "Revisione fatta" (intestazione + tooltip), ricerca clienti
   (k-cerca/b-cerca-clienti/Mostra tutto, disegnaClienti()).
7. Tab Registro: ricerca (r-cerca + Mostra tutto, disegnaRegistro()) e
   nota "i 250 piu' recenti"; cerca cliente anche in HOME (h-cerca/
   b-cerca-home -> tabella risultato).
8. controlla_pannello.py: fix bug cosmetico ("manca: estrai TUTTO lo zip"
   appariva anche quando tutto OK -> dettaglio condizionale).
Test: test_pannello.py esteso a 84 controlli (novita' 4.7, FERMA senza
azione, scadenza libretto 03/2026, posizione Controllo PC/stato-gmail,
controlla_pannello condizionale) -> 84/84 OK; suite motore v4.4 -> 52/52 OK;
py_compile OK; node --check sul JS OK; verifica id incrociata (70/70).
Foto pannello NON rigenerabili in questa sessione (rete agent-browser
isolata: ERR_CONNECTION_REFUSED anche su 127.0.0.1; foto_pannello.sh
corretto comunque: demo ora copia anche revmemo.py/config ecc.).
LEGGIMI.txt patchato per la 4.7 (patch_leggimi_v47.py, CRLF mantenuto).
ZIP: ricrea_zip_revmemo_v45.py -> 258 file, 11.462.741 byte,
sha256 5c334b24a331e767e20cab1ea0061744a6c68158a5c745042fb3b8f47a2234e7,
testzip OK; upload https://temp.sh/SnuKC/RevMemo_v4.zip e verifica
download POST: sha256 identico.

## Task 20 — RevMemo v4.8: chiusure, registro Excel, statistiche, email qualsiasi (2026-10-28)
Richieste del cliente (tutte implementate):
1. Home: card "Oggi" rinominata "Email da inviare oggi" (conferma: li' ci
   sono esattamente i promemoria che partono con INVIA ORA).
2. AZIONI: il pulsante CHIUDI restava nascosto a lavoro finito (bug v4.7):
   ora la finestra del risultato resta finche' il centro non la chiude,
   sia durante che DOPO l'azione (JS: bottoneChiudi(term), fine-branch
   senza classList.add("nascosto")).
3. CONTROLLO PC: aggiunto bottone "✖ Chiudi" (b-chiudi-controllo) che
   nasconde uscita-controllo; visibile durante e dopo il controllo.
4. Prossime scadenze: risposta alla domanda "si compila all'infinito?" ->
   NO, sempre 6, chi scade prima in alto; nota esplicita nella card.
5. Clienti: SÌ/NO (e STOP) centrati sotto le intestazioni "Revisione
   fatta" e "Attivo" (tabella(testate, righe, centrate) + th.ctr/td.ctr;
   centrati anche nel risultato del Cerca in Home).
6. REGISTRO: nuovo bottone "📊 ESPORTA IN EXCEL" (POST /api/registro-excel
   -> esporta_registro(): TUTTO il log, non solo 250 righe, in
   registro_invii.xlsx con intestazione blu, larghezze, freeze A2,
   date gg/mm/aaaa e tipo tradotto "1 mese/10 giorni/1 giorno", poi lo
   apre; PermissionError gestito "file APERTO").
7. REGISTRO: STATISTICHE INVII MENSILI (proposta v4.7 accettata):
   statistiche_invii() legge TUTTO il log, /api/registro ritorna
   "statistiche" = {mesi: ultimi 6 (etichette it: gen/feb...), questo_mese,
   totale}; card "Statistiche invii" con grafico a barre CSS (colonna
   corrente in blu) + riga "Questo mese / Totale".
8. EMAIL NON SOLO GMAIL (confermato al cliente):
   - revmemo.py: Sender._connetti usa STARTTLS se porta=587 (smtplib.SMTP
     + starttls), altrimenti SMTP_SSL; porta robusta in carica_config
     (non numerica -> 465); host_imap_da mappa Outlook/Hotmail/Live ->
     outlook.office365.com; CONFIG_DEFAULT/commenti aggiornati con la
     tabella Outlook/Libero/Virgilio/Aruba.
   - pannello.py: CHIAVI smtp ora include host/porta; /api/impostazioni
     valida server (deve avere il punto) e porta (1-65535) con messaggi
     in italiano; host senza porta -> porta 465 da sola.
   - pannello.html: carta "Email del centro" (gmail O altra) + nuova carta
     "Casella email diversa da Gmail" (i-host/i-porta con esempi); stato
     email a 3 pallini: verde ok / AMBRA "parziale" (casella non Gmail ma
     server ancora Gmail -> compila Server e Porta) / rosso da_fare
     (dati_stato: email_stato + email_messaggio + smtp_host/smtp_porta).
   - controlla_pc.py: messaggi generici + controllo "casella non Gmail ma
     server Gmail" (MALE con la tabella dei server) + "Server email: host porta".
   - config_ESEMPIO.ini: commenti [smtp] generici con esempi.
9. LEGGIMI.txt patchato (patch_leggimi_v48.py, CRLF): 1-bis aggiornato
   (EMAIL DA INVIARE OGGI, CHIUDI, statistiche, export, Server/Porta) +
   nuovo blocco "NON USI GMAIL?" in sezione 5.
Test: test_pannello.py esteso a 112 controlli (v4.8: card oggi, chiudi a
fine azione e su controllo, SI/NO centrati, nota scadenze, i-host/i-porta,
statistiche 6 mesi con riga di oggi nel log, export registro completo
(intestazioni, 3 righe, tipo tradotto, data gg/mm/aaaa, secondo export),
host/porta salvati nel posto giusto, porta/server non validi rifiutati,
stato parziale, motore starttls/imap outlook) -> 112/112 VERDI;
suite motore v4.4 -> 52/52 VERDI; smoke test host_imap_da + porta
robusta OK; node --check sul JS OK; verifica id incrociata (77 id).
ZIP (registro_invii.xlsx aggiunto agli ESCLUSI): 258 file, 11.466.921
byte, sha256 5a6311629b856cfffcbdae27f0056a046d510cf021ea443ecf5d66cdab830930,
testzip OK; upload https://temp.sh/IhtZX/RevMemo_v4.zip e verifica
download POST: sha256 identico, contenuto v4.8 confermato.
Miglioramenti PROPOSTI al cliente (non ancora fatti): orario di invio
automatico ogni mattina; backup settimanale completo della cartella;
password di accesso al pannello; archivio anni precedenti del registro.

## Task 21 — RevMemo v4.9: feedback pannello (2026-10-29)
Richiesta cliente: "che senso ha risposta_a se poi stop non funziona?",
"migliora perché virgilio non ho capito", "statistiche invii non lo
utilizzo da sei mesi, lo sto utilizzando da ottobre, 6 mesi indietro non
ha senso", "prossime scadenze ci sara' uno scroll oltre 6? non ho
capito", "cerca cliente in home toglilo", "le statistiche stanno meglio
in home (dashboard)".
1. "RISPOSTE VERSO" TOLTO (pannello.html): il campo confondeva — se
   compilato lo STOP automatico non leggeva piu'. Le risposte arrivano
   SEMPRE nella casella mittente: scritto anche nell'aiuto sotto il
   campo email. pannello.py _salva_senza_risposta: forza
   puliti["risposta_a"] = "" ad ogni salvataggio smtp (svecchia le
   config vecchie). Commenti aggiornati in revmemo.py/controlla_pc.py/
   config_ESEMPIO.ini. Motore: Reply-To resta supportato solo via
   config a mano, ma il pannello lo azzera.
2. MENU PROVIDER (pannello.html): al posto della lista testuale confusa
   ora c'e' un menu a tendina "La tua casella e' di..." (Gmail,
   Outlook/Hotmail, Libero, Virgilio, Aruba, Altro): scelto il nome,
   host+porta si scrivono da soli (PROVIDER map); "Altro" svuota e
   focus su Server. providerDaHost() pre-seleziona il menu dal server
   salvato. CSS: select aggiunto ai selettori .campo.
3. STATISTICHE MENSILI DINAMICHE (pannello.py statistiche_invii()):
   finestra dal MESE DEL PRIMO INVIO nel registro a oggi (max 12
   colonne); senza dati = solo mese corrente a zero. Chi usa da ottobre
   vede ottobre, non 6 mesi indietro.
4. STATISTICHE SPOSTATE IN HOME (dashboard): carta "Statistiche invii"
   subito dopo i numeri dei clienti (stat-invi dentro vista-home),
   tolta dal Registro; /api/stato ora include "statistiche";
   caricaStato disegna il grafico, caricaRegistro non piu'.
5. PROSSIME SCADENZE CHIARE: desc riscritta ("Non c'e' uno scroll:
   l'elenco resta corto di proposito...") + /api/stato ritorna
   "prossime_altre" e il JS aggiunge "…e altre N con scadenza piu'
   lontana: le trovi nella scheda Clienti".
6. CERCA CLIENTE TOLTO DALLA HOME (su richiesta: "se poi mi dice di
   andare in clienti? toglilo"): via card, JS cercaInHome e listener;
   CERCA resta in Clienti e Registro.
Test: test_pannello.py 119/119 VERDI (nuovi: cerca solo in
clienti/registro, no scroll + resta corto di proposito, menu provider
con i 5 server smtp, stat-invi in Home e fuori dal Registro (sezione),
i-risposta assente, risposte/STOP spiegati, prossime_altre, statistiche
in /api/stato == mesi attesi dal 2026-01 (max 12), risposta_a vecchio
svuotato al salvataggio, mesi dinamici anche su /api/registro); motore
52/52; verifica_html (node --check + 74 id) OK.
LEGGIMI.txt patchato v4.9 (patch_leggimi_v49.py, CRLF): 1-bis HOME,
IMPOSTAZIONI e REGISTRO aggiornati.
ZIP: 258 file, 10.9 MB, sha256
4699237d1f25f6ea479a87538e659696c6b107cfd47193e75c27148cc4804d17,
testzip OK; upload https://temp.sh/WDDRl/RevMemo_v4.zip e verifica
download POST: sha256 identico, contenuto v4.9 confermato (menu
provider, stats in Home, i-risposta/h-cerca assenti, prossime_altre,
LEGGIMI aggiornato).
Nota upload: temp.sh vuole PUT per -T? NO: -T su temp.sh/<nome> da 404;
funziona `curl -F file=@... https://temp.sh/upload` e il DOWNLOAD va
fatto con `curl -X POST -L` (GET restituisce la pagina HTML).

## Task 22 — RevMemo v4.10: scroll 10 in Home, menu 20/30/50, "30 giorni" (2026-10-29)
OK del cliente ("ok metti anche 50 e procedi") su: 1 mese -> 30 giorni,
scroll 10 su "Email da inviare oggi" e "Prossime scadenze", scroll +
menu 20/30/50 su Elenco clienti e Registro.
1. "1 MESE" -> "30 GIORNI" ovunque: label "Oggetto — 30 giorni" + tooltip
   (pannello.html), tipoLbl registro JS, _TIPO_PROMEMORIA (pannello.py,
   usato anche dall'export Excel), etichette motore "30 giorni prima"
   (revmemo.py; note log nuovi invii "30 giorni prima (30 gg)"; chiavi
   interne finestra_1mese/oggetto_1mese INVARIATE per compatibilita'
   config), footer dei file Excel di esempio (patch_excel_footer_v410.py
   su clienti_revisioni.xlsx e clienti_ESEMPIO.xlsx) + genera_excel_clienti.py,
   LEGGIMI.txt (patch_leggimi_v410.py, CRLF), ISTRUZIONI_CENTRO.md,
   GUIDA_INSTALLATORE.md. "1 mese" residuo nel pacchetto: ZERO.
2. LISTE DI HOME COMPLETE CON SCROLL 10 (pannello.py + pannello.html):
   oggi_elenco senza taglio [:8], prossime = tutte_prossime (niente [:6],
   tolta la nota "…e altre N"); fissaAltezza()/fissaLista() misurano
   intestazione+righe reali e impostano maxHeight per 10 righe; CSS
   .scroll-y (overflow) + .scroll-y th{position:sticky} = intestazione
   FISSA durante lo scroll; tabella() ora emette thead/tbody.
   BUG COLTO DALLE FOTO: fissaAltezza/fissaLista non aggiungevano la
   classe scroll-y -> aggiunta (classList.add) altrimenti niente overflow.
3. MENU VEDI 20/30/50 RIGHE (pannello.html): select #k-righe (Clienti,
   accanto a CERCA) e #r-righe (Registro, dopo ESPORTA IN EXCEL),
   classe .menu-righe; scelta salvata in localStorage
   (revmemo_righe_clienti/revmemo_righe_registro, try/catch),
   righeAttuali() rilegge a ogni disegno.
4. BUG REALE TROVATO E CORRETTO (demo 28 clienti): file con NOTA UNITA
   (clienti_revisioni.xlsx di esempio ha footer unito B14:H14) ->
   _prima_riga_libera sceglieva la riga della nota (email/targa/scadenza
   vuote) e la scrittura crashava "'MergedCell' object attribute 'value'
   is read-only" (HTTP 500). FIX: le righe dentro un merge vengono
   SALTATE (il cliente va nella prima riga libera DOPO la nota) e
   _scrivi_campi_cliente stacca preventivamente i merge sulla riga
   target (unmerge_cells). Regression test dedicato: file con nota unita
   B7:H7 -> nuovo cliente riga 6, secondo riga 8, nota intatta, 3 clienti.
5. FOTO DI CONTROLLO (scripts/foto_v410.sh, foto_pannello/v410_*.png):
   home con 36 clienti demo, prossime scadenze = 10 righe visibili in
   ordine, clienti con menu "Vedi 20 righe" + scroll con intestazione
   fissa, registro con "30 giorni" e menu, messaggio con "Oggetto —
   30 giorni" allineato. Nota tecnica: foto_pannello.sh vecchio sediva
   "PORTA" ma ora e' "porta": la demo resta su 127.0.0.1 e il browser
   apre direttamente 127.0.0.1.
Test: test_pannello.py 132/132 VERDI (nuovi: label 30 giorni senza
"1 mese"/"UN MESE", tipoLbl registro, menu 20/30/50 + chiavi localStorage,
fissaLista/fissaAltezza su lista-oggi e tabella-prossime, thead+sticky,
email oggi senza taglio a 8, merged-skip sorgente + comportamentale,
export tipo "30 giorni"); motore 52/52; verifica_html OK (node --check,
76 id); py_compile OK.
ZIP: 258 file, 10.9 MB, sha256
10d1749fe6b9acb54a6b25907bebee5d09baf5da358126231337336c61416466,
testzip OK; upload https://temp.sh/PNbXq/RevMemo_v4.zip e verifica
download POST: sha256 identico, v4.10 confermato nel contenuto.

## 2026-10-29 — v4.11 (conferme interne, menu 100, istruzioni Registro)
Richiesta cliente: (a) istruzioni Registro aggiornate perche' c'e' gia' il
menu righe e aggiungere anche 100; (b) domanda: cambiando email serve
toccare config.ini? (c) togliere "127.0.0.1:8768 dice:" dalla finestra di
conferma; (d) "verifica tutto, sii critico, trova bug".
1. ISTRUZIONI REGISTRO (pannello.html): "Qui vedi i 250 più recenti" ->
   "Qui vedi gli invii più recenti: col menu qui sotto scegli quante righe
   vedere (20, 30, 50 o 100), le altre si raggiungono scorrendo con la
   barra a destra...". Anche il title di "✖ Mostra tutto" senza piu' "250".
   L'export Excel resta su TUTTI gli invii (title invariato, corretto).
2. MENU RIGHE 20/30/50/100: RIGHE_MENU=[20,30,50,100], option value="100"
   ("Vedi 100 righe") su #k-righe e #r-righe, tooltip "20, 30, 50 o 100",
   commento blocco scorimento aggiornato. Scelta sempre su localStorage
   (revmemo_righe_clienti/revmemo_righe_registro); righeAttuali accetta 100.
3. CONFERME INTERNE (fine del "127.0.0.1:8768 dice:"): il prefisso era del
   window.confirm NATIVO del browser (Chrome mostra l'indirizzo della
   pagina). Niente alert/prompt nel progetto: creato modal del pannello
   #velo-conferma (z-index 58, sopra velo 50 e velo-cliente 55) con
   chiediConferma(testo, tastoOk) -> Promise; 5 siti convertiti con
   "await": INVIA ORA / PROVA (azione), FERMA ("SÌ, FERMA"), riattiva
   ("SÌ, RIATTIVA"), revisione fatta ("SÌ, FATTA"), pausa ("SÌ, PAUSA").
   No/Annulla/ESC/clic fuori = No (a prova di errori); conferma gia'
   aperta + nuova = la vecchia vale No (niente promesse appese); focus su
   "Sì, procedi" (Invio conferma, come il confirm nativo); white-space:
   pre-line per il \n del messaggio di INVIA ORA. Il toast avviso() era
   gia' senza indirizzo: non e' stato toccato.
4. EMAIL -> CONFIG.INI: verificato nel codice, la risposta e' "fa da solo":
   /api/impostazioni -> _salva -> salva_campi() riscrive config.ini al
   posto giusto (commenti, ordine, CRLF, chiavi sconosciute intatte;
   password vuota = resta quella salvata; risposta_a svuotata). Niente da
   cambiare a mano, mai.
5. CRITICA/BUG: py_compile aveva creato __pycache__ dentro il pacchetto ->
   il check "pacchetto vero senza __pycache__" l'ha beccato (135/135 dopo
   pulizia); verifica_html 81 id (5 nuovi della conferma); nessun alert(
   residuo; foto v411: menu "20,30,50,100", altezza tabella 696px->3336px
   col 100, scelta memorizzata, conferma con SOLO la domanda, "No, annulla"
   chiude senza avviare nulla.
Test: test_pannello.py 135/135 VERDI (nuovi: menu 20/30/50/100 + tooltip +
RIGHE_MENU, istruzioni Registro senza "250 più recenti", conferme interne
con >=5 siti await + niente window.confirm/alert, etichette tasti);
motore e verifica_html OK; py_compile OK.
ZIP: 258 file, 10.9 MB, sha256
20dd9566d6b16caa6519660271fb4be946e870653a8e5c0f91e54531522feb05,
testzip OK; upload https://temp.sh/xkdFD/RevMemo_v4.zip e verifica
download POST: sha256 identico, v4.11 confermato nel contenuto.

## 2026-10-29 — v4.12 (reset fatta da solo, ELIMINA cliente, aiuto password)
Richiesta cliente: ok all'opzione A (scadenza nuova -> "Revisione fatta"
torna NO da sola), domanda sullo STOP (spiegato: Attivo=STOP blocca GIA'
tutto, niente da cambiare; non toccare fatta allo STOP perche' RIATTIVA
deve ripartire pulita), manca il bottone elimina cliente ("risolvi").
1. RESET AUTOMATICO (pannello.py modifica_cliente): confronto le date
   PARSE con r.converti_scadenza (vecchia cella vs nuova): se DIVERSE ->
   fatta="NO" scritto da solo + risposta {"fatta_reset": true}; stessa
   data anche scritta diversa (09/2026 vs 30/09/2026) -> niente cambia.
   Nota di log "; scadenza nuova -> revisione fatta torna NO". Frontend:
   toast "Cliente aggiornato ✔ (scadenza nuova: revisione fatta è tornata
   NO da sola)".
2. ELIMINA CLIENTE (pannello.py elimina_cliente + /api/cliente/elimina):
   bottone 🗑 Elimina (rosso) accanto a Modifica per ogni riga; conferma
   interna chiediConferma "Elimino davvero NOME (TARGA)? ..." con tasto
   "SÌ, ELIMINA"; ws.delete_rows(riga,1) con gestione MANUALE delle zone
   unite (openpyxl non le sposta: stacco le merge sotto/intersecanti,
   delete, rimetto le sotto un posto piu' su — la nota a fondo tabella
   scende/integra); guard: riga dentro una zona unita = rifiutata ("non
   e' un cliente"), riga vuota/assurda rifiutata; backup automatico via
   _carica_per_scrittura; log "cliente ELIMINATO: riga N (nome, targa)".
3. AIUTO PASSWORD (pannello.html): "Gmail: la Password per app di Google
   (16 caratteri). Con Libero, Virgilio e simili: la password normale
   della casella."
4. STOP: NESSUNA modifica (verificato nel codice: il motore controlla
   Attivo PRIMA di fatta -> STOP ferma gia' tutte le email; scrivere
   fatta=SÌ allo STOP romperebbe RIATTIVA che lascerebbe fatta=SÌ).
5. FOTO (scripts/foto_v412.sh, foto_pannello/v412_*.png): bottoni
   Modifica/Elimina per riga, conferma con nome+targa, "Cliente
   eliminato ✔" e conteggio 8->7, reset fatta con toast e badge NO.
Test: test_pannello.py 151/151 VERDI (nuovi: bottone+endpoint+SÌ ELIMINA,
avviso reset + fatta_reset, aiuto password; comportamentali: reset con
scadenza nuova scritto nell'Excel, stessa data non toccata, elimina con
righe che salgono, nota B7:H7 -> B6:H6 integra, riga-nota protetta,
nuovo cliente prima riga libera dopo elimina, doppio elimina, vuota/
assurda rifiutate); verifica_html 81 id; py_compile OK.
ZIP: 258 file, sha256
5a8b300949128d82921c296450fb7cf9f4df82bc21876d06fddfb32b290812c4,
testzip OK; upload https://temp.sh/ZCKIU/RevMemo_v4.zip e verifica
download POST: sha256 identico, v4.12 confermato nel contenuto.

## 2026-10-29 — archivio completo su GitHub (richiesta cliente)
Cliente: "ok metti tutto su github così non perdo nulla".
- Nessun token GitHub nel contenitore (env senza vari GitHub, assenti
  ~/.git-credentials ~/.netrc ~/.config/gh ~/.ssh; grep ghp_/github_pat_
  su tutto il progetto: nulla). Rete OK: github.com 200, api.github.com
  /zen risponde. Serve la chiave (token) dell'account del cliente.
- Repository LOCALE creata: /home/z/my-project/revmemo-github (git init
  -b main, commit 2decc89, tag v4.12, utente "RevMemo Dev"). Contenuto:
  revmemo/ (pacchetto 24M con runtime Windows), dist/ (RevMemo_v4.zip =
  v4.12 + archivio/RevMemo_v3.zip), scripts/ (pulita da __pycache__ e
  *.log), foto/, archivio/promemoria-revisioni/, worklog.md, README.md,
  .gitignore. 58M lavoro, .git 48M (pronta per il push).
- Controllo dati sensibili PRIMA del commit: config.ini pacchetto =
  modello con password vuota; test_env/config.ini = password finta
  ("abcd efgh ijkl mnop"); clienti_revisioni.xlsx = solo esempi (Mario
  Rossi AB123CD...); nessun reminders_log/backup/*.pyc nel pacchetto.
  OK per repository privata.
- scripts/push_github.sh pronto: crea la repo PRIVATA "revmemo" via API
  e fa push main --tags; basta UTENTE_GITHUB + CHIAVE_TOKEN del cliente.
- Archivio di sicurezza RevMemo_progetto_completo_2026-10-29.zip (senza
  .git, 46M, 379 file) sha256
  642c3c2f0c4b15a5ac6735e7eca0735a9833824bed163a1f9ad1951952e436db,
  testzip OK; upload https://temp.sh/sdnGu/RevMemo_progetto_completo_
  2026-10-29.zip e verifica download POST: sha256 identico.
- Ricontrollo v4.12: test_pannello.py 151/151 verdi.

## 2026-10-29 — push su GitHub completato
- Cliente ha fornito utente + chiave (token) dell'account riccardofreddi.
  Token verificato via GET /user (login OK). push_github.sh: creata repo
  PRIVATA riccardofreddi/revmemo, push main + tag v4.12 OK.
- Verifica API dopo il push: private=True, default_branch=main, radice
  completa (.gitignore README.md archivio dist foto revmemo scripts
  worklog.md), etichetta v4.12 presente.
- Igiene: chiave rimossa dall'URL remote in .git/config (resta solo
  l'URL pulito); la chiave NON e' stata scritta nel worklog ne' in
  alcun file nel repository.
- Copia di sicurezza temporanea consegnata al cliente: temp.sh/sdnGu
  (RevMemo_progetto_completo_2026-10-29.zip, sha256 642c3c2f...).

## 2026-10-29 — patch esterna v4.13: revisione completa (NON applicata)
- Altra sessione (chiave fornita dal cliente) ha corretto 15 bug e
  pubblicato il ramo fix/revisione-bug-v4.13 (commit 6b49f3e, 11 file,
  +843/-87). Consegnata anche revmemo_v4.13.patch: risultato IDENTICA
  byte per byte al diff del ramo (1478 righe, 11 file).
- VERIFICA NOSTRA (su copia /tmp, mai sul pacchetto vero): git apply
  pulita; test_correzioni_v413.py 55/55 VERDI su codice patchato e
  crolla sul vecchio (i test beccano davvero i bug); suite pannello
  151/151 (meglio dei 150/151 dichiarati: il rosso "Excel aperto"
  dipendeva dall'utente root del loro ambiente); v43/v44 verdi;
  v4/v41/v42 con 2 rossi l'una MA identici anche sul codice VECCHIO
  (pre-esistenti, non colpa della patch); verifica_html SUPERATA;
  py_compile OK; scansione sicurezza della patch (urllib/eval/exec/
  base64/socket): nessun codice sospetto, solo test locali.
- Tutti i 15 bug dichiarati sono VERI (verificati nel codice). Il piu'
  grave: fin dall'inizio QUALSIASI sito aperto nel browser poteva
  comandare il pannello su 127.0.0.1 (cambiare il server di posta e
  leggere la password, cancellare clienti, far partire invii) via
  CSRF/DNS rebinding. Le difese nuove (chiave di sessione + Host +
  Origin) coprono TUTTE le vie API: il pannello ha solo do_GET/do_POST
  e ogni operazione cliente (nuovo/modifica/elimina) passa da POST.
- CRITICITA' TROVATE DA NOI (da sistemare prima di consegnare):
  1. controlla_pannello.py NON aggiornato: chiama /api/stato senza
     chiave -> 403 -> risponde "porta occupata da un altro programma"
     invece di "pannello gia' aperto" (regressione vera).
  2. pannello.py _esegui cancella revmemo.lock quando il motore esce
     con codice non 0/2: se esce 1 PRIMA di prendere il lock (es.
     password mancante) mentre un invio pianificato sta girando, gli si
     cancella il lock VIVO (corsa rara ma possibile). Meglio cancellare
     solo se il lock e' vecchio (>15 minuti).
  3. Scheda browser lasciata aperta + pannello riavviato = chiave
     vecchia -> 403 "riapri il pannello" (comportamento nuovo, da
     documentare); manca Cache-Control: no-store sulla pagina col token.
  4. Test "SMTP_SSL con contesto verificato" non verifica davvero il
     contesto (assert debole: tentativi>=1).
  5. stop_gestiti.txt e revmemo.lock in BASE: cartella non scrivibile =
     fix 4/5 spenti in silenzio (accettabile sul PC del centro).
  6. Con i certificati verificati, un antivirus che intercetta la posta
     puo' FERMARE gli invii (prima partivano insicuri): previsto
     messaggio d'aiuto, ma va spiegato in LEGGIMI.
- Non ancora fatto (dichiarato anche da loro): lock Excel tra pannello
  e motore; README/worklog/tag/ZIP v4.13.
- AZIONI: diff completo del ramo salvato in upload/branch_v413.diff;
  ramo CANCELLATO da GitHub su richiesta cliente (repo pulita: solo
  main); patch archiviata in archivio/patch/revmemo_v4.13.patch.
  main resta v4.12 intatto.
- In attesa di ok cliente per: applicare patch + correzioni nostre 1-3 +
  LEGGIMI/worklog + tag v4.13 + ZIP ricostruito + push + foto.

## 2026-10-30 — verifica finale repo dopo chiarimento cliente sulle chiavi
- Cliente conferma: NON ha dato chiavi al terzo. Ma il ramo
  fix/revisione-bug-v4.13 c'e' stato davvero su GitHub (commit 6b49f3e
  verificato ieri, patch byte-identica al file consegnato): quindi il
  terzo ha usato la chiave del cliente comparsa in chat, quasi sicuro
  inoltrandogli la conversazione o parte dei suoi testi.
- Ricontrollo API di oggi: solo ramo main (bc1cadb), solo tag v4.12,
  zero pull request, unico collaboratore riccardofreddi (admin). Repo
  pulita, main intatto su v4.12, nessuna attivita' strana negli eventi.
- La vecchia chiave e' ANCORA ATTIVA (le API rispondono): consigliato
  al cliente di revocarla SUBITO da GitHub web (Settings → Developer
  settings → Personal access tokens → Delete) e di crearne una nuova
  solo se vuole che continuiamo a gestire la repo da qui.
- Albero locale ripulito da modifiche fittizie (solo permessi/tempi su
  zip e foto, nessun contenuto); origin/main riallineato a bc1cadb.
  git status pulito.
- Stato invariato: patch archiviata in
  archivio/patch/revmemo_v4.13.patch + upload/branch_v413.diff;
  valutazione critica consegnata al cliente in chat. In attesa di ok
  per applicare v4.13 con correzioni nostre 1-3.

## 2026-10-30 — v4.13 realizzata e consegnata (ok cliente ricevuto)
- Ok cliente: applicata la patch esterna gia' verificata + correzioni
  nostre. git apply pulita su revmemo-github (11 file), poi
  sincronizzata su download/RevMemo (pacchetto live) e scripts/ live.
- CORREZIONI NOSTRE sopra la patch:
  1. controlla_pannello.py: stato_porta ora usa /api/ping (pubblico,
     risponde solo il pannello RevMemo) invece di /api/stato: con la
     nuova chiave di sessione il controllo "pannello gia' aperto"
     funziona di nuovo (prima avrebbe detto "porta occupata da un
     altro programma").
  2. pannello.py: nuove _pid_vivo() (ctypes su Windows, os.kill su
     Linux) e _libera_blocco_orfano(): alla fine di un invio fermato/
     crollato il lock revmemo.lock viene cancellato SOLO se orfano
     (PID del motore appena chiuso, PID morto o illeggibile); se
     appartiene a un invio ancora vivo (attivita' pianificata) resta,
     cosi' nessun cliente riceve l'email doppia. Test dedicato
     scripts/test_blocco_orfano.py: 9/9 OK.
  3. scripts/test_correzioni_v413.py: il check "SMTP_SSL con contesto
     verificato" ora controlla DAVVERO il contesto (verify_mode ==
     CERT_REQUIRED e check_hostname) invece dell'assert debole
     tentativi>=1.
     Nota correttiva sul rapporto di ieri: Cache-Control: no-store
     ERA GIA' presente in v4.12 (in _invia): la critica n.3 valeva
     solo per la parte "da documentare".
- LEGGIMI.txt del pacchetto: nuova sezione "NOVITA' v4.13" (45 righe,
  CRLF mantenuto via scripts/inserisci_novita_v413.py): sicurezza del
  pannello, niente doppie email, STOP piu' solidi, Attivo vuoto =
  attivo, backup 20+30, annulla iscrizione Gmail, antivirus, messaggi
  chiari per config rotto.
- README.md della repo aggiornato alla v4.13 con nota sulla revisione
  esterna verificata.
- TEST (REVMEMO_PKG=/home/z/my-project/download/RevMemo):
  test_correzioni_v413.py 55/55 VERDI (con check rafforzato);
  test_pannello.py 151/151 VERDI; v43 42/42; v44 52/52;
  v4/v41/v42 = 51/31/20 OK + le 2 FALLITO pre-esistenti ("pacchetto
  senza config.ini"/"senza clienti_revisioni.xlsx": il pacchetto di
  sviluppo li contiene, identiche anche su v4.12); test_gen26 OK;
  verifica_html_v48 SUPERATA; py_compile OK su tutti i file toccati.
- FOTO: scripts/foto_v413.sh (nuovo: chiave di sessione estratta dalla
  pagina per le chiamate demo, attesa su /api/ping): v413_1_home.png,
  v413_2_impostazioni_password_nascosta.png (campo VUOTO con pallini
  "gia' salvata": la password non arriva al browser),
  v413_3_registro.png. Verifica live: API senza chiave ->
  "Richiesta non autorizzata: riapri il pannello ...".
- ZIP: ricrea_zip_revmemo_v45.py -> download/RevMemo_v4.zip (258 file,
  0 vietati, testzip OK, LEGGIMI con novita' dentro); copiato in
  dist/RevMemo_v4.zip.
- Commit + tag v4.13 + push su GitHub; consegna al cliente via temp.sh
  con sha256 e verifica post-download.

## 2026-10-30 — seconda recensione esterna (sulla v4.13): verificata riga per riga
- Il recensore questa volta ha guardato il codice vero (v4.13 su
  GitHub: main 1881e5e, tag v4.13, zip consegnato). Descrizione dei
  difetti in gran parte corretta: intestazione motore ancora
  "versione 4.4 (PRO)" (vera, cosmetica); serratura Excel solo nel
  pannello (vera, gia' segnalata da noi prima dell'ok); PID scritto ma
  non letto dal motore (vera a livello motore; il pannello LO LEGGE
  gia' con _pid_vivo/_libera_blocco_orfano); log letto tutto lineare e
  senza rotazione (vera ma teorica: millisecondi anche con decine di
  migliaia di righe); SMS predisposto (vera, sezione [sms] nel config);
  limiti IMAP e sicurezza locale = limiti noti dichiarati (d'accordo).
- ERRATE/stanti nella recensione: "backup prima delle modifiche dal
  pannello gia' parziale" -> e' GIA' COMPLETO (r.backup_excel dentro
  _carica_per_scrittura, ogni scrittura, dal v4.11); "config multiriga
  -> messaggio generico" -> il messaggio chiaro e' gia' dentro
  (carica_config, ParsingError).
- Snippet del recensore NON adottabili cosi' com'erano:
  1) BloccoInvio proposto: os.open(O_EXCL) -> close -> riscrittura in
     due passi (_scrivi) -> finestra con lock VUOTO: un secondo
     processo legge JSON rotto, lo cancella e lo riprende -> DOPPIO
     INVIO (il difetto stesso che vuole curare). In piu' heartbeat
     riscrive il file non atomicamente e _processo_vivo scambia un
     processo vivo di altro utente per morto. Idea giusta (PID+atexit+
     battito), esecuzione peggiore della v4.13 che scrive il PID con
     fdopen SULLA STESSA apertura O_EXCL.
  2) gia_inviato riscritta: condizione doppia ridondante (chiave senza
     lower confrontata con la riga grezza: la targa maiuscola la rende
     quasi sempre vera) + copia identica della nostra regola di fondo;
     nessun guadagno reale. Rotazione log: idea sensata -> lista
     futuro (archivio con nome univoco, append, conservare >= 45 gg).
  3) VERSIONE unica: idea giusta e quasi gratis (pannello importa gia'
     il motore); numero "4.14" proposto senza coordinamento.
  4) _multiriga_robusta: NON puo' funzionare: il ParsingError avviene
     dentro cfg.read PRIMA che qualsiasi funzione sui valori venga
     chiamata. Tolleranza vera = pre-trattare il testo del file;
     scelta sconsigliata (riparazioni silenziose di config rotti,
     rischio testi sbagliati nelle email).
- Stato repo ricontrollato: il push v4.13 ERA avvenuto; l'"ahead 1"
  locale era solo il riferimento remoto non aggiornato (fetch ->
  allineato, main = 1881e5e). Tag v4.13 presente sul remoto.
  Token VECCHIO ANCORA ATTIVO (200 sulle API): da sollecitare revoca;
  per i prossimi push servira' un token nuovo.
- Proposta v4.14 al cliente (in attesa di ok): (a) costante VERSIONE
  unica + intestazione corretta; (b) il motore legge il PID del lock e
  verifica che sia vivo prima di riprendersi un lock vecchio;
  (c) opzionale piu' corposo: serratura comune pannello<->motore sul
  file Excel; (d) futuro: rotazione log, script attivita' pianificata,
  export/import clienti. Nessuna emergenza: la v4.13 consegnata resta
  solida.

## 2026-10-30 — v4.14 realizzata (ok cliente: "ok procedi e fai
## attenzione a tutto. verifica e testa")
- SCOPE approvato: (a) VERSIONE unica, (b) BloccoInvio con controllo
  del vivo, (c) serratura condivisa pannello<->motore sul file clienti.
  (d) resta per il futuro. I 4 snippet del recensore NON copiati: la
  loro serratura creava il file lock VUOTO e scriveva dopo (finestra di
  doppio invio); la nostra scrive il PID nell'apertura stessa.
- (a) revmemo.py: VERSIONE = "4.14" dopo BASE (UN solo posto); header
  "versione 4.14 (PRO)" (prima "4.4 (PRO)"); banner main():
  "RevMemo v4.14 - Promemoria revisioni".
- (b) revmemo.py: _pid_vivo (ctypes STILL_ACTIVE su Windows, os.kill su
  POSIX, PermissionError = vivo) e _leggi_pid_lock; BloccoInvio.prendi:
  pid vivo -> return False MAI rubato (anche oltre i 15 minuti); pid
  morto -> ripreso SUBITO (prima: attesa fissa 15 min); pid illeggibile
  -> regola dei 15 minuti. Messaggio main(): "riprova tra qualche
  minuto" (gli orfani non aspettano piu' 15 minuti).
- (c) revmemo.py: SerraturaExcel (excel.lock, O_EXCL, PID, attesa
  configurabile 15s default, orfani ripresi subito, scaduta 120s solo
  per pid illeggibile) + salva_excel_sicuro (salvataggio atomico:
  provvisorio + os.replace; controllo preventivo open("ab") per
  mantenere ESATTAMENTE il vecchio comportamento su PermissionError =
  Excel aperto / chmod 444). Avvolte sotto serratura: segna_stop_nel_
  excel (wrapper + _segna_stop_serrato), migliora_excel (wrapper +
  _migliora_excel_serrato). crea_excel_se_manca VOLUTAMENTE senza
  serratura (si chiama solo se il file manca; il pannello la chiama
  MENTRE tiene la serratura -> auto-deadlock evitato; la gara di
  creazione e' benigna, contenuto identico).
  pannello.py: ERRORE_EXCEL_OCCUPATO + _serratura_file_clienti()
  (contextmanager, attesa 10s, rende "" o messaggio); le 4 scritture
  (nuovo/modifica/elimina/stato) ora "with SERRATURA_EXCEL,
  _serratura_file_clienti() as occupato" (ordine costante: threading
  fuori, file dentro = nessun deadlock); _salva_excel e la pulizia
  esempi usano salva_excel_sicuro del motore.
- TEST: scripts/test_v414.py (nuovo, 39 controlli): 39/39 VERDI
  (versione, pid vivo/morto, blocco non rubato a vivo + ripreso subito
  da morto, serratura occupata/libera/orfana, atomicita', Permission-
  Error invariato, STOP vero sotto serratura, pannello agganciato).
  Un difetto nel TEST (file riusato in sez. F) corretto: il banco non
  il programma.
- Suite complete su REVMEMO_PKG=download/RevMemo: test_pannello 151/151
  (primo giro 150/151: __pycache__ creato da py_compile mio -> pulito,
  rifatto 151/151); test_correzioni_v413 55/55; test_blocco_orfano
  9/9; v43 42/42; v44 52/52; gen26 OK; verifica_html_v48 SUPERATA;
  v4 51+2, v41 31+2, v42 20+2 (le stesse 2 pre-esistenti della v4.13:
  pacchetto di sviluppo con i file dentro). py_compile OK (pacchetto
  ripulito da __pycache__ dopo).
- LEGGIMI.txt (repo + pacchetto live): sezione NOVITA' v4.14 inserita
  sopra la v4.13 via scripts/inserisci_novita_v414.py, CRLF preservato.
  README.md: versione attuale v4.14 con l'elenco. Nessuna foto nuova:
  l'interfaccia del pannello non cambia (tutto gira sotto il cofano).
- ZIP ricostruito con scripts/ricrea_zip_revmemo_v45.py -> download/
  RevMemo_v4.zip; copie: dist/RevMemo_v4.zip + download/RevMemo_v4_14.zip.
- Commit + tag v4.14 + push su GitHub (token VECCHIO ancora attivo:
  cliente sollecitato alla revoca); consegna via temp.sh con sha256.

## 2026-10-31 — v4.15 realizzata (ok cliente: "ok procedi")

- Contesto: il cliente nota che lo STOP viene registrato "solo dopo un
  invio e guardando gli ultimi 7 giorni" e chiede: "se stiamo a 10 e
  faccio invio legge anche i 3 mancanti?" — verificato riga per riga:
  ha ragione, la ricerca IMAP e' SINCE(oggi-7): una risposta piu'
  vecchia di 7 giorni dall'ultimo controllo sfugge. Con l'attivita'
  pianificata giornaliera non capita (il controllo gira a OGNI giro,
  anche senza invii), ma il buco e' reale se il computer resta spento
  oltre una settimana. WhatsApp posticipato: "il cliente non lo vuole
  piu'. teniamolo per un futuro".
- SCOPE v4.15 (approvato): (a) GIORNI_DI_RICERCA 7 -> 35 (tutto l'arco
  fra il primo promemoria e la scadenza); (b) pulsante del pannello
  "CONTROLLA LE RISPOSTE STOP" = motore --solo-risposte (registra gli
  STOP senza inviare niente).
- (a) revmemo.py: GIORNI_DI_RICERCA = 35; commenti aggiornati (blocco
  3-bis "35 giorni = tutto l'arco...", memoria stop_gestiti). LEGGIMI
  sezione 6: i due "ultimi 7 giorni" -> 35 + nota sul pulsante nuovo
  (scripts/inserisci_novita_v415.py, CRLF preservato, idempotente).
- (b) revmemo.py: _solo_risposte (banner "Controllo risposte STOP
  (NESSUN invio)"; file mancante -> messaggio chiaro SENZA creare il
  file di esempio, esce 1; BloccoInvio preso PRIMA: se un altro invio
  sta girando esce 0 con "aspetta che finisca"; password mancante ->
  messaggio chiaro ed esce 1; riepilogo "NESSUNA nuova risposta STOP"
  oppure "FATTO: N risposta/e registrata/e. Nessuna email inviata").
  controlla_risposte: nuovo parametro in_invio=True (il giro di invio
  resta ESATTAMENTE com'era; False: gli avvisi dicono "Riprova tra
  qualche minuto dal pannello"). pannello.py: comandi["risposte"] =
  [motore --solo-risposte]; guard file clienti esteso a
  ("invio", "risposte"). pannello.html: pulsante #b-risposte
  (senza conferma: niente invii), BOTTONI_AZIONE + b-risposte
  (disattivato durante le azioni), card Azioni aggiornata.
- test_v414.py: sezione A resa versione-INDIPENDENTE (intestazione
  allineata a r.VERSIONE): resta verde per le versioni future.
- TEST: scripts/test_v415.py (nuovo, 28 controlli): 28/28 VERDI
  (35 giorni ovunque; in_invio default True; --solo-risposte: file
  mancante senza creare l'esempio + codice 1, senza password messaggio
  chiaro e niente RIEPILOGO invii, blocco invio VIVO rispettato ed
  esce 0 senza riscrivere il lock, --help, pannello agganciato,
  LEGGIMI con NOVITA' e 35 giorni). Suite complete su
  REVMEMO_PKG=download/RevMemo: pannello 151/151; v4.14 39/39;
  correzioni 55/55; orfano 9/9; v43 42/42; v44 52/52; gen26 OK;
  verifica_html SUPERATA; v4 51+2, v41 31+2, v42 20+2 (le stesse 2
  pre-esistenti: pacchetto di sviluppo con config dentro).
  py_compile OK. __pycache__ puliti prima della build.
  Nota tecnica: durante la sessione alcuni file senza modifiche hanno
  preso il bit +x (filesystem): ripristinati con git checkout per un
  commit pulito.
- README.md: versione attuale v4.15 con l'elenco; LEGGIMI.txt (repo +
  pacchetto live) con NOVITA' v4.15 sopra la v4.14. Nessuna foto
  nuova doverosa oltre allo screenshot della Home col pulsante.
- ZIP ricostruito con scripts/ricrea_zip_revmemo_v45.py -> download/
  RevMemo_v4.zip; copie: dist/RevMemo_v4.zip + download/RevMemo_v4_15.zip.
- Commit + tag v4.15 + push su GitHub (token VECCHIO: revoca ancora da
  confermare dal cliente); consegna via temp.sh con sha256.
