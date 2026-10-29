# -*- coding: utf-8 -*-
"""Test RevMemo v4.4 — "personalizzazione totale: ogni centro vuole il
SUO messaggio" (config.ini sezione [messaggio] + nuovi dati centro).

Verifica:
 1. aggiorna_config: un config.ini vecchio (CRLF) viene arricchito da
    solo con la sezione [messaggio] e i campi indirizzo/sito_web/orari
    del [centro], SENZA toccare i valori; rilanciato = nessuna modifica
 2. componi_email con i testi predefiniti: oggetto, saluto col NOME del
    cliente (fallback "Cliente"), testo senza ** nel plain, footer con
    contatti e STOP, segnaposto senza residui
 3. vecchio template (senza {TESTO_HTML}): sostituzione classica che
    continua a funzionare
 4. aggiorna_template: vecchio -> nuovo con backup (template_email_
    vecchio.html), idempotente, e ricrea il file se manca
 5. componi_email col template nuovo: **grassetto** solo nell'HTML,
    paragrafi da riga vuota, escape HTML dei nomi con & e <, footer
    con indirizzo/orari/riga extra/disiscrizione personalizzata
 6. config personalizzato (altro tmp): oggetto speciale per tipo,
    segnaposto {SITO_WEB}/{ORARI}/{INDIRIZZO} sostituiti
 7. config.ini rotto (riga senza spazio dopo testo =): messaggio
    chiaro e uscita 1 (subprocess)
 8. CLI --dry-run completa in una cartella vergine (subprocess)
"""
import io
import os
import shutil
import subprocess
import sys
import tempfile
from contextlib import redirect_stdout

import datetime as dt

PACCHETTO = os.environ.get("REVMEMO_PKG", "/home/z/my-project/download/RevMemo")

ok_count = 0
fail_count = 0


def check(nome, condizione, extra=""):
    global ok_count, fail_count
    if condizione:
        ok_count += 1
        print(f"  [OK]   {nome}")
    else:
        fail_count += 1
        print(f"  [FALLITO] {nome} {extra}")


OLD_TEMPLATE = """<html><body style="background:{COLORE};">
<p>Gentile Cliente,</p>
<p>la revisione {TARGA} scade il {SCADENZA}. Manca {GIORNI_LABEL} ({TRAFRA}).</p>
<p>{NOME_CENTRO} - {CITTA} - Tel {TELEFONO}{CITTA_FOOTER}{EMAIL_FOOTER}</p>
<span>{PRODOTTO} {ANNO}{LOGO_IMG}</span>
</body></html>
"""

CONFIG_VECCHIO_CRLF = (
    "[smtp]\r\n"
    "host = smtp.gmail.com\r\n"
    "porta = 465\r\n"
    "mittente = centro.prova.v44@gmail.com\r\n"
    "password_app = aaaabbbbccccdddd\r\n"
    "risposta_a =\r\n"
    "\r\n"
    "[centro]\r\n"
    "nome = Centro Vecchio\r\n"
    "citta = Torino\r\n"
    "telefono = 011 555666\r\n"
    "email_contatti = vecchio@test.it\r\n"
    "colore_tema = #C8102E\r\n"
    "\r\n"
    "[prodotto]\r\n"
    "nome = PromemoriaCentro\r\n"
    "\r\n"
    "[promemoria]\r\n"
    "finestra_1mese = 26-30\r\n"
    "finestra_10giorni = 8-10\r\n"
    "finestra_1giorno = 1-2\r\n"
    "\r\n"
    "[invii]\r\n"
    "pausa_secondi = 0\r\n"
    "notifica_titolare = SI\r\n"
    "email_titolare =\r\n"
    "file_excel = clienti_revisioni.xlsx\r\n"
    "file_log = reminders_log.txt\r\n"
    "\r\n"
    "[sms]\r\n"
    "attivo = NO\r\n"
)

tmp = tempfile.mkdtemp(prefix="revmemo_v44_")
for f in ("revmemo.py", "logo.png"):
    shutil.copy2(os.path.join(PACCHETTO, f), os.path.join(tmp, f))
with open(os.path.join(tmp, "config.ini"), "w", encoding="utf-8", newline="") as f:
    f.write(CONFIG_VECCHIO_CRLF)
with open(os.path.join(tmp, "template_email.html"), "w", encoding="utf-8",
          newline="") as f:
    f.write(OLD_TEMPLATE)

sys.path.insert(0, tmp)
buf = io.StringIO()
with redirect_stdout(buf):
    import revmemo  # noqa: E402
uscita_import = buf.getvalue()

print("\n== 1) aggiorna_config: config vecchio arricchito da solo ==")
check("stampata l'aggiunta della sezione [messaggio]",
      "aggiunta la sezione [messaggio]" in uscita_import, repr(uscita_import))
check("stampata l'aggiunta dei campi del [centro]",
      "indirizzo, sito_web e orari" in uscita_import)
percorso_config = os.path.join(tmp, "config.ini")
with open(percorso_config, "r", encoding="utf-8", newline="") as f:
    cfg_testo = f.read()
check("config ora contiene [messaggio]", "[messaggio]" in cfg_testo)
check("CRLF preservato", "\r\n" in cfg_testo)
check("nuova chiave con CRLF", "\r\nindirizzo =" in cfg_testo)
check("mittente NON toccato", "mittente = centro.prova.v44@gmail.com" in cfg_testo)
check("nome centro NON toccato", "nome = Centro Vecchio" in cfg_testo)
check("prodotto NON toccato", "nome = PromemoriaCentro" in cfg_testo)
check("indirizzo finisce DENTRO [centro] (prima di [prodotto])",
      cfg_testo.find("indirizzo =") < cfg_testo.find("[prodotto]"))
check("sezione [messaggio] letta: oggetto predefinito attivo",
      revmemo.CFG["msg_oggetto"] ==
      "[{NOME_CENTRO}] Promemoria revisione {TARGA} - scade {TRAFRA}")
check("nuovi campi vuoti ma presenti",
      revmemo.CFG["indirizzo"] == "" and revmemo.CFG["sito_web"] == ""
      and revmemo.CFG["orari"] == "")

buf2 = io.StringIO()
with redirect_stdout(buf2):
    revmemo.aggiorna_config()
with open(percorso_config, "r", encoding="utf-8", newline="") as f:
    cfg_testo2 = f.read()
check("rilancio di aggiorna_config: nessun doppione, file identico",
      cfg_testo2 == cfg_testo and buf2.getvalue() == "")

print("\n== 2) componi_email con i testi predefiniti (template VECCHIO) ==")
cliente = {"riga": 5, "nome": "Mario Rossi", "email": "mario@x.it",
           "telefono": "347", "targa": "AB123CD",
           "scadenza": dt.date(2026, 9, 30), "da_mese": True,
           "fatta": False, "attivo": True, "attivo_testo": "SÌ"}
ogg, testo, html = revmemo.componi_email(cliente, 29, 30)
check("oggetto generale con centro e targa",
      ogg == "[Centro Vecchio] Promemoria revisione AB123CD - scade tra 29 giorni", ogg)
check("saluto col NOME del cliente", testo.startswith("Gentile Mario Rossi,"))
check("vecchio template usato cosi' com'e'",
      html is not None and "Gentile Cliente," in html and "Manca 29 GIORNI" in html)
check("vecchio template: segnaposto riempiti",
      "AB123CD" in html and "{TARGA}" not in html and "{GIORNI_LABEL}" not in html)
check("testo semplice senza asterischi **", "**" not in testo)
check("STOP nel footer del testo semplice",
      "con oggetto STOP" in testo and "---" in testo)
check("white label nel testo", "Promemoria inviato con PromemoriaCentro" in testo)

print("\n== 3) aggiorna_template: backup + nuova versione ==")
buf3 = io.StringIO()
with redirect_stdout(buf3):
    revmemo.aggiorna_template()
check("annunciato l'aggiornamento", "nuova versione" in buf3.getvalue())
percorso_backup = os.path.join(tmp, "template_email_vecchio.html")
check("backup creato", os.path.exists(percorso_backup))
with open(percorso_backup, "r", encoding="utf-8") as f:
    backup_testo = f.read()
check("backup = contenuto vecchio", "Gentile Cliente," in backup_testo
      and "{TESTO_HTML}" not in backup_testo)
with open(os.path.join(tmp, "template_email.html"), "r", encoding="utf-8") as f:
    nuovo_testo = f.read()
check("template nuovo ha {TESTO_HTML}", "{TESTO_HTML}" in nuovo_testo)
check("template nuovo ha il piede esteso",
      "{INDIRIZZO_FOOTER}" in nuovo_testo and "{ORARI_FOOTER}" in nuovo_testo)

buf4 = io.StringIO()
with redirect_stdout(buf4):
    revmemo.aggiorna_template()
check("seconda chiamata: nessuna azione (idempotente)",
      buf4.getvalue() == "" and os.path.exists(percorso_backup))
with open(percorso_backup, "r", encoding="utf-8") as f:
    check("backup intatto dopo il secondo giro", "Gentile Cliente," in f.read())

print("\n== 4) componi_email col template NUOVO ==")
ogg, testo, html = revmemo.componi_email(cliente, 29, 30)
check("grassetto **{TARGA}** -> <b>AB123CD</b>", "<b>AB123CD</b>" in html)
check("grassetto **entro il ...** -> <b>", "<b>entro il 30/09/2026</b>" in html)
check("saluto nel HTML col nome", "Gentile Mario Rossi," in html)
check("firma nel HTML", "Grazie della fiducia, Centro Vecchio" in html)
check("paragrafi <p> generati dal testo", html.count("<p style=") >= 2)
check("nessun segnaposto residuo",
      not any(x in html for x in ("{SALUTO_HTML}", "{TESTO_HTML}",
                                  "{FIRMA_HTML}", "{NOME}", "{TARGA}",
                                  "{TESTO_DISISCRIZIONE}", "{GIORNI}")))
check("footer STOP dal predefinito",
      "Rispondi a questa email con oggetto STOP" in html)
check("footer senza orari/indirizzo se vuoti",
      "Orari:" not in html and "Via " not in html)
check("oggetti speciali vuoti -> oggetto generale per tipo 1",
      revmemo.componi_email(cliente, 1, 1)[0] ==
      "[Centro Vecchio] Promemoria revisione AB123CD - scade tra 1 giorno")

print("\n== 5) escape HTML e fallback 'Cliente' ==")
cliente_strano = dict(cliente, nome="Rossi & Figli <srl>")
_, testo_str, html_str = revmemo.componi_email(cliente_strano, 29, 30)
check("nome con & e < scappato nell'HTML",
      "Rossi &amp; Figli &lt;srl&gt;" in html_str)
check("nome intatto nel testo semplice", "Rossi & Figli <srl>" in testo_str)
cliente_senzanome = dict(cliente, nome="")
_, testo_sn, _ = revmemo.componi_email(cliente_senzanome, 29, 30)
check("senza nome -> 'Gentile Cliente,'", testo_sn.startswith("Gentile Cliente,"))

print("\n== 6) template mancante: testo semplice e ricrea da solo ==")
os.remove(os.path.join(tmp, "template_email.html"))
_, testo_nt, html_nt = revmemo.componi_email(cliente, 29, 30)
check("html None senza template", html_nt is None)
check("testo semplice comunque completo",
      "Gentile Mario Rossi," in testo_nt and "STOP" in testo_nt)
buf5 = io.StringIO()
with redirect_stdout(buf5):
    revmemo.aggiorna_template()
check("template ricreato dal modello interno",
      "l'ho ricreato" in buf5.getvalue()
      and os.path.exists(os.path.join(tmp, "template_email.html")))

# ---------------------------------------------------------------------
# 7) config personalizzato in un'altra cartella (subprocess)
print("\n== 7) config PERSONALIZZATO: oggetti speciali, grassetto, footer ==")
tmp2 = tempfile.mkdtemp(prefix="revmemo_v44b_")
for f in ("revmemo.py", "logo.png", "template_email.html"):
    shutil.copy2(os.path.join(PACCHETTO, f), os.path.join(tmp2, f))
CONFIG_CUSTOM = """\
[smtp]
host = smtp.gmail.com
porta = 465
mittente = centro2@test.it
password_app = aaaabbbbccccdddd
risposta_a =

[centro]
nome = Centro Nuovo
citta = Bologna
indirizzo = Via Indipendenza 10, 40121 Bologna (BO)
telefono = 051 777888
email_contatti = info@nuovo.it
sito_web = www.nuovo.it
orari = Lun-Ven 8:00-13:00 e 14:30-18:00
colore_tema = #1B7D46

[messaggio]
oggetto = Revisione {TARGA}: mancano {GIORNI} giorni - {NOME_CENTRO}
oggetto_1mese =
oggetto_10giorni =
oggetto_1giorno = URGENTE {NOME}: revisione {TARGA} domani!
saluto = Ciao {NOME},
testo =
    la tua revisione **{TARGA}** scade il {SCADENZA}.

    Vieni a trovarci: {SITO_WEB} - {INDIRIZZO}
    Oppure chiama il {TELEFONO}.
firma = A presto,
    {NOME_CENTRO}
riga_extra_piede = Parcheggio gratuito per i clienti
testo_disiscrizione = Vuoi smettere? Scrivi STOP in risposta

[prodotto]
nome = RevMemo

[promemoria]
finestra_1mese = 26-30
finestra_10giorni = 8-10
finestra_1giorno = 1-2

[invii]
pausa_secondi = 0
notifica_titolare = SI
email_titolare =
file_excel = clienti_revisioni.xlsx
file_log = reminders_log.txt

[sms]
attivo = NO
"""
with open(os.path.join(tmp2, "config.ini"), "w", encoding="utf-8") as f:
    f.write(CONFIG_CUSTOM)

runner2 = os.path.join(tmp2, "_runner.py")
with open(runner2, "w", encoding="utf-8") as f:
    f.write(f"""\
import sys, datetime
sys.path.insert(0, {tmp2!r})
import revmemo
assert revmemo.CFG["msg_oggetto"] == "Revisione {{TARGA}}: mancano {{GIORNI}} giorni - {{NOME_CENTRO}}", revmemo.CFG["msg_oggetto"]
c = {{"nome": "Luigi", "targa": "GH456JK", "scadenza": datetime.date(2026, 9, 30)}}
ogg, testo, html = revmemo.componi_email(c, 29, 30)
assert ogg == "Revisione GH456JK: mancano 29 giorni - Centro Nuovo", ogg
ogg1, _, _ = revmemo.componi_email(c, 1, 1)
assert ogg1 == "URGENTE Luigi: revisione GH456JK domani!", ogg1
assert "Ciao Luigi," in html and "Ciao Luigi," in testo
assert "<b>GH456JK</b>" in html and "**" not in testo
p = html.count("<p style=")
assert p == 5, f"paragrafi attesi 5 (saluto+2 testi+firma+1 piede), trovati {{p}}"
assert "www.nuovo.it" in html and "Via Indipendenza 10" in html
assert "Orari: Lun-Ven 8:00-13:00 e 14:30-18:00<br>" in html
assert "Parcheggio gratuito per i clienti<br>" in html
assert "Vuoi smettere? Scrivi STOP in risposta" in html
assert "A presto," in html and "Centro Nuovo" in html
assert "Via Indipendenza 10" in testo and "Orari: Lun-Ven" in testo
assert "Tel. 051 777888 | info@nuovo.it | www.nuovo.it" in testo
assert "{{TESTO_HTML}}" not in html and "{{NOME}}" not in html
assert revmemo.CFG["msg_oggetti_tipo"][30] is None  # oggetto_1mese vuoto
print("CUSTOM_OK")
""")
r2 = subprocess.run([sys.executable, runner2], capture_output=True, text=True)
check("config personalizzato: tutti i controlli HTML/testo/oggetti",
      r2.returncode == 0 and "CUSTOM_OK" in r2.stdout,
      (r2.stderr or r2.stdout)[-400:])

# ---------------------------------------------------------------------
# 8) config rotto: riga senza spazio dopo testo =
print("\n== 8) config rotto: messaggio chiaro e uscita 1 ==")
tmp3 = tempfile.mkdtemp(prefix="revmemo_v44c_")
shutil.copy2(os.path.join(PACCHETTO, "revmemo.py"), os.path.join(tmp3, "revmemo.py"))
with open(os.path.join(tmp3, "config.ini"), "w", encoding="utf-8") as f:
    f.write("[smtp]\nmittente = x@y.it\n\n[messaggio]\ntesto =\nriga senza spazio iniziale\n")
r3 = subprocess.run([sys.executable, os.path.join(tmp3, "revmemo.py"), "--dry-run"],
                    capture_output=True, text=True)
check("uscita 1 con config rotto", r3.returncode == 1, str(r3.returncode))
check("spiegazione chiara (spazio iniziale)",
      "QUALCHE SPAZIO" in r3.stdout and "testo =" in r3.stdout,
      (r3.stderr or r3.stdout)[-300:])

# ---------------------------------------------------------------------
# 9) CLI --dry-run completa in cartella vergine (con un cliente in finestra)
print("\n== 9) CLI --dry-run in cartella vergine (tutto automatico) ==")
tmp4 = tempfile.mkdtemp(prefix="revmemo_v44d_")
for f in ("revmemo.py", "logo.png", "template_email.html"):
    shutil.copy2(os.path.join(PACCHETTO, f), os.path.join(tmp4, f))

# Excel con UN cliente la cui scadenza cade nella finestra 26-30 giorni
# (i clienti di esempio creati da solo NON sono nella finestra con la
# data di oggi, quindi preparo il file a mano come farebbe il centro)
from openpyxl import Workbook  # noqa: E402
wb4 = Workbook()
ws4 = wb4.active
ws4.title = "Clienti"
ws4["B4"] = "Nome Cliente"
ws4["C4"] = "Email"
ws4["D4"] = "Telefono"
ws4["E4"] = "Targa"
ws4["F4"] = "Scadenza Revisione"
ws4["G4"] = "Revisione Fatta"
ws4["H4"] = "Attivo"
scad_fine = (dt.date.today() + dt.timedelta(days=29)).strftime("%d/%m/%Y")
righe4 = [("Mario Prova", "mario.prova@test.it", "347 0000000",
           "AB123CD", scad_fine, "NO", "SÌ")]
for i, riga in enumerate(righe4, start=5):
    for j, v in enumerate(riga, start=2):
        ws4.cell(row=i, column=j, value=v)
wb4.save(os.path.join(tmp4, "clienti_revisioni.xlsx"))

r4 = subprocess.run([sys.executable, os.path.join(tmp4, "revmemo.py"), "--dry-run"],
                    capture_output=True, text=True)
check("dry-run: uscita 0", r4.returncode == 0, (r4.stderr or r4.stdout)[-400:])
check("dry-run: nessun traceback", "Traceback" not in r4.stderr)
check("dry-run: banner prova", "MODALITA' PROVA" in r4.stdout)
check("dry-run: oggetto di esempio col centro predefinito",
      "[Centro Revisioni Esempio] Promemoria revisione AB123CD - scade tra 29 giorni"
      in r4.stdout, r4.stdout[-500:])
check("dry-run: segnalato il messaggio da config.ini",
      "config.ini sezione [messaggio]" in r4.stdout)
check("dry-run: niente 'config.ini aggiornato' (era gia' completo)",
      "config.ini aggiornato" not in r4.stdout)
check("dry-run: HTML con logo", "Email HTML con logo: SI" in r4.stdout)
check("dry-run: password mancante non blocca la prova",
      "password non ancora inserita" in r4.stdout)

print("\n" + "=" * 60)
print(f"RISULTATO v4.4: {ok_count} OK, {fail_count} FALLITI")
print("=" * 60)
sys.exit(1 if fail_count else 0)
