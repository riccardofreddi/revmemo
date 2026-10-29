# -*- coding: utf-8 -*-
"""Test RevMemo v4.2 — scenario esatto dell'utente:
"ho risposto STOP da riccardofreddi6@gmail.com e riavviato il programma".

Il titolare RISPONDE (Re:/In-Reply-To) a un promemoria dalla PROPIA casella
scrivendo STOP: ora deve essere accettato. Le COPIE automatiche (promemoria
e riepilogo mandati a se stessi) devono restare ignorate.

Messaggi nella Gmail finta:
 1. copia del promemoria a se stessi (niente Re:, footer con STOP) -> IGNORATA
 2. RISPOSTA del titolare dalla sua casella con "STOP" scritto a mano
    (riga 8 = la sua riga cliente) -> DEVE mettere STOP
 3. RISPOSTA del titolare "VA BENE" (cita il footer) -> niente
 4. email NUOVA (non risposta) dal secondo indirizzo del centro con
    oggetto "STOP" (riga 12 = sua riga cliente) -> DEVE mettere STOP
 5. copia del riepilogo titolare a se stessi (body con tabella STOP) -> IGNORATA
 6. RISPOSTA dall'email contatti del centro (non presente nell'Excel)
    con STOP -> [INFO], nessuna riga da fermare
 7. RISPOSTA di un cliente vero con STOP -> DEVE mettere STOP
"""
import io
import os
import shutil
import sys
import tempfile
from contextlib import redirect_stdout

from email.message import EmailMessage

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


tmp = tempfile.mkdtemp(prefix="revmemo_v42_")
for f in ("revmemo.py", "template_email.html", "logo.png"):
    shutil.copy2(os.path.join(PACCHETTO, f), os.path.join(tmp, f))

CONFIG_TEST = """\
[smtp]
host = smtp.gmail.com
porta = 465
mittente = centro.prova.revmemo@gmail.com
password_app = aaaabbbbccccdddd
risposta_a =

[centro]
nome = Centro Revisioni Test
citta = Milano
telefono = 02 1234567
email_contatti = info@test.it
colore_tema = #1F3B66

[prodotto]
nome = RevMemo

[promemoria]
finestra_1mese = 26-30
finestra_10giorni = 8-10
finestra_1giorno = 1-2

[invii]
pausa_secondi = 0
notifica_titolare = SI
email_titolare = titolare@test.it
file_excel = clienti_revisioni.xlsx
file_log = reminders_log.txt

[sms]
attivo = NO
"""
with open(os.path.join(tmp, "config.ini"), "w", encoding="utf-8") as f:
    f.write(CONFIG_TEST)

sys.path.insert(0, tmp)
import revmemo  # noqa: E402

FOOTER = ("Non vuoi piu' ricevere promemoria? Rispondi a questa email "
          "con oggetto STOP.")


def msg(da, oggetto, corpo, in_risposta_a=None):
    m = EmailMessage()
    m["From"] = da
    m["To"] = "centro.prova.revmemo@gmail.com"
    m["Subject"] = oggetto
    if in_risposta_a:
        m["In-Reply-To"] = in_risposta_a
        m["References"] = in_risposta_a
    m.set_content(corpo)
    return m.as_bytes()


percorso_excel = os.path.join(tmp, "clienti_revisioni.xlsx")
revmemo.crea_excel_se_manca(percorso_excel)

# La MIA email del centro su una riga cliente (riga 8, Carlo Neri) e
# l'email del titolare su un'altra (riga 12, Franco Viola): come nell'errore
# di prova dell'utente, cosi' gli STOP del titolare hanno una riga da fermare.
from openpyxl import load_workbook  # noqa: E402
_wb = load_workbook(percorso_excel)
_wb["Clienti"].cell(row=8, column=3).value = "centro.prova.revmemo@gmail.com"
_wb["Clienti"].cell(row=12, column=3).value = "titolare@test.it"
_wb.save(percorso_excel)
_wb.close()

CITAZIONE = ("Il giorno mar 28 ott 2026 alle ore 09:00, Centro Revisioni Test "
             "<centro.prova.revmemo@gmail.com> ha scritto:\n"
             "> Gentile Carlo,\n> la revisione scade il 30/09/2026.\n"
             "> " + FOOTER + "\n")

MESSAGGI = {
    1: msg("Centro Revisioni Test <centro.prova.revmemo@gmail.com>",
           "Promemoria revisione AB123CD - scadenza 09/2026",
           "Gentile Cliente,\n\nla revisione scade il 30/09/2026.\n\n" + FOOTER + "\n"),
    2: msg("Centro Revisioni Test <centro.prova.revmemo@gmail.com>",
           "Re: Promemoria revisione AB123CD - scadenza 09/2026",
           "STOP\n\n" + CITAZIONE,
           in_risposta_a="<promemoria-ab123cd@centro.prova.revmemo.gmail.com>"),
    3: msg("Centro Revisioni Test <centro.prova.revmemo@gmail.com>",
           "Re: Promemoria revisione AB123CD - scadenza 09/2026",
           "VA BENE GRAZIE\n\n" + CITAZIONE,
           in_risposta_a="<promemoria-ab123cd@centro.prova.revmemo.gmail.com>"),
    4: msg("Titolare <titolare@test.it>",
           "STOP",
           "non voglio piu' ricevere nulla\n"),
    5: msg("Centro Revisioni Test <centro.prova.revmemo@gmail.com>",
           "RevMemo - riepilogo del 28/10",
           "Clienti che hanno risposto STOP:\n- Carlo Neri (OP321QR)\n"),
    6: msg("Segreteria <info@test.it>",
           "Re: Promemoria revisione AB123CD - scadenza 09/2026",
           "STOP\n\n" + CITAZIONE,
           in_risposta_a="<promemoria-ab123cd@centro.prova.revmemo.gmail.com>"),
    7: msg("Mario Rossi <mario.rossi@email.it>",
           "Re: Promemoria revisione AB123CD - scadenza 09/2026",
           "STOP per favore\n\n" + CITAZIONE,
           in_risposta_a="<promemoria-ab123cd@centro.prova.revmemo.gmail.com>"),
}


class FakeIMAP:
    istanza = None

    def __init__(self, host, port, timeout=None, **kw):
        self.marked = []
        self.fetch_full = []
        FakeIMAP.istanza = self

    def login(self, u, p):
        return ("OK", [b""])

    def select(self, box):
        return ("OK", [b"7"])

    def search(self, charset, *crit):
        return ("OK", [b"1 2 3 4 5 6 7"])

    def fetch(self, num, spec):
        raw = MESSAGGI[int(num)].replace(b"\n", b"\r\n")
        if "HEADER" in spec:
            cut = raw.find(b"\r\n\r\n")
            hdr = raw[:cut]
            return ("OK", [(num + b" (RFC822.HEADER {%d}" % len(hdr), hdr), b")"])
        self.fetch_full.append(num)
        return ("OK", [(num + b" (RFC822 {%d}" % len(raw), raw), b")"])

    def store(self, num, flags, flag):
        self.marked.append(num)
        return ("OK", [b""])

    def logout(self):
        return ("BYE", [b""])


imaplib_originale = revmemo.imaplib.IMAP4_SSL
revmemo.imaplib.IMAP4_SSL = FakeIMAP

print("== 1) risposta STOP dalla PROPRIA casella (caso dell'utente) ==")
buf = io.StringIO()
with redirect_stdout(buf):
    eventi = revmemo.controlla_risposte(percorso_excel, dry_run=False)
uscita = buf.getvalue()

check("3 eventi (titolare riga 8, titolare riga 12, cliente riga 5)",
      len(eventi) == 3, str(eventi))
check("evento 1: riga 8 (la MIA riga cliente)",
      eventi and eventi[0]["righe"] and eventi[0]["righe"][0][0] == 8
      and eventi[0]["email"] == "centro.prova.revmemo@gmail.com", str(eventi))
check("evento 2: riga 12 (riga del titolare)",
      eventi and eventi[1]["righe"] and eventi[1]["righe"][0][0] == 12
      and eventi[1]["email"] == "titolare@test.it", str(eventi))
check("evento 3: cliente mario riga 5",
      eventi and eventi[2]["righe"] and eventi[2]["righe"][0][0] == 5
      and eventi[2]["email"] == "mario.rossi@email.it", str(eventi))

ws = load_workbook(percorso_excel)["Clienti"]
check("riga 8 = STOP", ws.cell(row=8, column=8).value == "STOP")
check("riga 12 = STOP", ws.cell(row=12, column=8).value == "STOP")
check("riga 5 = STOP", ws.cell(row=5, column=8).value == "STOP")
check("riga 6 resta SÌ (nessun falso STOP)",
      ws.cell(row=6, column=8).value == "SÌ")

check("copia promemoria ignorata (contatore = 2)",
      "(ignorate 2 email inviate dal centro stesso" in uscita, uscita)
check("[INFO] per STOP di info@test.it (email non nell'Excel)",
      "[INFO] STOP letto da info@test.it" in uscita, uscita)
check("mario segnalato con [STOP]", "[STOP] mario.rossi@email.it" in uscita)
check("marcate lette solo le 3 email con STOP +INFO = 4",
      FakeIMAP.istanza.marked == [b"2", b"4", b"7"], str(FakeIMAP.istanza.marked))

print("== 2) rilancio: niente doppioni, copie ancora ignorate ==")
buf = io.StringIO()
with redirect_stdout(buf):
    eventi2 = revmemo.controlla_risposte(percorso_excel, dry_run=False)
uscita2 = buf.getvalue()
check("rilancio: zero nuovi eventi", eventi2 == [], str(eventi2))
check("rilancio: contatore copie ancora presente",
      "(ignorate 2 email inviate dal centro stesso" in uscita2)

print("== 3) main() completo in PROVA con Gmail finta ==")
sys.argv = ["revmemo.py", "--dry-run"]
buf = io.StringIO()
with redirect_stdout(buf):
    codice = revmemo.main()
uscita3 = buf.getvalue()
check("main ritorna 0", codice == 0)
check("MODALITA' PROVA", "MODALITA' PROVA" in uscita3)
check("righe in STOP escluse dall'invio", "ha chiesto STOP" in uscita3)
check("riepilogo presente", "RIEPILOGO:" in uscita3)

revmemo.imaplib.IMAP4_SSL = imaplib_originale

print("== 4) il pacchetto vero resta pulito ==")
for f in ("config.ini", "clienti_revisioni.xlsx", "reminders_log.txt", "__pycache__"):
    check(f"pacchetto senza {f}", not os.path.exists(os.path.join(PACCHETTO, f)))

print()
print("=" * 50)
print(f"RISULTATO: {ok_count} OK, {fail_count} FALLITI")
print("=" * 50)
sys.exit(1 if fail_count else 0)
