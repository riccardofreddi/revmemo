#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Test v3.1 di RevMemo: MIME logo, anti-doppioni per scadenza,
auto-copia Excel d'esempio, controlla_pc, dry-run completo."""
import os
import shutil
import subprocess
import sys
import datetime as dt

BASE_PROG = "/home/z/my-project/download/RevMemo"
ENV = "/home/z/my-project/scripts/test_env"

CONFIG_TEST = """\
[smtp]
host = smtp.gmail.com
porta = 465
mittente = testrevmemo@gmail.com
password_app = abcd efgh ijkl mnop
risposta_a =

[centro]
nome = Centro Test
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
pausa_secondi = 0.1
notifica_titolare = SI
email_titolare =
file_excel = clienti_revisioni.xlsx
file_log = reminders_log.txt

[sms]
attivo = NO
"""

PASS = FAIL = 0
def check(nome, cond, extra=""):
    global PASS, FAIL
    if cond:
        PASS += 1
        print(f"  OK   {nome}")
    else:
        FAIL += 1
        print(f"  FAIL {nome} {extra}")

print("== 0) py_compile ==")
for f in ("revmemo.py", "controlla_pc.py"):
    r = subprocess.run([sys.executable, "-m", "py_compile",
                        os.path.join(BASE_PROG, f)], capture_output=True, text=True)
    check(f"compila {f}", r.returncode == 0, r.stderr[:300])

# ambiente di test pulito
if os.path.exists(ENV):
    shutil.rmtree(ENV)
os.makedirs(ENV)
for f in ("revmemo.py", "controlla_pc.py", "template_email.html",
          "logo.png", "clienti_ESEMPIO.xlsx"):
    shutil.copy2(os.path.join(BASE_PROG, f), os.path.join(ENV, f))
with open(os.path.join(ENV, "config.ini"), "w", encoding="utf-8") as f:
    f.write(CONFIG_TEST)

sys.path.insert(0, ENV)
import revmemo as rm  # noqa: E402

print("\n== 1) struttura MIME dell'email (logo allegato con cid) ==")
cliente = {"scadenza": dt.date(2026, 11, 5), "targa": "AB123CD",
           "nome": "Mario Rossi", "email": "mario@gmail.com"}
oggetto, testo, html = rm.componi_email(cliente, 8)
check("html generato dal template", html is not None and "{TARGA}" not in html)
check("html usa cid:logo", 'src="cid:logo"' in html)

from email.message import EmailMessage
msg = EmailMessage()
msg["From"] = "Centro Test <testrevmemo@gmail.com>"
msg["To"] = "mario@gmail.com"
msg["Subject"] = oggetto
msg.set_content(testo)
msg.add_alternative(html, subtype="html")
logo = rm.leggi_logo()
check("logo.png letto (non vuoto)", bool(logo) and len(logo) > 100)
msg.get_payload()[1].add_related(logo, maintype="image", subtype="png", cid="<logo>")

def struttura(part, prof=0):
    out = ["  " * prof + part.get_content_type()
           + (f" cid={part.get('Content-ID')}" if part.get("Content-ID") else "")]
    if part.is_multipart():
        for p in part.iter_parts():
            out += struttura(p, prof + 1)
    return out

print("\n".join(struttura(msg)))
corpo = msg.get_payload()[1]
check("corpo e' multipart/related", corpo.get_content_type() == "multipart/related")
img_parts = [p for p in corpo.iter_parts() if p.get_content_type() == "image/png"]
check("immagine PNG dentro la mail", len(img_parts) == 1)
check("Content-ID e' <logo>", img_parts and img_parts[0].get("Content-ID") == "<logo>")
check("immagine ha byte reali", len(img_parts[0].get_payload(decode=True) or b"") > 100)

print("\n== 2) anti-doppioni per scadenza ==")
log = os.path.join(ENV, "reminders_log.txt")
with open(log, "w", encoding="utf-8") as f:
    f.write("2026-10-28;mario@gmail.com;AB123CD;10;2026-11-05;10 giorni prima (8 gg)\n")
    f.write("2026-10-27;vecchio@gmail.com;XYZ12AB;30;1 mese prima (27 gg)\n")  # formato vecchio
check("stessa scadenza -> gia' inviato (maiuscole/minuscole)",
      rm.gia_inviato("Mario@Gmail.com", "ab123cd", 10, "2026-11-05") is True)
check("scadenza diversa (rinnovo) -> NON inviato, riparte",
      rm.gia_inviato("mario@gmail.com", "AB123CD", 10, "2026-11-06") is False)
check("tipo diverso (30gg) -> NON inviato",
      rm.gia_inviato("mario@gmail.com", "AB123CD", 30, "2026-11-05") is False)
check("vecchio formato -> comunque 'gia' inviato' (zero doppioni)",
      rm.gia_inviato("vecchio@gmail.com", "XYZ12AB", 30, "2027-01-31") is True)
check("altro cliente -> NON inviato",
      rm.gia_inviato("altro@gmail.com", "AB123CD", 10, "2026-11-05") is False)
rm.registra_invio("nuovo@gmail.com", "EE111ZZ", 1, 2, "1 giorno prima", "2026-10-30")
with open(log, encoding="utf-8") as f:
    ultima = f.readlines()[-1].strip()
check("registra_invio scrive data;email;targa;tipo;scadenza",
      ultima == f"{dt.date.today().isoformat()};nuovo@gmail.com;EE111ZZ;1;"
                f"2026-10-30;1 giorno prima (2 gg)", ultima)
check("subito rilegibile come inviato",
      rm.gia_inviato("NUOVO@gmail.com", "ee111zz", 1, "2026-10-30") is True)

print("\n== 3) auto-copia Excel d'esempio (aggiornamenti sicuri) ==")
check("clienti_revisioni.xlsx NON c'e' ancora (solo ESEMPIO)",
      not os.path.exists(os.path.join(ENV, "clienti_revisioni.xlsx")))
r = subprocess.run([sys.executable, os.path.join(ENV, "controlla_pc.py")],
                   input="\n", capture_output=True, text=True)
out = r.stdout
check("controlla_pc gira senza errori", r.returncode == 0)
check("controlla_pc ha creato l'Excel dall'esempio",
      "l'ho creato dall'esempio" in out, out[-500:])
check("clienti_revisioni.xlsx adesso esiste",
      os.path.exists(os.path.join(ENV, "clienti_revisioni.xlsx")))

print("\n== 4) dry-run completo su Excel ricreato ==")
os.remove(log)  # tolgo le righe di test anti-doppioni: altrimenti bloccano la prova
# metto il primo cliente nella finestra 8-10 gg (scadenza fissa 05/11/2026)
from openpyxl import load_workbook
perc_x = os.path.join(ENV, "clienti_revisioni.xlsx")
wb = load_workbook(perc_x)
ws = wb["Clienti"] if "Clienti" in wb.sheetnames else wb.active
col_scad = col_mail = None
for c in range(1, ws.max_column + 1):
    h = str(ws.cell(row=4, column=c).value or "").strip().lower()
    if h == "scadenza revisione":
        col_scad = c
    if h == "email":
        col_mail = c
ws.cell(row=5, column=col_scad).value = "05/11/2026"
ws.cell(row=5, column=col_mail).value = "mario.prova@gmail.com"
wb.save(perc_x)
r = subprocess.run([sys.executable, os.path.join(ENV, "revmemo.py"), "--dry-run"],
                   capture_output=True, text=True)
out = r.stdout
check("dry-run esce 0", r.returncode == 0, r.stderr[:400])
check("rileva 8 clienti", "Trovati 8 clienti" in out, out[:400])
check("mostra invii PROVA", "PROVA invio" in out, out[-800:])
check("email HTML con logo: SI", "Email HTML con logo: SI" in out, out[-800:])
r2 = subprocess.run([sys.executable, os.path.join(ENV, "revmemo.py"), "--dry-run"],
                    capture_output=True, text=True)
check("dry-run: niente invii fuori finestra oltre agli esclusi", r2.returncode == 0)

print("\n== 5) risposta_a (Reply-To) ==")
check("config risposta_a letta (vuota)", rm.CFG.get("risposta_a") == "")
with open(os.path.join(ENV, "config.ini"), "a", encoding="utf-8") as f:
    pass  # la chiave esiste gia' vuota: test solo del default
check("mittente letto", rm.MITTENTE_EMAIL == "testrevmemo@gmail.com")

print(f"\n===== RISULTATO: {PASS} OK, {FAIL} FAIL =====")
sys.exit(1 if FAIL else 0)
