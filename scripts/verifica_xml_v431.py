# -*- coding: utf-8 -*-
"""v4.3.1 — verifica XML del file Excel generato:
solo celle G/H con 5 regole, NESSUNA tinta di riga, menu SÌ,NO senza STOP."""
import os
import re
import shutil
import sys
import zipfile

CONFIG_MINIMO = """\
[smtp]
host = smtp.gmail.com
porta = 465
mittente =
password_app =

[centro]
nome = X
citta = Y
telefono =
email_contatti =
colore_tema = #1F3B66

[prodotto]
nome = RevMemo

[promemoria]
finestra_1mese = 26-30
finestra_10giorni = 8-10
finestra_1giorno = 1-2

[invii]
pausa_secondi = 0
notifica_titolare = NO
email_titolare =
file_excel = clienti_revisioni.xlsx
file_log = reminders_log.txt

[sms]
attivo = NO
"""

tmp = "/tmp/revdemo/pkg3"
shutil.rmtree(tmp, ignore_errors=True)
os.makedirs(tmp, exist_ok=True)
for f in ("revmemo.py", "template_email.html", "logo.png"):
    shutil.copy2("/home/z/my-project/download/RevMemo/" + f, tmp + "/" + f)
with open(tmp + "/config.ini", "w", encoding="utf-8") as fh:
    fh.write(CONFIG_MINIMO)

sys.path.insert(0, tmp)
import revmemo  # noqa: E402

percorso = tmp + "/clienti_revisioni.xlsx"
revmemo.crea_excel_se_manca(percorso)

z = zipfile.ZipFile(percorso)
xml = z.read("xl/worksheets/sheet1.xml").decode("utf-8")

intervalli = re.findall(r'<conditionalFormatting sqref="([^"]+)">', xml)
print("Intervalli formattazione condizionale:", intervalli)
assert intervalli == ["G5:G300", "H5:H300"], f"attesi solo G5/H5, trovati {intervalli}"

n_regole = len(re.findall(r"<cfRule ", xml))
print("Regole totali:", n_regole)
assert n_regole == 5, f"attese 5, trovate {n_regole}"

assert "B5:H300" not in xml, "tinta di riga presente!"
print("Nessuna tinta di riga: OK")

# il menu vecchio SBAGLIATO ("SÌ,NO,STOP") non deve esserci in nessuna forma
assert "NO,STOP" not in xml, "STOP ancora nel menu a tendina!"
m = re.search(r"<dataValidation[^>]*>\s*<formula1>([^<]+)</formula1>", xml)
menu = m.group(1) if m else "??"
print("Menu a tendina nel XML:", menu)
assert "STOP" not in menu, "STOP ancora nel menu!"
assert "SI" in menu.upper().replace("S&#204;", "SI"), "menu non contiene SI"

# scadenza in formato testo: controllo su una cella
stili = z.read("xl/styles.xml").decode("utf-8")
assert "C00000" in stili, "rosso scuro STOP assente dagli stili"
print("Rosso scuro STOP (C00000) negli stili: OK")
print("\nXML OK: solo celle G5:G300 + H5:H300, 5 regole, menu SÌ,NO senza STOP")
