# -*- coding: utf-8 -*-
"""
Genera clienti_revisioni.xlsx (versione 2) — elenco clienti con scadenze revisione.
Novità v2:
 - colonna "Revisione Fatta" (SÌ/NO): SÌ = nessuna email per quel cliente
 - colonna "Scadenza Revisione" in formato MESE/ANNO come sul libretto (es. 09/2026)
Le date di esempio sono relative a OGGI (mese corrente +/-) così i casi sono realistici.
"""
import sys, os
import datetime as dt

XLSX_SKILL_DIR = "/home/z/my-project/skills/xlsx"
for sub in [XLSX_SKILL_DIR, os.path.join(XLSX_SKILL_DIR, "templates")]:
    if sub not in sys.path:
        sys.path.insert(0, sub)

import base as B
from openpyxl import Workbook
from openpyxl.styles import Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule

B.FONT_NAME = "Calibri"
B.HEADER_BOLD = True

OUT = "/home/z/my-project/download/RevMemo/clienti_revisioni.xlsx"

OGGI = dt.date.today()


def mese_libretto(offset_mesi):
    """Restituisce 'MM/YYYY' spostandosi di offset_mesi dal mese corrente."""
    y = OGGI.year + (OGGI.month - 1 + offset_mesi) // 12
    m = (OGGI.month - 1 + offset_mesi) % 12 + 1
    return f"{m:02d}/{y}"


# (Nome, Email, Telefono, Targa, mese libretto, Revisione Fatta, Attivo)
CLIENTI = [
    ("Mario Rossi",   "mario.rossi@email.it",   "347 1122334", "AB123CD", mese_libretto(1),  "NO", "SÌ"),
    ("Luigi Verdi",   "luigi.verdi@email.it",   "339 5566778", "EF456GH", mese_libretto(2),  "NO", "SÌ"),
    ("Anna Bianchi",  "anna.bianchi@email.it",  "331 8899001", "JK789LM", mese_libretto(0),  "NO", "SÌ"),
    ("Carlo Neri",    "carlo.neri@email.it",    "320 2233445", "OP321QR", mese_libretto(11), "NO", "SÌ"),
    ("Sofia Gialli",  "sofia.gialli@email.it",  "349 6677889", "ST654UV", mese_libretto(-1), "NO", "SÌ"),
    ("Paolo Marroni", "paolo.marroni@email.it", "333 4455667", "WX987YZ", mese_libretto(0),  "SÌ", "SÌ"),
    ("Elena Azzurri", "elena.azzurri@email.it", "346 7788990", "ZT147HP", mese_libretto(0),  "NO", "NO"),
    ("Franco Viola",  "franco.viola@email.it",  "328 2244668", "QR258WX", mese_libretto(3),  "NO", "NO"),
]

HEADERS = ["Nome Cliente", "Email", "Telefono", "Targa", "Scadenza Revisione", "Revisione Fatta", "Attivo"]

wb = Workbook()
ws = wb.active
ws.title = "Clienti"

last_col = len(HEADERS) + 1  # dati da colonna B (2) -> ultima G (7)

B.setup_sheet(ws, title="Scadenze Revisioni – Elenco Clienti", last_col=last_col)

for col_idx, header in enumerate(HEADERS, start=2):
    ws.cell(row=4, column=col_idx, value=header)
B.style_header_row(ws, row_num=4, col_start=2, col_end=last_col)

DATA_START = 5
for i, (nome, email, telefono, targa, scadenza, fatta, attivo) in enumerate(CLIENTI):
    r = DATA_START + i
    ws.cell(row=r, column=2, value=nome)
    ws.cell(row=r, column=3, value=email)
    ws.cell(row=r, column=4, value=telefono)
    ws.cell(row=r, column=5, value=targa)
    c_scad = ws.cell(row=r, column=6, value=scadenza)   # testo: "09/2026"
    c_scad.number_format = "@"                           # forza formato testo
    ws.cell(row=r, column=7, value=fatta)
    ws.cell(row=r, column=8, value=attivo)
    B.style_data_row(ws, row_num=r, col_start=2, col_end=last_col, row_index=i)
    for col in (4, 6, 7, 8):                             # telefono/scadenza/fatta/attivo centrati
        ws.cell(row=r, column=col).alignment = B.align_date()

LAST_DATA = DATA_START + len(CLIENTI) - 1

# --- nota a fondo tabella ---
nota_row = LAST_DATA + 2
ws.cell(row=nota_row, column=2,
        value="Come si usa: una riga per VEICOLO (stesso cliente con 2 macchine = 2 righe). Scrivi la scadenza "
              "come sul libretto (es. 09/2026): il programma calcola da solo l'ultimo giorno del mese. "
              "Revisione Fatta = SÌ blocca tutte le email per quel veicolo (quando rinnovi: aggiorna il mese e rimetti NO). "
              "Attivo = NO esclude il cliente. Promemoria a 30 giorni, 10 giorni e 1 giorno dalla fine del mese di scadenza.")
ws.merge_cells(start_row=nota_row, start_column=2, end_row=nota_row, end_column=last_col)
nc = ws.cell(row=nota_row, column=2)
nc.font = B.font_caption()
nc.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
ws.row_dimensions[nota_row].height = 40

# --- menu a tendina SÌ/NO su Revisione Fatta (G) e Attivo (H) ---
dv = DataValidation(type="list", formula1='"SÌ,NO"', allow_blank=True,
                    showDropDown=False, errorTitle="Valore non valido",
                    error="Scrivi SÌ oppure NO (usa il menu a tendina).")
ws.add_data_validation(dv)
dv.add(f"G{DATA_START}:G300")
dv.add(f"H{DATA_START}:H300")

# --- formattazione condizionale: SÌ verde, NO rosso su entrambe le colonne ---
for rng in (f"G{DATA_START}:G300", f"H{DATA_START}:H300"):
    ws.conditional_formatting.add(rng, CellIsRule(
        operator="equal", formula=['"SÌ"'], fill=B.CF_POSITIVE_FILL, font=B.CF_POSITIVE_FONT))
    ws.conditional_formatting.add(rng, CellIsRule(
        operator="equal", formula=['"NO"'], fill=B.CF_NEGATIVE_FILL, font=B.CF_NEGATIVE_FONT))

ws.freeze_panes = "B5"
B.auto_fit_columns(ws, min_width=8, max_width=30, header_row=4, data_start_row=DATA_START)

wb.properties.creator = "Z.ai"
wb.save(OUT)
print("Creato:", OUT)
print("Mese corrente:", mese_libretto(0))
for nome, _, telefono, targa, scad, fatta, attivo in CLIENTI:
    print(f"  {targa}  {nome:<15} tel={telefono}  libretto={scad}  Fatta={fatta}  Attivo={attivo}")
