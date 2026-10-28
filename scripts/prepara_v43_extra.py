# -*- coding: utf-8 -*-
"""v4.3 — crea colora_excel.bat (CRLF) e aggiorna LEGGIMI.txt mantenendo
i fine riga CRLF. Ogni sostituzione deve avvenire UNA sola volta."""
import io
import os

PACCHETTO = "/home/z/my-project/download/RevMemo"
BAT = os.path.join(PACCHETTO, "colora_excel.bat")
LEGGIMI = os.path.join(PACCHETTO, "LEGGIMI.txt")

# ---------------------------------------------------------------- bat
CONTENUTO_BAT = "\r\n".join([
    "@echo off",
    "chcp 65001 >nul",
    'cd /d "%~dp0"',
    'set "PYCMD="',
    'if exist "runtime\\python.exe" set "PYCMD=runtime\\python.exe"',
    'if not defined PYCMD where python >nul 2>nul && set "PYCMD=python"',
    'if not defined PYCMD where py >nul 2>nul && set "PYCMD=py"',
    'if not defined PYCMD echo Python non trovato: vedi LEGGIMI.txt '
    '& pause & exit /b 1',
    "echo.",
    "echo   COLORI NELL'EXCEL: SI = verde, NO = rosso, STOP = rosso scuro",
    "echo   (non tocca i tuoi clienti)",
    "echo.",
    "%PYCMD% revmemo.py --colora-excel",
    "echo.",
    "pause",
    "",
])
with io.open(BAT, "w", encoding="utf-8", newline="") as f:
    f.write(CONTENUTO_BAT)
print(f"Creato: {BAT} ({os.path.getsize(BAT)} byte)")

# ------------------------------------------------------------- leggimi
with io.open(LEGGIMI, "r", encoding="utf-8", newline="") as f:
    testo = f.read()


def sostituisci(testo, vecchio, nuovo, etichetta):
    """Sostituisce esattamente UNA volta, altrimenti solleva errore."""
    quante = testo.count(vecchio)
    if quante != 1:
        raise SystemExit(f"ERRORE ({etichetta}): trovate {quante} occorrenze, "
                         f"attesa 1. Fragmento: {vecchio[:60]!r}")
    print(f"OK patch: {etichetta}")
    return testo.replace(vecchio, nuovo)


# 1) sezione 2: elenco file del pacchetto
testo = sostituisci(
    testo,
    "- controlla_pc.bat ........ controllo quando qualcosa non va (te lo dice)\r\n",
    "- controlla_pc.bat ........ controllo quando qualcosa non va (te lo dice)\r\n"
    "- colora_excel.bat ........ mette i colori nell'Excel (verde/rosso),"
    " si usa una volta\r\n",
    "sezione 2: elenco file")

# 2) sezione 4: paragrafo sui colori prima di "Regole d'oro:"
testo = sostituisci(
    testo,
    "                         STOP = il cliente l'ha chiesto (vedi sezione 6)\r\n"
    "\r\n"
    "Regole d'oro:",
    "                         STOP = il cliente l'ha chiesto (vedi sezione 6)\r\n"
    "\r\n"
    "Colori automatici: riga VERDE = cliente attivo, riga ROSSA = NO oppure\r\n"
    "STOP (la scritta STOP e' rosso scuro). I colori cambiano da soli quando\r\n"
    "scrivi SÌ o NO, senza rilanciare il programma.\r\n"
    "Se il tuo file non si colora (nato con una versione vecchia): doppio\r\n"
    "click su colora_excel.bat - non tocca i tuoi clienti.\r\n"
    "La colonna Scadenza e' in formato TESTO: scrivi 01/2026 e resta cosi'\r\n"
    "(non diventa \"gen-26\"; se era gia' diventato gen-26, colora_excel.bat\r\n"
    "lo riscrive 01/2026 da solo).\r\n"
    "\r\n"
    "Regole d'oro:",
    "sezione 4: colori automatici")

# 3) sezione 9: FAQ "gen-26" prima di "- Ho sbagliato un invio"
testo = sostituisci(
    testo,
    "- Ho sbagliato un invio",
    "- Nell'Excel vedo \"gen-26\" al posto di 01/2026\r\n"
    "    E' la stessa cosa per il programma (calcola sempre il FINE mese\r\n"
    "    di quel mese, come sul libretto). Se vuoi vedere 01/2026:\r\n"
    "    doppio click su colora_excel.bat e la colonna diventa testo\r\n"
    "- Ho sbagliato un invio",
    "sezione 9: FAQ gen-26")

with io.open(LEGGIMI, "w", encoding="utf-8", newline="") as f:
    f.write(testo)

crfl_ok = "\r\n" in testo and testo.count("\n") == testo.count("\r\n")
print(f"Aggiornato: {LEGGIMI} ({len(testo)} caratteri, CRLF coerente: {crfl_ok})")
