#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
=====================================================================
 PROMEMORIA REVISIONI AUTOMATICI  —  versione 2
=====================================================================
Cosa fa:
 - Legge il file Excel dei clienti (clienti_revisioni.xlsx)
 - Per ogni cliente con Attivo = SÌ e Revisione Fatta = NO, calcola
   quanti giorni mancano alla scadenza della revisione
 - Invia un'email di promemoria: 1 mese prima, 10 giorni prima e
   1 giorno prima (vedi FINESTRE_PROMEMORIA qui sotto)

Novità versione 2:
 - Colonna "Revisione Fatta": se SÌ, non parte più NESSUNA email
   per quel cliente (finché non aggiorni la scadenza e rimetti NO)
 - Scadenza in formato mese/anno come sul libretto (es. 09/2026):
   la scadenza è l'ULTIMO giorno di quel mese, come da legge
 - Finestre di tolleranza: se il PC è spento il giorno esatto,
   il promemoria parte comunque nei giorni successivi (una volta sola)

Uso:
    python promemoria_revisioni.py              # invio reale delle email
    python promemoria_revisioni.py --dry-run    # PROVA: mostra senza inviare
    python promemoria_revisioni.py --file "altro_file.xlsx"

Requisiti:  pip install openpyxl
Gmail:      serve una "Password per app" di Google (vedi ISTRUZIONI.md)
=====================================================================
"""

import argparse
import calendar
import datetime as dt
import os
import smtplib
import sys
from email.message import EmailMessage

try:
    from openpyxl import load_workbook
except ImportError:
    print("Manca la libreria openpyxl. Installala con:  pip install openpyxl")
    sys.exit(1)

# =====================================================================
# 1) CONFIGURAZIONE — cambia QUI i tuoi dati
# =====================================================================

FILE_EXCEL = "clienti_revisioni.xlsx"   # elenco clienti (stessa cartella dello script)
FILE_LOG   = "reminders_log.txt"        # registro invii (anti-doppioni), si crea da solo

# --- Gmail SMTP ---
SMTP_HOST         = "smtp.gmail.com"
SMTP_PORT         = 465                                  # 465 = SSL
MITTENTE_EMAIL    = "riccardofreddi6@gmail.com"          # la tua casella Gmail
MITTENTE_PASSWORD = ""                                   # <-- INCOLLA qui la Password per app di Google (16 caratteri)

NOME_OFFICINA = "Officina Rossi"     # nome che appare nel mittente e nell'oggetto

# Punti di invio con finestra di tolleranza:
#   chiave = tipo promemoria, valore = (dal giorno, al giorno) di anticipo
#   30 -> invia quando alla scadenza restano da 26 a 30 giorni ("1 mese prima")
#   10 -> invia quando restano da 8 a 10 giorni ("10 giorni prima")
#    1 -> invia quando restano 1 o 2 giorni ("1 giorno prima", ultimo avviso)
# Il registro anti-doppioni fa sì che ogni tipo parta UNA SOLA VOLTA.
FINESTRE_PROMEMORIA = {
    30: (26, 30),
    10: (8, 10),
     1: (1, 2),
}
ETICHETTE = {30: "1 mese prima", 10: "10 giorni prima", 1: "1 giorno prima"}

# Nome delle colonne nel foglio Excel (riga 4). Puoi riordinare le colonne,
# ma NON rinominare queste intestazioni.
COLONNE = {
    "nome":     "Nome Cliente",
    "email":    "Email",
    "targa":    "Targa",
    "scadenza": "Scadenza Revisione",
    "fatta":    "Revisione Fatta",
    "attivo":   "Attivo",
}
RIGA_INTESTAZIONI = 4     # nel file Excel le intestazioni sono alla riga 4
FOGLIO            = "Clienti"

# =====================================================================
# 2) LETTURA DEL FILE EXCEL
# =====================================================================

def percorsi_base():
    """Cartella dove si trova questo script (così funziona anche da
    Utilità di pianificazione Windows, che parte da un'altra cartella)."""
    return os.path.dirname(os.path.abspath(__file__))


def valore_si(v):
    """Considera SÌ: SÌ / SI / S / TRUE / VERO / 1 (maiuscole indifferenti)."""
    if v is None:
        return False
    return str(v).strip().upper() in ("SÌ", "SI", "S", "TRUE", "VERO", "1")


def ultimo_giorno_mese(anno, mese):
    """Ultimo giorno del mese: 09/2026 -> 30/09/2026."""
    return dt.date(anno, mese, calendar.monthrange(anno, mese)[1])


def converti_scadenza(val):
    """Converte la cella 'Scadenza Revisione' in (data_scadenza, da_mese).
    Formati accettati:
      "09/2026" oppure "9/2026"  -> scade l'ULTIMO giorno di settembre 2026
      data di Excel 01/09/2026   -> se Excel trasforma 09/2026 in data,
                                    giorno 1 = segnale del mese -> ultimo giorno
      "15/01/2026"               -> data esatta (raro, ma gestito)
    Restituisce (None, False) se non riesce a leggere."""
    if isinstance(val, dt.datetime):
        val = val.date()
    if isinstance(val, dt.date):
        if val.day == 1:   # Excel ha trasformato "MM/YYYY" nella data del giorno 1
            return ultimo_giorno_mese(val.year, val.month), True
        return val, False
    if isinstance(val, str):
        v = val.strip()
        # prima il formato mese/anno del libretto: 09/2026
        for formato in ("%m/%Y", "%m-%Y"):
            try:
                d = dt.datetime.strptime(v, formato)
                return ultimo_giorno_mese(d.year, d.month), True
            except ValueError:
                continue
        # poi le date complete
        for formato in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d"):
            try:
                return dt.datetime.strptime(v, formato).date(), False
            except ValueError:
                continue
    return None, False


def leggi_clienti(percorso_excel):
    """Legge il foglio Clienti e restituisce una lista di dict.
    Trova le colonne dal nome dell'intestazione (non dalla posizione)."""
    if not os.path.exists(percorso_excel):
        print(f"ERRORE: non trovo il file Excel: {percorso_excel}")
        sys.exit(1)

    wb = load_workbook(percorso_excel, data_only=True)
    ws = wb[FOGLIO] if FOGLIO in wb.sheetnames else wb.active

    # mappa nome-intestazione -> numero colonna
    indice = {}
    for col in range(1, ws.max_column + 1):
        val = ws.cell(row=RIGA_INTESTAZIONI, column=col).value
        if val is not None:
            indice[str(val).strip().lower()] = col

    colonna_di, mancanti = {}, []
    for chiave, nome_col in COLONNE.items():
        c = indice.get(nome_col.strip().lower())
        if c is None:
            mancanti.append(nome_col)
        else:
            colonna_di[chiave] = c
    if mancanti:
        print("ERRORE: nel foglio Excel mancano le colonne: " + ", ".join(mancanti))
        print("Le intestazioni alla riga 4 devono essere esattamente: "
              + " | ".join(COLONNE.values()))
        print("(oppure scarica di nuovo il file clienti_revisioni aggiornato)")
        sys.exit(1)

    clienti = []
    for row in range(RIGA_INTESTAZIONI + 1, ws.max_row + 1):
        nome     = ws.cell(row=row, column=colonna_di["nome"]).value
        email    = ws.cell(row=row, column=colonna_di["email"]).value
        targa    = ws.cell(row=row, column=colonna_di["targa"]).value
        scadenza = ws.cell(row=row, column=colonna_di["scadenza"]).value
        fatta    = ws.cell(row=row, column=colonna_di["fatta"]).value
        attivo   = ws.cell(row=row, column=colonna_di["attivo"]).value

        # riga vuota o nota a fondo tabella -> salta.
        # Una riga cliente ha SEMPRE almeno uno tra email, targa e scadenza.
        if all(v is None or str(v).strip() == "" for v in (email, targa, scadenza)):
            continue

        data_scad, da_mese = converti_scadenza(scadenza)
        clienti.append({
            "riga":    row,
            "nome":    (str(nome).strip() if nome else ""),
            "email":   (str(email).strip() if email else ""),
            "targa":   (str(targa).strip().upper() if targa else ""),
            "scadenza": data_scad,
            "da_mese": da_mese,
            "fatta":   valore_si(fatta),
            "attivo":  valore_si(attivo),
        })
    return clienti

# =====================================================================
# 3) REGISTRO INVII (per non mandare email doppie)
# =====================================================================

def percorso_log():
    return os.path.join(percorsi_base(), FILE_LOG)


def gia_inviato(email, targa, tipo):
    """True se oggi è già partito questo tipo di promemoria a questo cliente."""
    percorso = percorso_log()
    if not os.path.exists(percorso):
        return False
    riga_cercata = f"{dt.date.today().isoformat()};{email.lower()};{targa};{tipo};"
    with open(percorso, "r", encoding="utf-8") as f:
        return any(line.strip().startswith(riga_cercata) for line in f)


def registra_invio(email, targa, tipo, giorni, descrizione):
    """Aggiunge una riga al registro: data;email;targa;tipo;descrizione"""
    riga = f"{dt.date.today().isoformat()};{email.lower()};{targa};{tipo};{descrizione} ({giorni} gg)"
    with open(percorso_log(), "a", encoding="utf-8") as f:
        f.write(riga + "\n")

# =====================================================================
# 4) COMPOSIZIONE E INVIO EMAIL
# =====================================================================

def data_it(d):
    return d.strftime("%d/%m/%Y")


def componi_email(cliente, giorni):
    """Email essenziale di 3 righe. 'entro il' = ultimo giorno utile."""
    targa   = cliente["targa"]
    oggetto = (f"[{NOME_OFFICINA}] Promemoria revisione {targa} – scade tra {giorni} giorni"
               if giorni > 1 else
               f"[{NOME_OFFICINA}] Promemoria revisione {targa} – scade tra 1 giorno")
    corpo = (
        f"La revisione del veicolo targato {targa} scade entro il {data_it(cliente['scadenza'])}.\n"
        f"Giorni rimanenti: {giorni}.\n"
        f"La preghiamo di provvedere per tempo.\n"
    )
    return oggetto, corpo


def invia_email(destinatario, oggetto, corpo):
    msg = EmailMessage()
    msg["From"]    = f"{NOME_OFFICINA} <{MITTENTE_EMAIL}>"   # es: Officina Rossi <riccardofreddi6@gmail.com>
    msg["To"]      = destinatario
    msg["Subject"] = oggetto
    msg.set_content(corpo)

    with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT) as server:
        server.login(MITTENTE_EMAIL, MITTENTE_PASSWORD)
        server.send_message(msg)

# =====================================================================
# 5) PROGRAMMA PRINCIPALE
# =====================================================================

def trova_tipo(giorni):
    """Restituisce il tipo di promemoria se 'giorni' cade in una finestra."""
    for tipo, (da, a) in FINESTRE_PROMEMORIA.items():
        if da <= giorni <= a:
            return tipo
    return None


def main():
    parser = argparse.ArgumentParser(description="Promemoria scadenza revisioni via email.")
    parser.add_argument("--dry-run", action="store_true",
                        help="modalità prova: mostra cosa verrebbe inviato SENZA inviare nulla")
    parser.add_argument("--file", default=None,
                        help="percorso alternativo del file Excel dei clienti")
    args = parser.parse_args()

    percorso_excel = args.file or os.path.join(percorsi_base(), FILE_EXCEL)

    print("=" * 62)
    print("  PROMEMORIA REVISIONI", "(MODALITÀ PROVA – nessuna email partirà)" if args.dry_run else "")
    print("=" * 62)
    print(f"Oggi: {data_it(dt.date.today())}")
    print(f"File clienti: {percorso_excel}")
    if args.dry_run and not MITTENTE_PASSWORD:
        print("Nota: password di Gmail non ancora inserita — la prova funziona lo stesso.")
    print()

    if not args.dry_run:
        if not MITTENTE_EMAIL or MITTENTE_PASSWORD == "":
            print("ERRORE: prima dell'invio vero devi inserire la Password per app di Google\n"
                  "nella sezione CONFIGURAZIONE in cima allo script (riga ~45):\n"
                  '    MITTENTE_PASSWORD = "la-tua-password-di-16-caratteri"\n'
                  "Vedi ISTRUZIONI.md punto 3 per come crearla passo passo.")
            sys.exit(1)

    clienti = leggi_clienti(percorso_excel)
    print(f"Trovati {len(clienti)} clienti nel file.\n")

    inviati = saltati_log = esclusi = errori = 0
    ogggi = dt.date.today()

    for c in clienti:
        targa = c["targa"] or "(targa mancante)"
        etichetta_riga = f"riga {c['riga']:>2} | {c['nome'] or '(senza nome)':<15} | {targa:<8} | "

        # -- controlli preliminari -----------------------------------
        if not c["attivo"]:
            print(f"  - {etichetta_riga}cliente NON attivo -> saltato")
            esclusi += 1
            continue
        if c["fatta"]:
            print(f"  - {etichetta_riga}revisione GIÀ FATTA -> nessuna email (bloccata dalla colonna Revisione Fatta)")
            esclusi += 1
            continue
        if not c["email"] or "@" not in c["email"]:
            print(f"  - {etichetta_riga}email mancante o non valida -> saltato")
            esclusi += 1
            continue
        if c["scadenza"] is None:
            print(f"  - {etichetta_riga}scadenza non leggibile (scrivi es. 09/2026) -> saltato")
            esclusi += 1
            continue

        giorni = (c["scadenza"] - ogggi).days

        if giorni < 0:
            print(f"  - {etichetta_riga}revisione SCADUTA da {-giorni} gg (fine mese {data_it(c['scadenza'])}) -> nessun invio")
            esclusi += 1
            continue
        if giorni == 0:
            print(f"  - {etichetta_riga}scade OGGI ({data_it(c['scadenza'])}): ultimo giorno utile, nessun promemoria")
            esclusi += 1
            continue

        tipo = trova_tipo(giorni)
        if tipo is None:
            print(f"  - {etichetta_riga}mancano {giorni:>3} gg (entro il {data_it(c['scadenza'])}) -> non è momento di inviare")
            esclusi += 1
            continue

        descrizione = ETICHETTE[tipo]
        oggetto, corpo = componi_email(c, giorni)

        # -- doppioni? ------------------------------------------------
        if gia_inviato(c["email"], c["targa"], tipo):
            print(f"  - {etichetta_riga}promemoria '{descrizione}' GIÀ INVIATO -> saltato")
            saltati_log += 1
            continue

        # -- invio ----------------------------------------------------
        if args.dry_run:
            print(f"  > {etichetta_riga}PROVA invio '{descrizione}' a {c['email']}")
            print(f"      Oggetto: {oggetto}")
            for linea in corpo.strip().splitlines():
                print(f"      {linea}")
            inviati += 1
            continue

        try:
            invia_email(c["email"], oggetto, corpo)
            registra_invio(c["email"], c["targa"], tipo, giorni, descrizione)
            print(f"  ✓ {etichetta_riga}email INVIATA a {c['email']} ({descrizione})")
            inviati += 1
        except Exception as e:
            print(f"  ✗ {etichetta_riga}ERRORE invio a {c['email']}: {e}")
            errori += 1

    # -- riepilogo ---------------------------------------------------
    print()
    print("-" * 62)
    if args.dry_run:
        print(f"RIEPILOGO PROVA: {inviati} email verrebbero inviate, "
              f"{saltati_log} già inviate oggi, {esclusi} esclusi.")
        print("Tutto corretto? Lancia l'invio reale con:  avvia_promemoria.bat")
    else:
        print(f"RIEPILOGO: {inviati} inviate, {saltati_log} doppioni saltati, "
              f"{esclusi} esclusi, {errori} errori.")
        if inviati > 0:
            print(f"Registro aggiornato: {percorso_log()}")
    print("-" * 62)
    return 0 if errori == 0 else 2


if __name__ == "__main__":
    sys.exit(main())
