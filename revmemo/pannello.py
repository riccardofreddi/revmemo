#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
=====================================================================
 RevMemo - PANNELLO DI CONTROLLO (interfaccia grafica locale)
=====================================================================
Cosa fa:
 - Apre una pagina semplice sul computer del centro (niente internet:
   gira SOLO su questo PC, indirizzo 127.0.0.1)
 - Fa tutto con i pulsanti: messaggio delle email, impostazioni,
   GESTIONE CLIENTI (nuovo/modifica/SI-NO, con backup automatico),
   prova, invio vero (fermabile), controllo PC, cerca cliente,
   registro invii
 - NON tocca il motore: usa revmemo.py (gia' collaudato) e scrive
   i valori nello stesso config.ini di sempre (commenti compresi)
 - Nessun programma nuovo da installare: serve solo il Python
   gia' incluso nel pacchetto (nessun modulo extra)

Avvio: doppio clic su apri_pannello.bat
=====================================================================
"""
import json
import html as html_lib
import os
import secrets
import subprocess
import sys
import threading
import time
import webbrowser
from contextlib import contextmanager
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse

BASE = os.path.dirname(os.path.abspath(__file__))
# Il Python incluso (cartella runtime) gira in modalita' ISOLATA per via
# del file runtime\python312._pth: la cartella dello script NON viene
# aggiunta ai percorsi di ricerca, quindi "import revmemo" non troverebbe
# il motore (ModuleNotFoundError). La aggiungo io qui, per tutto.
if BASE not in sys.path:
    sys.path.insert(0, BASE)
PORTA = int(os.environ.get("PANNELLO_PORTA", "8768"))
URL_PANNELLO = f"http://127.0.0.1:{PORTA}/"

# Chiave segreta di QUESTA sessione del pannello: la pagina la riceve e la
# rimanda in ogni richiesta. Senza, una pagina web qualunque aperta nello
# stesso browser potrebbe far partire invii, cancellare clienti o cambiare
# il server di posta (e quindi rubare la password) parlando con 127.0.0.1.
TOKEN = secrets.token_urlsafe(24)

PERCORSO_CONFIG = os.path.join(BASE, "config.ini")
PERCORSO_EXCEL = os.path.join(BASE, "clienti_revisioni.xlsx")
PERCORSO_LOG = os.path.join(BASE, "reminders_log.txt")
PERCORSO_HTML = os.path.join(BASE, "pannello.html")
PERCORSO_LOG_PANNELLO = os.path.join(BASE, "pannello_log.txt")

# i campi che possono avere piu' righe (nel config: righe successive
# identate con qualche spazio)
MULTIRIGA = {"saluto", "testo", "firma"}

# chiavi gestite dal pannello, per sezione (l'ordine qui decide solo
# la lettura; la scrittura avviene AL POSTO GIUSTO nel file)
CHIAVI = {
    "smtp": ["host", "porta", "mittente", "password_app", "risposta_a"],
    "centro": ["nome", "citta", "indirizzo", "telefono", "email_contatti",
               "sito_web", "orari", "colore_tema"],
    "messaggio": ["oggetto", "oggetto_1mese", "oggetto_10giorni",
                  "oggetto_1giorno", "saluto", "testo", "firma",
                  "riga_extra_piede", "testo_disiscrizione"],
    "prodotto": ["nome"],
}

MOTORE = os.path.join(BASE, "revmemo.py")
CONTROLLO = os.path.join(BASE, "controlla_pc.py")

# ---------------------------------------------------------------------
# stato dell'azione in corso (prova / invio / controllo)
SERRATURE = threading.Lock()
SERRATURA_EXCEL = threading.Lock()      # una scrittura clienti per volta
AZIONE = {"in_corso": False, "tipo": None, "righe": [], "finito": True,
          "codice": None, "fermata": False}
PROCESSO = {"p": None}                  # processo figlio (per il pulsante FERMA)


def nota(testo):
    """piccolo log su file: utile per l'assistenza"""
    try:
        with open(PERCORSO_LOG_PANNELLO, "a", encoding="utf-8") as f:
            f.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')} {testo}\n")
    except Exception:
        pass


def eseguibile():
    """python.exe per i programmi figli (pythonw non ha console/stdout)"""
    exe = sys.executable
    if exe.lower().endswith("pythonw.exe"):
        alt = os.path.join(os.path.dirname(exe), "python.exe")
        if os.path.exists(alt):
            return alt
    return exe


def _bandiere():
    """niente finestra nera quando si lanciano i programmi figli"""
    if os.name == "nt":
        return 0x08000000          # CREATE_NO_WINDOW
    return 0


def ambiente_utf8():
    env = dict(os.environ)
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUTF8"] = "1"
    return env


# =====================================================================
#  LETTURA / SCRITTURA DI config.ini (AL POSTO GIUSTO, commenti ok)
# =====================================================================
def leggi_config_raw():
    """valori GREZZI scritti nel file (senza aggiungere i predefiniti),
    divisi per sezione: {"smtp": {...}, "centro": {...}, ...}"""
    import configparser
    cfg = configparser.ConfigParser(interpolation=None)
    try:
        cfg.read(PERCORSO_CONFIG, encoding="utf-8")
    except Exception as e:
        return {"errore": f"config.ini non si riesce a leggere: {e}"}
    fuori = {"errore": ""}
    for sezione, chiavi in CHIAVI.items():
        sec = cfg[sezione] if cfg.has_section(sezione) else {}
        fuori[sezione] = {k: sec.get(k, "") for k in chiavi}
    fuori["smtp"]["password_app"] = ""     # mai in chiaro verso il browser
    fuori["password_presente"] = bool(
        "".join((cfg["smtp"].get("password_app", "") if cfg.has_section("smtp")
                 else "").split()))
    return fuori


def salva_campi(modifiche):
    """aggiorna config.ini AL POSTO GIUSTO: conserva commenti, ordine,
    a capo Windows (CRLF) e le chiavi che il pannello non gestisce.
    modifiche = {"sezione": {"chiave": "valore"}} (valore None = ignora)"""
    with open(PERCORSO_CONFIG, "r", encoding="utf-8", newline="") as f:
        testo = f.read()
    nl = "\r\n" if "\r\n" in testo else "\n"
    righe = testo.split(nl)

    # valori multiriga: prima riga dopo il =, le altre identate
    def righe_valore(chiave, valore):
        valore = (valore or "").replace("\r\n", "\n").replace("\r", "\n")
        pezzi = [p.rstrip() for p in valore.split("\n")]
        while pezzi and not pezzi[0]:
            pezzi.pop(0)
        while pezzi and not pezzi[-1]:
            pezzi.pop()
        if not pezzi:
            return [f"{chiave} ="]
        prima = pezzi[0]
        resto = ["    " + p for p in pezzi[1:]]
        return [f"{chiave} = {prima}"] + resto

    def chiave_di(riga):
        nuda = riga.strip()
        if not nuda or nuda.startswith(";") or nuda.startswith("#"):
            return None
        if "=" not in riga:
            return None
        if riga[:1] in (" ", "\t"):
            return None          # riga di continuazione (multiriga)
        return riga.split("=", 1)[0].strip().lower()

    sezione = None
    fatte = []
    i = 0
    while i < len(righe):
        riga = righe[i]
        nuda = riga.strip()
        if nuda.startswith("[") and nuda.endswith("]") and "=" not in nuda:
            sezione = nuda[1:-1].strip().lower()
            i += 1
            continue
        if sezione in modifiche:
            k = chiave_di(riga)
            if k and k in modifiche[sezione] and modifiche[sezione][k] is not None:
                k_maiuscolo = riga.split("=", 1)[0].rstrip()
                nuove = righe_valore(k_maiuscolo, modifiche[sezione][k])
                # cancello le vecchie righe di continuazione: righe identate
                # E righe vuote INTERNE al blocco (una vuota conta solo se
                # dopo c'e' ancora testo identato; i separatori tra chiavi
                # restano al loro posto)
                j = i + 1
                while j < len(righe):
                    r2 = righe[j]
                    nuda2 = r2.strip()
                    if nuda2 and r2[:1] in (" ", "\t"):
                        j += 1
                        continue
                    if nuda2 == "":
                        k2 = j
                        while k2 < len(righe) and righe[k2].strip() == "":
                            k2 += 1
                        if k2 < len(righe) and righe[k2].strip() \
                                and righe[k2][:1] in (" ", "\t"):
                            j = k2
                            continue
                    break          # chiave/seziona/commento a colonna 0: stop
                righe[i:j] = nuove
                i += len(nuove)
                if sezione not in fatte:
                    fatte.append(sezione)
                continue
        i += 1

    # chiavi NUOVE (il file era di una versione vecchia senza quel campo)
    for sezione_nuova, campi in modifiche.items():
        for chiave, valore in campi.items():
            if valore is None or chiave in fatte or f"\n{chiave}" in testo \
                    or testo.startswith(f"{chiave}") \
                    or (f"{chiave} =" in testo):
                continue
            # trova la sezione nel file
            inizio = None
            for idx, r in enumerate(righe):
                if r.strip().lower() == f"[{sezione_nuova}]":
                    inizio = idx + 1
                    break
            if inizio is not None:
                fine = len(righe)
                for j in range(inizio, len(righe)):
                    nuda = righe[j].strip()
                    if nuda.startswith("[") and nuda.endswith("]"):
                        fine = j
                        break
                righe[fine:fine] = righe_valore(chiave, valore)

    with open(PERCORSO_CONFIG, "w", encoding="utf-8", newline="") as f:
        f.write(nl.join(righe))
    nota("config.ini aggiornato dal pannello: " + json.dumps(
        {s: sorted([k for k, v in c.items() if v is not None])
         for s, c in modifiche.items()}, ensure_ascii=False))


# =====================================================================
#  MOTORE (revmemo.py gia' collaudato: lo uso per i dati veri)
# =====================================================================
MSG_CONFIG_ROTTA = ("config.ini ha una riga con formato sbagliato. Causa piu' "
                    "comune: le righe dopo 'testo =' devono iniziare con "
                    "QUALCHE SPAZIO. Correggi il file (o riprendi "
                    "config_ESEMPIO.ini) e riapri il pannello.")


def motore():
    try:
        import revmemo
    except SystemExit:          # il motore esce se config.ini e' illeggibile
        raise RuntimeError(MSG_CONFIG_ROTTA) from None
    return revmemo


def ricarica_config(r):
    """config FRESCO dal disco. Ritorna "" se va bene, altrimenti il
    messaggio da mostrare (prima il pannello mostrava solo "1")."""
    try:
        r.CFG = r.carica_config()
    except SystemExit:
        return MSG_CONFIG_ROTTA
    r.NOME_CENTRO = r.CFG["nome_centro"]
    r.PRODOTTO = r.CFG["prodotto"]
    return ""


def _testo_scadenze():
    """numero di riga -> scadenza scritta NEL FILE cosi' com'e' (es.
    '09/2026' come da libretto, o la data completa se scritta completa).
    Serve per MOSTRARLA uguale a come la scrive il centro, senza
    trasformarla in 30/09/2026."""
    try:
        import datetime as dt_mod
        r = motore()
        wb = r.load_workbook(PERCORSO_EXCEL)
        ws = wb[r.FOGLIO] if r.FOGLIO in wb.sheetnames else wb.active
        colonna_di, _ = r.individua_colonne(ws)
        fuori = {}
        if "scadenza" in colonna_di:
            for row in range(r.RIGA_INTESTAZIONI + 1, LIMITE_RIGA + 1):
                v = ws.cell(row=row, column=colonna_di["scadenza"]).value
                if v in (None, ""):
                    continue
                if isinstance(v, dt_mod.date):
                    fuori[row] = v.strftime("%d/%m/%Y")
                    continue
                t = str(v).strip()
                # numero seriale di Excel senza formato data: strano,
                # lascio la data che ha gia' calcolato il motore
                if t.replace(".", "").isdigit() and len(t) <= 6:
                    continue
                fuori[row] = t
        return fuori
    except Exception:
        return {}


def dati_stato():
    r = motore()
    # config FRESCO dal disco (il pannello l'ha appena aggiornato)
    errore_config = ricarica_config(r)
    cfg = leggi_config_raw()
    smtp_cfg = cfg.get("smtp", {})
    oggi = r.dt.date.today()
    mittente_val = (smtp_cfg.get("mittente", "")
                    if isinstance(smtp_cfg, dict) else "")
    host_val = (r.CFG.get("smtp_host") or "smtp.gmail.com").lower()
    porta_val = r.CFG.get("smtp_porta", 465)
    password_ok = bool(cfg.get("password_presente"))
    # stato della casella email: ok / parziale (casella non Gmail ma
    # server ancora Gmail) / da_fare (mancano email o password)
    if not mittente_val or not password_ok:
        email_stato, email_msg = "da_fare", (
            "Da completare: scrivi l'email e la Password per app qui sopra.")
    elif "gmail.com" not in mittente_val.lower().split("@")[-1] \
            and "gmail" in host_val:
        email_stato, email_msg = "parziale", (
            "Hai messo una casella NON Gmail ma il server è ancora quello "
            "di Gmail: compila anche SERVER e PORTA nella carta qui sotto.")
    else:
        email_stato, email_msg = "ok", ""
    stato = {
        "nome_centro": r.CFG["nome_centro"],
        "prodotto": r.CFG["prodotto"],
        "mittente": mittente_val,
        "gmail_ok": bool(mittente_val) and password_ok,
        "smtp_host": host_val,
        "smtp_porta": porta_val,
        "email_stato": email_stato,
        "email_messaggio": email_msg,
        "excel_presente": os.path.exists(PERCORSO_EXCEL),
        "log_presente": os.path.exists(PERCORSO_LOG),
        "clienti": None,
        "errore": cfg.get("errore", "") or errore_config,
    }
    if stato["excel_presente"]:
        try:
            clienti = r.leggi_clienti(PERCORSO_EXCEL)
            attivi = [c for c in clienti if c["attivo"]]
            testi = _testo_scadenze()

            def scadenza_mostrata(c):
                """come scritta nel file (es. 09/2026); se non c'e' testo,
                la data calcolata dal motore"""
                return testi.get(c["riga"]) or (c["scadenza"].strftime(
                    "%d/%m/%Y") if c["scadenza"] else "")

            in_finestra = []
            for c in attivi:
                if c["fatta"] or not c["scadenza"] or not c["email"]:
                    continue
                giorni = (c["scadenza"] - oggi).days
                if giorni <= 0:
                    continue
                tipo = r.trova_tipo(giorni)
                if tipo is None:
                    continue
                gia = False
                try:
                    gia = r.gia_inviato(c["email"], c["targa"], tipo,
                                        c["scadenza"].isoformat())
                except Exception:
                    pass
                in_finestra.append({"nome": c["nome"] or "(senza nome)",
                                    "targa": c["targa"] or "",
                                    "scadenza": scadenza_mostrata(c),
                                    "giorni": giorni,
                                    "inviato": gia})
            tutte_prossime = sorted(
                ({"nome": c["nome"] or "(senza nome)",
                  "targa": c["targa"] or "",
                  "scadenza": scadenza_mostrata(c),
                  "giorni": (c["scadenza"] - oggi).days}
                 for c in attivi
                 if c["scadenza"] and not c["fatta"]
                 and (c["scadenza"] - oggi).days > 0),
                key=lambda x: x["giorni"])
            # TUTTE le prossime scadenze: in Home la tabella resta alta
            # 10 righe e il resto si raggiunge con la barra di scorrimento
            prossime = tutte_prossime
            elenco = [{"nome": c["nome"] or "(senza nome)",
                       "targa": c["targa"] or "",
                       "email": c["email"] or "",
                       "scadenza": scadenza_mostrata(c),
                       "fatta": bool(c["fatta"]),
                       "attivo": _attivo_visuale(c["attivo_testo"],
                                                 c["attivo"])}
                      for c in clienti]
            stato["clienti"] = {
                "totali": len(clienti),
                "attivi": len(attivi),
                "in_stop": sum(1 for c in clienti
                               if c["attivo_testo"] == "STOP"),
                "fatte": sum(1 for c in clienti if c["fatta"]),
                "in_finestra_oggi": len(in_finestra),
                "da_inviare_oggi": sum(1 for x in in_finestra if not x["inviato"]),
                "oggi_elenco": in_finestra,
                "prossime": prossime,
                "prossime_altre": max(0, len(tutte_prossime)
                                      - len(prossime)),
                "elenco": elenco[:40],
                "elenco_tagliato": max(0, len(clienti) - 40),
            }
        except BaseException as e:      # anche SystemExit del motore
            msg = str(e).strip() or "controlla le intestazioni (riga 4)"
            stato["errore"] = f"Il file clienti ha un problema: {msg}"
    # statistiche invii (mesi dinamici): servono alla DASHBOARD in Home
    stato["statistiche"] = statistiche_invii()
    return stato


def leggi_registro(completo=False):
    """righe del registro (reminders_log.txt), dalla piu' recente.
    completo=False -> le 250 per la tabella; completo=True -> TUTTE
    (serve per l'esportazione in Excel e le statistiche)."""
    if not os.path.exists(PERCORSO_LOG):
        return []
    righe = []
    try:
        with open(PERCORSO_LOG, "r", encoding="utf-8", errors="replace") as f:
            righe = f.read().splitlines()
    except Exception:
        return []
    fuori = []
    for riga in reversed(righe):
        if not riga.strip():
            continue
        pezzi = riga.split(";")
        if len(pezzi) >= 6:
            fuori.append({"data": pezzi[0], "email": pezzi[1],
                          "targa": pezzi[2], "tipo": pezzi[3],
                          "scadenza": pezzi[4],
                          "descrizione": ";".join(pezzi[5:])})
        else:
            fuori.append({"data": riga[:10], "email": "", "targa": "",
                          "tipo": "", "scadenza": "", "descrizione": riga})
        if not completo and len(fuori) >= 250:
            break
    return fuori


_MESI_IT = ("gennaio", "febbraio", "marzo", "aprile", "maggio", "giugno",
            "luglio", "agosto", "settembre", "ottobre", "novembre",
            "dicembre")
_TIPO_PROMEMORIA = {"30": "30 giorni", "10": "10 giorni", "1": "1 giorno"}


def statistiche_invii():
    """quanti promemoria sono partiti ogni mese: la finestra parte dal
    MESE DEL PRIMO INVIO registrato nel registro (non da 6 mesi fissi:
    se il centro usa il programma da ottobre, a ottobre si ferma) e
    arriva fino al mese corrente, con un massimo di 12 colonne per non
    affollare il grafico. Prima del primo invio mostra solo il mese
    corrente a zero, senza colonne vuote di mesi mai usati."""
    import datetime as dt_mod
    tutti = leggi_registro(completo=True)
    conteggi = {}
    tot_validi = 0
    primo_mese = None
    for x in tutti:
        d = (x.get("data") or "").strip()
        if len(d) >= 7 and d[:4].isdigit() and d[5:7].isdigit():
            chiave = d[:7]                    # AAAA-MM
            conteggi[chiave] = conteggi.get(chiave, 0) + 1
            tot_validi += 1
            if primo_mese is None or chiave < primo_mese:
                primo_mese = chiave
    oggi = dt_mod.date.today()
    # gli ultimi 12 mesi (dal piu' vecchio al corrente)
    chiavi = []
    anno, mese = oggi.year, oggi.month
    for _ in range(12):
        chiavi.append(f"{anno:04d}-{mese:02d}")
        mese -= 1
        if mese == 0:
            mese, anno = 12, anno - 1
    chiavi.reverse()
    if primo_mese and primo_mese in chiavi:
        chiavi = chiavi[chiavi.index(primo_mese):]
    else:
        chiavi = chiavi[-1:]     # niente invii: solo il mese corrente
    mesi = [{"chiave": k,
             "etichetta": f"{_MESI_IT[int(k[5:7]) - 1][:3]} {k[:4]}",
             "quanti": conteggi.get(k, 0)} for k in chiavi]
    return {"mesi": mesi,
            "questo_mese": conteggi.get(f"{oggi.year:04d}-{oggi.month:02d}", 0),
            "totale": tot_validi}


PERCORSO_REGISTRO_EXCEL = os.path.join(BASE, "registro_invii.xlsx")


def esporta_registro():
    """esporta TUTTO il registro (non solo le ultime 250) in un file
    Excel (registro_invii.xlsx) e lo apre: il centro puo' cosi'
    archiviare/stampare i promemoria inviati."""
    if not os.path.exists(PERCORSO_LOG):
        return False, "Il registro e' ancora vuoto: non c'e' niente da " \
                      "esportare. Esporta dopo il primo INVIA ORA.", 0
    righe = leggi_registro(completo=True)
    righe = list(reversed(righe))          # dalla piu' vecchia: cronologico
    if not righe:
        return False, "Il registro e' ancora vuoto: non c'e' niente da " \
                      "esportare. Esporta dopo il primo INVIA ORA.", 0
    try:
        import openpyxl
        from openpyxl.styles import Alignment, Font, PatternFill
        from openpyxl.utils import get_column_letter
    except Exception as e:
        return False, f"Esportazione non riuscita (openpyxl): {e}", 0
    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Registro invii"
        intestazioni = ["Data invio", "Email cliente", "Targa",
                        "Promemoria", "Scadenza revisione", "Dettagli"]
        fill = PatternFill("solid", fgColor="0F2B46")
        font = Font(bold=True, color="FFFFFF")
        for col, titolo in enumerate(intestazioni, start=1):
            cella = ws.cell(row=1, column=col, value=titolo)
            cella.fill = fill
            cella.font = font
            cella.alignment = Alignment(vertical="center")
        for num, x in enumerate(righe, start=2):
            d = (x.get("data") or "").strip()
            if len(d) == 10 and d[4] == "-" and d[7] == "-":
                d = f"{d[8:10]}/{d[5:7]}/{d[0:4]}"
            scad = (x.get("scadenza") or "").strip()
            if len(scad) == 10 and scad[4] == "-" and scad[7] == "-":
                scad = f"{scad[8:10]}/{scad[5:7]}/{scad[0:4]}"
            ws.cell(row=num, column=1, value=d)
            ws.cell(row=num, column=2, value=x.get("email") or "")
            ws.cell(row=num, column=3, value=x.get("targa") or "")
            ws.cell(row=num, column=4,
                    value=_TIPO_PROMEMORIA.get(x.get("tipo"), x.get("tipo") or ""))
            ws.cell(row=num, column=5, value=scad)
            ws.cell(row=num, column=6, value=x.get("descrizione") or "")
        larghezze = [12, 30, 12, 12, 18, 34]
        for col, largh in enumerate(larghezze, start=1):
            ws.column_dimensions[get_column_letter(col)].width = largh
        ws.freeze_panes = "A2"
        try:
            wb.save(PERCORSO_REGISTRO_EXCEL)
        except PermissionError:
            return False, ("Il file registro_invii.xlsx risulta APERTO: "
                           "chiudi Excel e riprova."), 0
        except Exception as e:
            return False, f"Non riesco a salvare il registro: {e}", 0
    except Exception as e:
        return False, f"Esportazione non riuscita: {e}", 0
    try:
        if os.name == "nt":
            os.startfile(PERCORSO_REGISTRO_EXCEL)      # noqa: solo Windows
        elif sys.platform == "darwin":
            subprocess.Popen(["open", PERCORSO_REGISTRO_EXCEL])
        else:
            subprocess.Popen(["xdg-open", PERCORSO_REGISTRO_EXCEL])
    except Exception:
        pass                     # il file c'e': l'apertura non e' critica
    nota(f"registro esportato in Excel: {len(righe)} righe")
    return True, f"Esportate {len(righe)} righe in registro_invii.xlsx", \
        len(righe)


# =====================================================================
#  AZIONI (prova / invio / controllo) in background
# =====================================================================
def lancia_azione(tipo):
    with SERRATURE:
        if AZIONE["in_corso"]:
            return False, "C'e' gia' un'azione in corso: aspetta che finisca."
        comandi = {
            "prova": [eseguibile(), MOTORE, "--dry-run"],
            "invio": [eseguibile(), MOTORE],
            "controllo": [eseguibile(), CONTROLLO],
        }
        if tipo not in comandi:
            return False, "Azione sconosciuta."
        if tipo == "invio" and not os.path.exists(PERCORSO_EXCEL):
            return False, "Prima crea il file clienti (si crea aprendo il pannello)."
        AZIONE.update({"in_corso": True, "tipo": tipo, "righe": [],
                       "finito": False, "codice": None, "fermata": False})
    threading.Thread(target=_esegui, args=(tipo, comandi[tipo]),
                     daemon=True).start()
    return True, ""


def _pid_vivo(pid):
    """True se quel numero di processo sta girando adesso (Windows e Linux)."""
    if not pid or int(pid) <= 0:
        return False
    if os.name == "nt":
        try:
            import ctypes
            k32 = ctypes.windll.kernel32
            handle = k32.OpenProcess(0x1000, False, int(pid))
            if not handle:
                return False       # non esiste piu' (o non si puo' guardare)
            try:
                codice_uscita = ctypes.c_ulong()
                if not k32.GetExitCodeProcess(handle, ctypes.byref(codice_uscita)):
                    return False
                return codice_uscita.value == 259      # STILL_ACTIVE
            finally:
                k32.CloseHandle(handle)
        except Exception:
            return False
    try:
        os.kill(int(pid), 0)
    except ProcessLookupError:
        return False
    except PermissionError:
        return True
    except OSError:
        return False
    return True


def _libera_blocco_orfano(pid_motore=None):
    """Cancella revmemo.lock SOLO se e' sicuramente rimasto orfano: contiene
    il PID del motore appena chiuso, un PID illeggibile, o il PID di un
    processo morto. Se invece il blocco appartiene a un altro invio ANCORA
    VIVO (es. attivita' pianificata partita mentre qui mancava la password)
    resta dov'e': cosi' nessun cliente riceve l'email doppia. Il battito da
    15 minuti del blocco ripulisce comunque i casi rimasti indietro."""
    percorso = os.path.join(BASE, "revmemo.lock")
    try:
        with open(percorso, "r", encoding="utf-8") as f:
            testo = (f.read() or "").strip()
        pid = int("".join(c for c in testo if c.isdigit()) or "0")
    except (OSError, ValueError):
        pid = 0
    if pid and pid != pid_motore and _pid_vivo(pid):
        return          # l'invio di un altro programma sta girando: non toccare
    try:
        os.remove(percorso)
    except OSError:
        pass


def _esegui(tipo, comando):
    codice = 1
    try:
        proc = subprocess.Popen(comando, cwd=BASE,
                                stdin=subprocess.PIPE if tipo == "controllo"
                                else subprocess.DEVNULL,
                                stdout=subprocess.PIPE,
                                stderr=subprocess.STDOUT,
                                text=True, encoding="utf-8",
                                errors="replace", env=ambiente_utf8(),
                                creationflags=_bandiere())
        with SERRATURE:
            PROCESSO["p"] = proc
        if tipo == "controllo":
            # controlla_pc.py chiede INVIO alla fine: glielo mando io
            try:
                proc.stdin.write("\n")
                proc.stdin.close()
            except Exception:
                pass
        for riga in proc.stdout:
            with SERRATURE:
                AZIONE["righe"].append(riga.rstrip("\n"))
        if tipo == "controllo":
            try:
                codice = proc.wait(timeout=300)
            except subprocess.TimeoutExpired:
                proc.terminate()
                codice = 1
                with SERRATURE:
                    AZIONE["righe"].append(
                        "Tempo scaduto (5 minuti): controllo fermato.")
        else:
            codice = proc.wait()
    except Exception as e:
        with SERRATURE:
            AZIONE["righe"].append(f"ERRORE pannello: {e}")
        codice = 1
    finally:
        if tipo == "invio" and codice not in (0, 2):
            # fermato o crollato: il motore non ha potuto liberare il suo
            # blocco anti-doppio-invio (codice 2 = blocco di ALTRO invio:
            # non e' nostro, non si tocca). Lo liberiamo solo se e' davvero
            # orfano: se appartiene a un invio ancora vivo (es. attivita'
            # pianificata partita mentre qui mancava la password) resta
            # dov'e', cosi' non si rischia il doppio invio.
            with SERRATURE:
                pid_motore = PROCESSO["p"].pid if PROCESSO["p"] else None
            try:
                _libera_blocco_orfano(pid_motore)
            except Exception:
                pass
        with SERRATURE:
            PROCESSO["p"] = None
            AZIONE["finito"] = True
            AZIONE["codice"] = codice
            AZIONE["in_corso"] = False
        nota(f"azione '{tipo}' finita con codice {codice}")


def ferma_azione():
    """pulsante FERMA: interrompe l'azione in corso (prova/invio)"""
    with SERRATURE:
        if not AZIONE["in_corso"]:
            return False, "Non c'e' nessuna azione in corso."
        proc = PROCESSO["p"]
        if proc is None:
            return False, "L'azione sta gia' finendo: aspetta un momento."
        AZIONE["fermata"] = True
        try:
            proc.terminate()
        except Exception as e:
            return False, f"Non riesco a fermarla: {e}"
        AZIONE["righe"].append("--- AZIONE FERMATA da te ---")
    return True, ""  # il processo muore: il pannello mostra FERMATA da solo


def apri_excel():
    if not os.path.exists(PERCORSO_EXCEL):
        return False, "Il file clienti non esiste ancora: fai una PROVA e lo creo io."
    try:
        if os.name == "nt":
            os.startfile(PERCORSO_EXCEL)          # noqa: solo Windows
        elif sys.platform == "darwin":
            subprocess.Popen(["open", PERCORSO_EXCEL])
        else:
            subprocess.Popen(["xdg-open", PERCORSO_EXCEL])
        return True, ""
    except Exception as e:
        return False, f"Non riesco ad aprire l'Excel: {e}"


# =====================================================================
#  GESTIONE CLIENTI (legge e scrive clienti_revisioni.xlsx)
# =====================================================================
ERRORE_EXCEL_APERTO = ("Il file Excel risulta APERTO: salvalo (Ctrl+S), "
                       "chiudi Excel, aspetta qualche secondo e riprova.")
ERRORE_EXCEL_OCCUPATO = ("Il file clienti e' occupato da un'altra scrittura "
                         "(forse un invio in corso?): riprova tra qualche "
                         "secondo.")
LIMITE_RIGA = 300          # come le tendine e i colori del file


@contextmanager
def _serratura_file_clienti():
    """La serratura CONDIVISA col motore (revmemo.SerraturaExcel): mentre
    un invio scrive gli STOP, il pannello non salva clienti e viceversa
    (prima la serratura interna qui sopra tiene in fila solo i thread del
    pannello). Rende '' se presa, il messaggio d'errore se occupata."""
    try:
        serratura = motore().SerraturaExcel(attesa=10)
    except Exception:
        yield ""                     # motore non disponibile: come prima
        return
    if not serratura.prendi():
        yield ERRORE_EXCEL_OCCUPATO
        return
    try:
        yield ""
    finally:
        serratura.rilascia()


def _valida_cliente(dati):
    """ritorna (campi_puliti, errore): nome/email/telefono/targa/scadenza"""
    nome = str(dati.get("nome") or "").strip()
    email = str(dati.get("email") or "").strip()
    telefono = str(dati.get("telefono") or "").strip()
    targa = str(dati.get("targa") or "").strip().upper()
    scadenza = str(dati.get("scadenza") or "").strip()
    if not nome:
        return None, "Manca il NOME del cliente."
    if not motore().email_valida(email):
        return None, "L'EMAIL non sembra giusta (esempio: mario.rossi@email.it)."
    if not targa:
        return None, "Manca la TARGA."
    if not scadenza:
        return None, "Manca la SCADENZA: scrivila come MESE/ANNO (es. 09/2026)."
    r = motore()
    data, _ = r.converti_scadenza(scadenza)
    if data is None:
        return None, ("SCADENZA non chiara: scrivila come MESE/ANNO "
                      "(es. 09/2026) o data completa (es. 30/09/2026).")
    if data.year < 2000:
        return None, "Nella SCADENZA scrivi l'anno per esteso (es. 09/2026)."
    # 3/2026 -> 03/2026 (solo il formato MESE/ANNO)
    pezzi = scadenza.replace("-", "/").split("/")
    if (len(pezzi) == 2 and pezzi[0].isdigit() and len(pezzi[1]) == 4
            and pezzi[1].isdigit() and 1 <= int(pezzi[0]) <= 12):
        scadenza = f"{int(pezzi[0]):02d}/{pezzi[1]}"
    return {"nome": nome, "email": email, "telefono": telefono,
            "targa": targa, "scadenza": scadenza}, ""


def _carica_per_scrittura():
    """(wb, ws, colonna_di, errore) - con copia di sicurezza automatica"""
    r = motore()
    if not os.path.exists(PERCORSO_EXCEL):
        return None, None, None, "Il file clienti non esiste ancora."
    try:
        r.backup_excel(PERCORSO_EXCEL)
    except Exception:
        pass
    try:
        wb = r.load_workbook(PERCORSO_EXCEL)   # senza data_only: resta tutto
    except PermissionError:
        return None, None, None, ERRORE_EXCEL_APERTO
    except Exception as e:
        return None, None, None, f"Non riesco ad aprire il file clienti: {e}"
    ws = wb[r.FOGLIO] if r.FOGLIO in wb.sheetnames else wb.active
    colonna_di, mancanti = r.individua_colonne(ws)
    if mancanti:
        return None, None, None, ("Nel file Excel mancano le colonne: "
                                  + ", ".join(mancanti)
                                  + " (le intestazioni stanno alla riga 4).")
    return wb, ws, colonna_di, ""


def _salva_excel(wb):
    try:
        motore().salva_excel_sicuro(wb, PERCORSO_EXCEL)
    except PermissionError:
        return ERRORE_EXCEL_APERTO
    except Exception as e:
        return f"Non riesco a salvare il file clienti: {e}"
    return ""


def _riga_piena(ws, colonna_di, riga):
    return any((ws.cell(row=riga, column=c).value not in (None, ""))
               for c in colonna_di.values())


def _riga_gia_con_targa(ws, colonna_di, targa, tranne_riga=0):
    r = motore()
    if "targa" not in colonna_di:
        return 0
    for row in range(r.RIGA_INTESTAZIONI + 1, LIMITE_RIGA + 1):
        if row == tranne_riga:
            continue
        v = ws.cell(row=row, column=colonna_di["targa"]).value
        if v and str(v).strip().upper() == targa:
            return row
    return 0


def _prima_riga_libera(ws, colonna_di):
    r = motore()
    chiavi = [k for k in ("email", "targa", "scadenza") if k in colonna_di]
    # le righe dentro una zona UNITA (es. la nota di istruzioni unita in
    # fondo al file di esempio) non si possono scrivere: vanno saltate,
    # il cliente nuovo va nella prima riga libera DOPO la nota
    unite = [(m.min_row, m.max_row) for m in ws.merged_cells.ranges]
    for row in range(r.RIGA_INTESTAZIONI + 1, LIMITE_RIGA + 1):
        if any(a <= row <= b for a, b in unite):
            continue
        if all((ws.cell(row=row, column=colonna_di[k]).value in (None, ""))
               for k in chiavi):
            return row
    return None


def _scrivi_campi_cliente(ws, colonna_di, riga, campi, fatta=None, attivo=None):
    # se sulla riga c'e' una zona UNITA (fatta a mano in Excel) la stacco
    # prima: openpyxl rifiuta la scrittura sulle celle unite
    for m in [m for m in ws.merged_cells.ranges
              if m.min_row <= riga <= m.max_row]:
        ws.unmerge_cells(str(m))
    def scrivi(chiave, valore):
        if chiave in colonna_di:
            ws.cell(row=riga, column=colonna_di[chiave], value=valore)
    scrivi("nome", campi["nome"])
    scrivi("email", campi["email"])
    scrivi("telefono", campi["telefono"] or None)
    scrivi("targa", campi["targa"])
    scrivi("scadenza", campi["scadenza"])
    ws.cell(row=riga, column=colonna_di["scadenza"]).number_format = "@"
    if fatta is not None:
        scrivi("fatta", fatta)
    if attivo is not None:
        scrivi("attivo", attivo)


def _attivo_visuale(testo, attivo_bool):
    """STOP / SI / NO per la visualizzazione: il motore accetta piu'
    scritture (S\u00cc, SI, S, TRUE, ...) e le celle possono essere vuote"""
    t = str(testo or "").strip().upper()
    if t == "STOP":
        return "STOP"
    if t in ("SI", "S\u00cc", "S", "TRUE", "VERO", "1", "YES", "Y"):
        return "SI"
    if t:
        return "NO"
    return "SI" if attivo_bool else "NO"


def lista_clienti():
    """tutti i clienti con il numero di riga (per la scheda Clienti)"""
    r = motore()
    if not os.path.exists(PERCORSO_EXCEL):
        return {"clienti": []}
    try:
        clienti = r.leggi_clienti(PERCORSO_EXCEL)
    except BaseException as e:          # anche SystemExit del motore
        msg = str(e).strip() or "controlla le intestazioni (riga 4)"
        return {"errore": f"Il file clienti ha un problema: {msg}"}
    testi = _testo_scadenze()
    fuori = [{"riga": c["riga"], "nome": c["nome"], "email": c["email"],
              "telefono": c["telefono"], "targa": c["targa"],
              "scadenza": testi.get(c["riga"])
              or (c["scadenza"].strftime("%d/%m/%Y") if c["scadenza"] else ""),
              "fatta": bool(c["fatta"]),
              "attivo": _attivo_visuale(c["attivo_testo"], c["attivo"])}
             for c in clienti]
    return {"clienti": fuori}


def nuovo_cliente(dati):
    campi, errore = _valida_cliente(dati)
    if errore:
        return {"errore": errore}
    r = motore()
    with SERRATURA_EXCEL, _serratura_file_clienti() as occupato:
        if occupato:
            return {"errore": occupato}
        creato = False
        if not os.path.exists(PERCORSO_EXCEL):
            try:
                r.crea_excel_se_manca(PERCORSO_EXCEL)
            except Exception as e:
                return {"errore": f"Non riesco a creare il file clienti: {e}"}
            creato = True
            # tolgo i clienti di esempio: file vuoto, pronto per il centro
            try:
                wbe = r.load_workbook(PERCORSO_EXCEL)
                wse = wbe[r.FOGLIO] if r.FOGLIO in wbe.sheetnames \
                    else wbe.active
                for row in range(r.RIGA_INTESTAZIONI + 1,
                                 r.RIGA_INTESTAZIONI + 9):
                    for col in range(2, 9):
                        wse.cell(row=row, column=col).value = None
                r.salva_excel_sicuro(wbe, PERCORSO_EXCEL)
            except Exception:
                pass
        wb, ws, colonna_di, errore = _carica_per_scrittura()
        if errore:
            return {"errore": errore}
        doppia = _riga_gia_con_targa(ws, colonna_di, campi["targa"])
        if doppia:
            return {"errore": f"La targa {campi['targa']} e' gia' nell'elenco "
                              f"(riga {doppia}): usa MODIFICA per correggerla."}
        riga = _prima_riga_libera(ws, colonna_di)
        if riga is None:
            return {"errore": ("Il file clienti e' pieno (300 righe): "
                               "aggiungi il cliente direttamente in Excel.")}
        _scrivi_campi_cliente(ws, colonna_di, riga, campi,
                              fatta="NO", attivo="S\u00cc")
        errore = _salva_excel(wb)
        if errore:
            return {"errore": errore}
        nota(f"cliente aggiunto: {campi['targa']} alla riga {riga}")
    return {"ok": True, "riga": riga, "creato": creato}


def modifica_cliente(dati):
    r = motore()
    try:
        riga = int(dati.get("riga") or 0)
    except (TypeError, ValueError):
        return {"errore": "Riga non valida."}
    if riga < r.RIGA_INTESTAZIONI + 1 or riga > LIMITE_RIGA:
        return {"errore": "Riga non valida (5-300)."}
    campi, errore = _valida_cliente(dati)
    if errore:
        return {"errore": errore}
    with SERRATURA_EXCEL, _serratura_file_clienti() as occupato:
        if occupato:
            return {"errore": occupato}
        wb, ws, colonna_di, errore = _carica_per_scrittura()
        if errore:
            return {"errore": errore}
        if not _riga_piena(ws, colonna_di, riga):
            return {"errore": f"La riga {riga} e' vuota: aggiorna l'elenco "
                              "(bottone su una riga piena)."}
        doppia = _riga_gia_con_targa(ws, colonna_di, campi["targa"],
                                     tranne_riga=riga)
        if doppia:
            return {"errore": f"La targa {campi['targa']} e' gia' usata "
                              f"alla riga {doppia}."}
        # scadenza NUOVA (diversa da prima) -> "Revisione fatta" torna NO
        # DA SOLO: il cliente e' in un giro di revisione nuovo e i
        # promemoria devono ripartire senza che nessuno si ricordi di
        # rimettere NO a mano. Stessa data (anche scritta diversa,
        # es. 09/2026 e 30/09/2026): niente cambia.
        fatta_reset = False
        if "fatta" in colonna_di:
            vecchia_data, _ = r.converti_scadenza(
                ws.cell(row=riga, column=colonna_di["scadenza"]).value)
            nuova_data, _ = r.converti_scadenza(campi["scadenza"])
            if nuova_data is not None and vecchia_data is not None \
                    and nuova_data != vecchia_data:
                fatta_reset = True
        _scrivi_campi_cliente(ws, colonna_di, riga, campi,
                              fatta="NO" if fatta_reset else None)
        errore = _salva_excel(wb)
        if errore:
            return {"errore": errore}
        nota(f"cliente modificato: riga {riga} ({campi['targa']})"
             + ("; scadenza nuova -> revisione fatta torna NO"
                if fatta_reset else ""))
    return {"ok": True, "fatta_reset": fatta_reset}


def elimina_cliente(dati):
    """TOGLIE il cliente dal file Excel: la riga sparisce e le righe
    sotto salgono di un posto (niente buchi). Le zone UNITA (note a
    fondo tabella) sotto la riga eliminata vengono spostate integre.
    La copia di sicurezza parte da sola con _carica_per_scrittura."""
    r = motore()
    try:
        riga = int(dati.get("riga") or 0)
    except (TypeError, ValueError):
        return {"errore": "Riga non valida."}
    if riga < r.RIGA_INTESTAZIONI + 1 or riga > LIMITE_RIGA:
        return {"errore": "Riga non valida (5-300)."}
    with SERRATURA_EXCEL, _serratura_file_clienti() as occupato:
        if occupato:
            return {"errore": occupato}
        wb, ws, colonna_di, errore = _carica_per_scrittura()
        if errore:
            return {"errore": errore}
        if not _riga_piena(ws, colonna_di, riga):
            return {"errore": f"La riga {riga} e' vuota: aggiorna l'elenco."}
        if any(m.min_row <= riga <= m.max_row
               for m in ws.merged_cells.ranges):
            return {"errore": f"La riga {riga} fa parte di una zona unita "
                              "(una nota): non e' un cliente, la lascio "
                              "stare."}
        nome = str(ws.cell(row=riga, column=colonna_di["nome"]).value or "")
        targa = str(ws.cell(row=riga, column=colonna_di["targa"]).value or "")
        # openpyxl con delete_rows NON sposta le zone unite da solo:
        # le stacco prima e rimetto quelle SOTTO la riga, un posto piu' su
        sotto = [(m.min_row, m.min_col, m.max_row, m.max_col)
                 for m in ws.merged_cells.ranges if m.min_row > riga]
        toccate = [(m.min_row, m.min_col, m.max_row, m.max_col)
                   for m in ws.merged_cells.ranges
                   if m.min_row <= riga <= m.max_row]
        for a, b, c, d in toccate + sotto:
            ws.unmerge_cells(start_row=a, start_column=b,
                             end_row=c, end_column=d)
        ws.delete_rows(riga, 1)
        for a, b, c, d in sotto:
            ws.merge_cells(start_row=a - 1, start_column=b,
                           end_row=c - 1, end_column=d)
        errore = _salva_excel(wb)
        if errore:
            return {"errore": errore}
        nota(f"cliente ELIMINATO: riga {riga} ({nome}, {targa})")
    return {"ok": True, "nome": nome, "targa": targa}


def stato_cliente(dati):
    """S\u00cc/NO su Revisione Fatta o Attivo (riattivare un STOP = S\u00cc:
    come da istruzioni del motore). STOP NON e' scrivibile da qui: lo
    scrive solo il programma quando il cliente lo chiede."""
    r = motore()
    try:
        riga = int(dati.get("riga") or 0)
    except (TypeError, ValueError):
        return {"errore": "Riga non valida."}
    campo = str(dati.get("campo") or "").strip()
    if campo not in ("fatta", "attivo"):
        return {"errore": "Campo non valido."}
    valore = str(dati.get("valore") or "").strip().upper()
    if valore in ("SI", "S\u00cc"):
        valore = "S\u00cc"
    elif valore == "NO":
        valore = "NO"
    else:
        return {"errore": "Valore non valido: ammessi solo S\u00cc e NO."}
    if riga < r.RIGA_INTESTAZIONI + 1 or riga > LIMITE_RIGA:
        return {"errore": "Riga non valida (5-300)."}
    with SERRATURA_EXCEL, _serratura_file_clienti() as occupato:
        if occupato:
            return {"errore": occupato}
        wb, ws, colonna_di, errore = _carica_per_scrittura()
        if errore:
            return {"errore": errore}
        if campo not in colonna_di:
            return {"errore": "Nel file Excel manca la colonna " + campo + "."}
        if not _riga_piena(ws, colonna_di, riga):
            return {"errore": f"La riga {riga} e' vuota: aggiorna l'elenco."}
        ws.cell(row=riga, column=colonna_di[campo], value=valore)
        errore = _salva_excel(wb)
        if errore:
            return {"errore": errore}
        nota(f"clienti: riga {riga} {campo} = {valore}")
    return {"ok": True, "valore": valore}


# =====================================================================
#  ANTEPRIMA VERA (usa il motore: stessa email che ricevera' il cliente)
# =====================================================================
def anteprima():
    r = motore()
    errore_config = ricarica_config(r)
    if errore_config:
        raise ValueError(errore_config)
    oggi = r.dt.date.today()
    cliente = {"nome": "Mario Rossi", "targa": "AB123CD",
               "scadenza": oggi + r.dt.timedelta(days=10)}
    oggetto, testo, html = r.componi_email(cliente, 10, 10)
    if not html:
        html = "<pre style='font-family:inherit'>" \
               + html_lib.escape(testo) + "</pre>"
    else:
        # nel preview il logo non puo' usare cid: (standard per le email):
        # mostro il file logo.png del pannello
        html = html.replace('src="cid:logo"', 'src="/logo.png"')
    return {"oggetto": oggetto, "html": html}


# =====================================================================
#  SERVER HTTP LOCALE (solo su questo computer)
# =====================================================================
class Gestore(BaseHTTPRequestHandler):
    def _invia(self, codice, corpo, tipo="application/json; charset=utf-8"):
        dati = corpo if isinstance(corpo, bytes) else corpo.encode("utf-8")
        self.send_response(codice)
        self.send_header("Content-Type", tipo)
        self.send_header("Content-Length", str(len(dati)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(dati)

    def _json(self, oggetto, codice=200):
        self._invia(codice, json.dumps(oggetto, ensure_ascii=False))

    def _permesso(self, percorso):
        """Solo la pagina del pannello (stesso indirizzo, con la chiave di
        sessione) puo' usare le API. Blocca le richieste partite da altri
        siti (CSRF) e quelle con un nome host falso (DNS rebinding)."""
        if not percorso.startswith("/api/") or percorso == "/api/ping":
            return True
        porta = self.server.server_address[1]
        validi = {f"127.0.0.1:{porta}", f"localhost:{porta}"}
        if (self.headers.get("Host") or "").strip().lower() not in validi:
            return False
        origine = (self.headers.get("Origin") or "").strip().lower()
        if origine and origine not in {f"http://{h}" for h in validi}:
            return False
        dato = self.headers.get("X-RevMemo-Token") or ""
        return secrets.compare_digest(dato, TOKEN)

    def _rifiuta(self):
        self._json({"errore": "Richiesta non autorizzata: riapri il pannello "
                              "dal collegamento (apri_pannello.bat)."}, 403)

    def _corpo_json(self):
        try:
            lunghezza = int(self.headers.get("Content-Length") or 0)
        except ValueError:
            return {}
        if lunghezza <= 0 or lunghezza > 2_000_000:
            return {}
        try:
            return json.loads(self.rfile.read(lunghezza).decode("utf-8"))
        except Exception:
            return {}

    def do_GET(self):
        percorso = urlparse(self.path).path
        try:
            if not self._permesso(percorso):
                return self._rifiuta()
            if percorso == "/api/ping":
                self._json({"revmemo": True})
            elif percorso in ("/", "/index.html"):
                with open(PERCORSO_HTML, "rb") as f:
                    pagina = f.read().decode("utf-8").replace(
                        "</head>", '<meta name="revmemo-token" content="'
                        + TOKEN + '"></head>', 1)
                    self._invia(200, pagina, "text/html; charset=utf-8")
            elif percorso in ("/logo.png", "/favicon.ico"):
                percorso_file = os.path.join(BASE, "logo.png")
                if os.path.exists(percorso_file):
                    with open(percorso_file, "rb") as f:
                        self._invia(200, f.read(), "image/png")
                else:
                    self._invia(404, "logo mancante", "text/plain")
            elif percorso == "/api/stato":
                self._json(dati_stato())
            elif percorso == "/api/clienti":
                self._json(lista_clienti())
            elif percorso == "/api/messaggio":
                self._json(leggi_config_raw())
            elif percorso == "/api/registro":
                self._json({"righe": leggi_registro(),
                            "statistiche": statistiche_invii()})
            elif percorso == "/api/azione-stato":
                with SERRATURE:
                    self._json(dict(AZIONE))
            elif percorso == "/api/anteprima":
                self._json(anteprima())
            else:
                self._invia(404, "non trovato", "text/plain")
        except BaseException as e:
            try:
                self._json({"errore": str(e) or "errore interno"}, 500)
            except Exception:
                pass

    def do_POST(self):
        percorso = urlparse(self.path).path
        if not self._permesso(percorso):
            return self._rifiuta()
        corpo = self._corpo_json()
        try:
            if percorso == "/api/messaggio":
                self._salva(corpo, ["messaggio", "centro"])
            elif percorso == "/api/impostazioni":
                self._salva(corpo, ["smtp", "centro", "prodotto"])
            elif percorso == "/api/anteprima":
                if corpo:
                    self._salva_senza_risposta(corpo, ["messaggio", "centro"])
                self._json(anteprima())
            elif percorso == "/api/azione":
                ok, messaggio = lancia_azione(
                    str(corpo.get("tipo", "")).strip())
                self._json({"ok": ok, "messaggio": messaggio},
                           200 if ok else 409)
            elif percorso == "/api/azione-ferma":
                ok, messaggio = ferma_azione()
                self._json({"ok": ok, "messaggio": messaggio},
                           200 if ok else 409)
            elif percorso == "/api/apri-excel":
                ok, messaggio = apri_excel()
                self._json({"ok": ok, "messaggio": messaggio})
            elif percorso == "/api/registro-excel":
                ok, messaggio, quanti = esporta_registro()
                self._json({"ok": ok, "messaggio": messaggio,
                            "quanti": quanti})
            elif percorso == "/api/cliente/nuovo":
                self._json(nuovo_cliente(corpo))
            elif percorso == "/api/cliente/modifica":
                self._json(modifica_cliente(corpo))
            elif percorso == "/api/cliente/elimina":
                self._json(elimina_cliente(corpo))
            elif percorso == "/api/cliente/stato":
                self._json(stato_cliente(corpo))
            else:
                self._invia(404, "non trovato", "text/plain")
        except BaseException as e:
            nota(f"errore POST {percorso}: {e}")
            try:
                self._json({"errore": str(e) or "errore interno"}, 500)
            except Exception:
                pass

    def _salva_senza_risposta(self, corpo, sezioni):
        # controlli semplici sulla casella email (server/porta): cosi' un
        # errore di battitura non rovina gli invii
        smtp = corpo.get("smtp")
        if isinstance(smtp, dict):
            host = str(smtp.get("host") or "").strip()
            porta = str(smtp.get("porta") or "").strip()
            if host and "." not in host:
                raise ValueError("Il SERVER non sembra giusto (esempio: "
                                 "smtp.libero.it): controlla e riprova.")
            if porta and (not porta.isdigit() or not (1 <= int(porta) <= 65535)):
                raise ValueError("La PORTA deve essere un numero (di solito "
                                 "465 oppure 587): controlla e riprova.")
            if host and not porta:
                smtp["porta"] = "465"     # default se si mette solo il server
        modifiche = {}
        for sez in sezioni:
            dati = corpo.get(sez)
            if isinstance(dati, dict):
                puliti = {k: (v if isinstance(v, str) else None)
                          for k, v in dati.items()
                          if k in CHIAVI.get(sez, [])}
                # password vuota = NON toccare quella salvata
                if sez == "smtp" and puliti.get("password_app", None) == "":
                    puliti["password_app"] = None
                if sez == "smtp":
                    # "risposta_a" non esiste piu' nel pannello: confondeva
                    # (se compilato, lo STOP automatico non leggeva piu' le
                    # risposte). Le risposte devono arrivare nella casella
                    # mittente: se una vecchia config lo aveva impostato,
                    # lo svuoto qui ad ogni salvataggio.
                    puliti["risposta_a"] = ""
                if puliti:
                    modifiche[sez] = puliti
        if modifiche:
            salva_campi(modifiche)

    def _salva(self, corpo, sezioni):
        self._salva_senza_risposta(corpo, sezioni)
        self._json({"ok": True})

    def log_message(self, fmt, *args):
        pass          # niente log in console (il pannello gira silenzioso)


def _pannello_nostro(porta):
    """sulla porta gira gia' un pannello RevMemo? (se il doppio clic
    parte due volte, basta riaprire la pagina gia' attiva)"""
    try:
        import urllib.request
        with urllib.request.urlopen(
                f"http://127.0.0.1:{porta}/api/ping", timeout=2) as r:
            dati = json.loads(r.read().decode("utf-8"))
        return isinstance(dati, dict) and dati.get("revmemo") is True
    except Exception:
        return False


def _apri_browser(url):
    """apre la pagina; se il browser non parte, la scritta dice cosa fare"""
    try:
        if webbrowser.open(url):
            return True
    except Exception:
        pass
    print()
    print("=" * 62)
    print("NON SONO RIUSCITO AD APRIRE IL BROWSER DA SOLO.")
    print("Apri tu questo indirizzo (copia e incolla nella barra")
    print("indirizzi di Chrome o Edge):")
    print(url)
    print("=" * 62)
    return False


def main():
    nota(f"avvio pannello (python {sys.version.split()[0]}, {sys.platform})")
    # crea config.ini se manca (il motore lo fa già al primo uso)
    try:
        motore()
    except (SystemExit, RuntimeError):
        pass          # config.ini rotto: il pannello parte lo stesso e lo spiega
    # porta libera? se occupata: pannello gia' aperto -> riapro la pagina;
    # altrimenti provo le porte dopo (un altro programma puo' usarla)
    server = None
    for porta in (PORTA, PORTA + 1, PORTA + 2, PORTA + 3, PORTA + 4):
        if _pannello_nostro(porta):
            nota(f"pannello gia' attivo sulla porta {porta}: riapro la pagina")
            print("Il pannello era gia' aperto: te lo faccio vedere nel "
                  "browser.")
            _apri_browser(f"http://127.0.0.1:{porta}/")
            return 0
        try:
            server = ThreadingHTTPServer(("127.0.0.1", porta), Gestore)
            porta_usata = porta
            break
        except OSError:
            nota(f"porta {porta} occupata da altro programma")
    if server is None:
        print(f"PROBLEMA: nessuna porta libera tra {PORTA} e {PORTA + 4}.")
        print("Riavvia il computer e riprova; se continua, chi ti assiste")
        print("sa cosa guardare nel file pannello_log.txt")
        return 1
    url = f"http://127.0.0.1:{porta_usata}/"
    threading.Timer(0.8, lambda: _apri_browser(url)).start()
    nota(f"pannello attivo su {url}")
    print(f"RevMemo - pannello attivo su {url}")
    print("Lascia aperta questa finestra mentre usi il pannello.")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    return 0


if __name__ == "__main__":
    try:
        CODICE_USCITA = main()
    except SystemExit:
        raise
    except BaseException:
        import traceback
        ERRORE = traceback.format_exc()
        nota("ERRORE AVVIO: " + ERRORE.replace("\n", " | ")[:1500])
        print()
        print("=" * 62)
        print("ERRORE ALL'AVVIO DEL PANNELLO")
        print("=" * 62)
        print(ERRORE)
        print()
        print("Fai una FOTO di questa finestra e inviala a chi ti assiste.")
        print("(Aiuta anche: doppio click su controlla_pannello.bat)")
        try:
            input("Premi INVIO per chiudere...")
        except Exception:
            pass
        CODICE_USCITA = 1
    sys.exit(CODICE_USCITA)
