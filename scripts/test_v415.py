#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Prove della v4.15: ricerca delle risposte STOP allargata a 35 giorni
(prima 7: a programma fermo oltre una settimana le richieste sfuggivano)
e pulsante 'CONTROLLA LE RISPOSTE STOP' del pannello (--solo-risposte:
registra gli STOP senza mandare niente)."""
import importlib.util
import inspect
import os
import shutil
import subprocess
import sys
import tempfile

PACCHETTO = os.environ.get("REVMEMO_PKG", "/home/z/my-project/download/RevMemo")
PANNELLO = os.path.join(PACCHETTO, "pannello.py")
MOTORE = os.path.join(PACCHETTO, "revmemo.py")
CONFIG_TEMPLATE = "/home/z/my-project/revmemo-github/revmemo/config.ini"

spec = importlib.util.spec_from_file_location("pannello", PANNELLO)
p = importlib.util.module_from_spec(spec)
spec.loader.exec_module(p)
r = p.motore()                       # il motore importato dal pannello

testo = open(MOTORE, encoding="utf-8").read()
testo_p = open(PANNELLO, encoding="utf-8").read()
html = open(os.path.join(PACCHETTO, "pannello.html"), encoding="utf-8").read()
leggimi = open(os.path.join(PACCHETTO, "LEGGIMI.txt"), encoding="utf-8").read()

ok, ko = 0, []


def check(nome, cond, extra=""):
    global ok
    if cond:
        ok += 1
        print(f"  [OK]   {nome}")
    else:
        ko.append(nome)
        print(f"  [FALL] {nome}  {extra}")


print("== A) Versione 4.15 ==")
check("VERSIONE = 4.15", getattr(r, "VERSIONE", "") == "4.15")
check("intestazione 4.15 (PRO)", "versione 4.15 (PRO)" in testo)
check("niente piu' intestazione 4.14 (PRO)",
      "versione 4.14 (PRO)" not in testo)

print("== B) Ricerca risposte STOP su 35 giorni ==")
check("GIORNI_DI_RICERCA = 35", r.GIORNI_DI_RICERCA == 35)
check("il commento spiega il cambio (prima 7)", "prima 7" in testo)
check("nessun rimasto 'ultimi 7 giorni' nel motore",
      "ultimi 7 giorni" not in testo)
check("LEGGIMI parla di 35 giorni", "35 giorni" in leggimi)
check("LEGGIMI ha la sezione NOVITA' v4.15", "NOVITA' v4.15" in leggimi)

print("== C) controlla_risposte: modalita' 'fuori invio' ==")
pars = inspect.signature(r.controlla_risposte).parameters
check("parametro in_invio esiste", "in_invio" in pars)
check("in_invio True di default (il giro di invio non cambia)",
      pars["in_invio"].default is True)
check("avviso dedicato per il pulsante (riprova, non 'proseguo')",
      "Riprova tra qualche minuto dal pannello" in testo)
check("il giro di invio chiama ancora controlla_risposte com'era",
      "controlla_risposte(percorso_excel, args.dry_run)" in testo)

print("== D) --solo-risposte: niente invii, niente creazioni ==")
cart = tempfile.mkdtemp(prefix="revmemo_v415_")
shutil.copy(MOTORE, os.path.join(cart, "revmemo.py"))
shutil.copy(CONFIG_TEMPLATE, os.path.join(cart, "config.ini"))
PERCORSO_EXCEL_PROVA = os.path.join(cart, "clienti_revisioni.xlsx")


def avvia(argomenti):
    esito = subprocess.run(
        [sys.executable, os.path.join(cart, "revmemo.py")] + argomenti,
        cwd=cart, capture_output=True, text=True, timeout=180)
    return esito.returncode, (esito.stdout or "") + (esito.stderr or "")


# D1: file clienti mancante -> messaggio chiaro e NESSUN file di esempio
codice, fuori = avvia(["--solo-risposte"])
check("file mancante: lo dice chiaro", "Non trovo il file clienti" in fuori)
check("file mancante: NON crea il file di esempio",
      not os.path.exists(PERCORSO_EXCEL_PROVA))
check("file mancante: esce con codice 1 come previsto", codice == 1,
      f"codice={codice}")

# D2: file presente, senza password -> lo dice e non tocca la rete
from openpyxl import Workbook
Workbook().save(PERCORSO_EXCEL_PROVA)
codice, fuori = avvia(["--solo-risposte"])
check("senza password: messaggio chiaro",
      "Per leggere le risposte servono mittente e password per app" in fuori)
check("senza password: nessun riepilogo di invii",
      "RIEPILOGO:" not in fuori)

# D3: blocco invio di un processo VIVO -> non tocca niente e lo dice
open(os.path.join(cart, "revmemo.lock"), "w").write(str(os.getpid()))
codice, fuori = avvia(["--solo-risposte"])
check("blocco di un invio vivo: rispettato", "Un altro invio sta girando" in fuori)
check("blocco di un invio vivo: esce tranquillo (codice 0)", codice == 0,
      f"codice={codice}")
check("blocco di un invio vivo: NON riscritto da noi",
      open(os.path.join(cart, "revmemo.lock")).read() == str(os.getpid()))
os.remove(os.path.join(cart, "revmemo.lock"))

# D4: --help elenca l'opzione nuova
codice, fuori = avvia(["--help"])
check("--help mostra --solo-risposte", "--solo-risposte" in fuori)

print("== E) Pannello agganciato al pulsante nuovo ==")
check("il pannello avvia il motore con --solo-risposte",
      '"--solo-risposte"' in testo_p)
check("il tipo 'risposte' e' nei comandi", '"risposte":' in testo_p)
check("il guard del file clienti copre anche 'risposte'",
      'tipo in ("invio", "risposte")' in testo_p)
check("pulsante CONTROLLA LE RISPOSTE STOP presente",
      'id="b-risposte"' in html and "CONTROLLA LE RISPOSTE STOP" in html)
check("il pulsante e' agganciato all'azione senza conferma (niente invii)",
      '$("b-risposte").addEventListener("click", () =>\n  azione("risposte"));' in html)
riga_bottoni = [l for l in html.splitlines() if "BOTTONI_AZIONE" in l]
check("BOTTONI_AZIONE include b-risposte (si disattiva durante le azioni)",
      bool(riga_bottoni) and "b-risposte" in riga_bottoni[0])
check("la card Azioni spiega il pulsante nuovo",
      "senza mandare niente" in html)

print(f"\nRISULTATO v4.15: {ok} OK, {len(ko)} FALLITI {ko if ko else ''}")
sys.exit(1 if ko else 0)
