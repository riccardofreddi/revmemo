#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Prova veloce di _pid_vivo / _libera_blocco_orfano (pannello.py v4.13).
Il blocco va cancellato SOLO se orfano (PID morto, nostro, o illeggibile);
non va toccato se appartiene a un processo vivo di qualcun altro."""
import importlib.util
import os
import subprocess
import sys
import tempfile

PANNELLO = "/home/z/my-project/download/RevMemo/pannello.py"
spec = importlib.util.spec_from_file_location("pannello", PANNELLO)
p = importlib.util.module_from_spec(spec)
spec.loader.exec_module(p)

cart = tempfile.mkdtemp(prefix="blocco_orfano_")
p.BASE = cart
percorso = os.path.join(cart, "revmemo.lock")
ok, ko = 0, []


def check(nome, cond):
    global ok
    if cond:
        ok += 1
        print(f"  [OK]   {nome}")
    else:
        ko.append(nome)
        print(f"  [FALL] {nome}")


# 1) pid vivo (questo script)
check("_pid_vivo(se stesso) = True", p._pid_vivo(os.getpid()) is True)
check("_pid_vivo(0) = False", p._pid_vivo(0) is False)
check("_pid_vivo(None) = False", p._pid_vivo(None) is False)

# 2) pid morto: un figlio che esce subito
figlio = subprocess.Popen([sys.executable, "-c", "pass"])
figlio.wait()
morto = figlio.pid
check("_pid_vivo(pid appena morto) = False", p._pid_vivo(morto) is False)

# 3) blocco di un processo MORTO e di nessuno nostro -> lo cancella
open(percorso, "w").write(str(morto))
p._libera_blocco_orfano(pid_motore=None)
check("blocco di pid morto: cancellato", not os.path.exists(percorso))

# 4) blocco di un processo VIVO di qualcun altro -> NON toccato
open(percorso, "w").write(str(os.getpid()))
p._libera_blocco_orfano(pid_motore=None)
check("blocco di pid vivo (altro): resta", os.path.exists(percorso))

# 5) blocco con il NOSTRO pid di motore (figlio appena uscito) -> cancellato
p._libera_blocco_orfano(pid_motore=os.getpid())
check("blocco dello stesso motore appena chiuso: cancellato",
      not os.path.exists(percorso))

# 6) blocco illeggibile/vuoto -> cancellato (orfano presunto)
open(percorso, "w").write("   ")
p._libera_blocco_orfano(pid_motore=None)
check("blocco illeggibile: cancellato", not os.path.exists(percorso))

# 7) blocco mancate -> nessun errore
p._libera_blocco_orfano(pid_motore=None)
check("blocco assente: nessun errore", True)

print(f"\nRISULTATO blocco orfano: {ok} OK, {len(ko)} FALLITI {ko if ko else ''}")
sys.exit(1 if ko else 0)
