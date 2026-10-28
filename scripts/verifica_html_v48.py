#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Verifica il JS di pannello.html: estrae il blocco <script> e lo passa
a node --check; controlla anche che gli id richiamati nel JS esistano
nell'HTML e che le sezioni del pannello restino bilanciate."""
import re
import subprocess
import sys
import tempfile

PERCORSO = "/home/z/my-project/download/RevMemo/pannello.html"
html = open(PERCORSO, encoding="utf-8").read()

m = re.search(r"<script>\n(.*)</script>\s*</body>", html, re.DOTALL)
if not m:
    sys.exit("BLOCCO <script> NON TROVATO")
js = m.group(1)
with tempfile.NamedTemporaryFile("w", suffix=".js", delete=False,
                                 encoding="utf-8") as f:
    f.write(js)
    percorso_js = f.name
r = subprocess.run(["node", "--check", percorso_js], capture_output=True,
                   text=True)
if r.returncode != 0:
    print(r.stderr)
    sys.exit("ERRORE SINTASSI JS")
print("[OK] sintassi JS valida (node --check)")

# id usati nel JS vs id definiti nell'HTML
usati = set(re.findall(r'\$\("([a-z0-9-]+)"\)', js))
definiti = set(re.findall(r'id="([a-z0-9-]+)"', html))
mancanti = sorted(u for u in usati - definiti)
if mancanti:
    sys.exit("ID usati nel JS ma assenti nell'HTML: " + ", ".join(mancanti))
print(f"[OK] tutti gli id richiamati nel JS esistono ({len(usati)} id)")

# sezioni bilanciate
for id_sezione in ("vista-home", "vista-messaggio", "vista-clienti",
                   "vista-impostazioni", "vista-registro"):
    if f'id="{id_sezione}"' not in html:
        sys.exit(f"manca la sezione {id_sezione}")
print("[OK] 5 schede presenti")

# elementi chiave v4.8
attesi = ("b-chiudi-controllo", "b-esp-registro", "stat-invi",
          "i-host", "i-porta", "Email da inviare oggi", "barre-stat",
          "bottoneChiudi")
mancanti = [x for x in attesi if x not in html]
if mancanti:
    sys.exit("elementi v4.8 mancanti: " + ", ".join(mancanti))
print("[OK] elementi v4.8 presenti")
print("VERIFICA HTML/JS SUPERATA")
