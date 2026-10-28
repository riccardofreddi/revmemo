#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Mini server di download per RevMemo.
Serve la cartella /home/z/my-project/download sulla porta 3000
(il gateway del sistema la espone con l'anteprima web).
La pagina iniziale e' in italiano semplice con il bottone di download.
"""
import os
from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler

CARTELLA = "/home/z/my-project/download"
PORTA = 3000

PAGINA = """<!DOCTYPE html>
<html lang="it">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Download RevMemo</title>
<style>
  * { box-sizing: border-box; }
  body {
    margin: 0; min-height: 100vh; display: flex;
    align-items: center; justify-content: center;
    font-family: Arial, Helvetica, sans-serif;
    background: #0f2b46; color: #ffffff; padding: 16px;
  }
  .card {
    background: #ffffff; color: #16324f;
    border-radius: 14px; padding: 34px 28px;
    max-width: 430px; width: 100%;
    box-shadow: 0 12px 40px rgba(0,0,0,.35);
    text-align: center;
  }
  .logo { font-size: 44px; line-height: 1; }
  h1 { font-size: 26px; margin: 12px 0 6px; color: #0f2b46; }
  .sub { color: #5a7184; font-size: 15px; margin: 0 0 22px; }
  a.btn {
    display: block; background: #1a9e4b; color: #fff;
    text-decoration: none; font-size: 19px; font-weight: bold;
    padding: 16px 18px; border-radius: 10px; letter-spacing: .3px;
  }
  a.btn:active { background: #178a42; }
  .meta { color: #7a8c9c; font-size: 13px; margin-top: 10px; }
  .istruzioni {
    text-align: left; background: #f2f7fb; border-radius: 10px;
    padding: 14px 16px; margin-top: 22px; font-size: 14px; color: #33475b;
  }
  .istruzioni b { color: #0f2b46; }
  .istruzioni ol { margin: 8px 0 0; padding-left: 20px; }
  .istruzioni li { margin: 6px 0; }
</style>
</head>
<body>
  <div class="card">
    <div class="logo">📦</div>
    <h1>RevMemo v4.4</h1>
    <p class="sub">Pacchetto completo &mdash; pronto da installare</p>
    <a class="btn" href="/RevMemo_v4.zip" download>⬇️ SCARICA RevMemo_v4.zip</a>
    <p class="meta">10.9 MB &bull; 249 file &bull; verificato OK</p>
    <div class="istruzioni">
      <b>Dopo il download:</b>
      <ol>
        <li>Estrai lo ZIP sul PC (tasto destro &rarr; Estrai tutto)</li>
        <li>Apri la cartella estratta</li>
        <li>Fai doppio clic su <b>controlla_pc.bat</b> (controllo automatico)</li>
        <li>Poi <b>avvia_prova.bat</b> per la prova senza invii veri</li>
      </ol>
    </div>
  </div>
</body>
</html>
"""


class Handler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=CARTELLA, **kwargs)

    def do_GET(self):
        if self.path in ("/", "/index.html"):
            corpo = PAGINA.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(corpo)))
            self.end_headers()
            self.wfile.write(corpo)
        else:
            super().do_GET()

    def log_message(self, fmt, *args):
        # log essenziale su stdout (va nel file di log del server)
        print("[%s] %s" % (self.log_date_time_string(), fmt % args), flush=True)


if __name__ == "__main__":
    os.chdir(CARTELLA)
    server = ThreadingHTTPServer(("0.0.0.0", PORTA), Handler)
    print("Server download attivo su porta %d, cartella %s" % (PORTA, CARTELLA), flush=True)
    server.serve_forever()
