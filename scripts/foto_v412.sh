#!/bin/bash
# Foto del pannello RevMemo v4.12: bottone ELIMINA con conferma interna,
# eliminazione reale (righe sotto che salgono) e reset automatico di
# "Revisione fatta" al salvataggio di una scadenza nuova.
set -e
PORTE=8768
FOTO=/home/z/my-project/download/foto_pannello
DEMO=/tmp/demo_pannello_v412

mkdir -p "$DEMO" "$FOTO"
cp /home/z/my-project/download/RevMemo/pannello.py \
   /home/z/my-project/download/RevMemo/pannello.html \
   /home/z/my-project/download/RevMemo/revmemo.py \
   /home/z/my-project/download/RevMemo/controlla_pc.py \
   /home/z/my-project/download/RevMemo/template_email.html \
   /home/z/my-project/download/RevMemo/logo.png \
   /home/z/my-project/download/RevMemo/clienti_revisioni.xlsx \
   /home/z/my-project/download/RevMemo/config.ini "$DEMO/"
rm -f "$DEMO/reminders_log.txt"

pkill -f "pannello.py" 2>/dev/null || true
sleep 1
cd "$DEMO"
PANNELLO_DEMO=1 PANNELLO_PORTA=$PORTE nohup python3 pannello.py > server.log 2>&1 &
PID=$!
for i in $(seq 1 20); do
  if curl -s -m 2 "http://127.0.0.1:$PORTE/api/stato" > /dev/null; then break; fi
  sleep 0.5
done
echo "demo attiva (pid $PID)"

agent-browser open "http://127.0.0.1:$PORTE/" --timeout 30000 > /dev/null 2>&1 || true
sleep 3

# ---- 1) CLIENTI: la riga 5 ha i tre bottoni (Modifica / Elimina)
agent-browser eval "document.querySelector('nav button[data-vista=clienti]').click()" > /dev/null 2>&1 || true
sleep 1.5
echo "TEST bottoni eliminina presenti:"; agent-browser eval "document.querySelectorAll('[data-elimina]').length + ' bottoni Elimina'"
agent-browser screenshot "$FOTO/v412_1_clienti.png" 2>&1 | tail -1

# ---- 2) ELIMINA: conferma interna con nome e targa
echo "TEST clienti prima:"; agent-browser eval "document.querySelectorAll('[data-elimina]').length"
agent-browser eval "document.querySelector('[data-elimina=\"5\"]').click()" > /dev/null 2>&1 || true
sleep 1
echo "TEST domanda di conferma:"; agent-browser eval "document.getElementById('conferma-testo').textContent"
echo "TEST tasto verde:"; agent-browser eval "document.getElementById('conferma-ok').textContent"
agent-browser screenshot "$FOTO/v412_2_conferma_elimina.png" 2>&1 | tail -1

# ---- 3) SÌ, ELIMINA: il cliente sparisce e gli altri salgono
agent-browser eval "document.getElementById('conferma-ok').click()" > /dev/null 2>&1 || true
sleep 1.5
echo "TEST dopo elimina:"; agent-browser eval "document.querySelectorAll('[data-elimina]').length + ' bottoni (prima 5)'"
echo "TEST avviso:"; agent-browser eval "document.getElementById('avviso').textContent"
agent-browser screenshot "$FOTO/v412_3_eliminato.png" 2>&1 | tail -1

# ---- 4) RESET AUTOMATICO: fatta=SÌ poi scadenza nuova -> torna NO da sola
agent-browser eval "document.querySelector('[data-fatta=\"6\"]').click()" > /dev/null 2>&1 || true
sleep 0.8
agent-browser eval "document.getElementById('conferma-ok').click()" > /dev/null 2>&1 || true
sleep 1.5
echo "TEST fatta=SÌ messa:"; agent-browser eval "document.querySelector('[data-fatta=\"6\"]') ? document.querySelector('[data-fatta=\"6\"]').textContent : 'manca'"
agent-browser eval "document.querySelector('[data-modifica=\"6\"]').click()" > /dev/null 2>&1 || true
sleep 0.6
agent-browser eval "const c=document.getElementById('c-scadenza'); c.value='12/2028'; 'ok'" > /dev/null 2>&1 || true
agent-browser eval "document.getElementById('b-salva-cliente').click()" > /dev/null 2>&1 || true
sleep 1.2
echo "TEST avviso reset (scadenza nuova):"; agent-browser eval "document.getElementById('avviso').textContent"
echo "TEST badge fatta della riga 6 (deve essere NO):"; agent-browser eval "document.querySelector('[data-fatta=\"6\"]').textContent"
agent-browser screenshot "$FOTO/v412_4_reset_fatta.png" 2>&1 | tail -1

kill $PID 2>/dev/null || true
pkill -f "pannello.py" 2>/dev/null || true
echo "foto v4.12 scattate"
ls "$FOTO" | grep v412 || true
