#!/bin/bash
# Foto del pannello RevMemo (v3 con gestione clienti): click via eval.
set -e
IP=$(hostname -I | awk '{print $1}')
PORTE=8768
FOTO=/home/z/my-project/download/foto_pannello
DEMO=/tmp/demo_pannello

# aggiorna la copia demo con i file correnti e il bind dimostrativo
mkdir -p "$DEMO"
cp /home/z/my-project/download/RevMemo/pannello.py \
   /home/z/my-project/download/RevMemo/pannello.html \
   /home/z/my-project/download/RevMemo/revmemo.py \
   /home/z/my-project/download/RevMemo/controlla_pc.py \
   /home/z/my-project/download/RevMemo/template_email.html \
   /home/z/my-project/download/RevMemo/logo.png \
   /home/z/my-project/download/RevMemo/clienti_revisioni.xlsx \
   /home/z/my-project/download/RevMemo/config.ini "$DEMO/"
rm -f "$DEMO/reminders_log.txt"
sed -i 's/("127.0.0.1", PORTA)/(("0.0.0.0" if os.environ.get("PANNELLO_DEMO") else "127.0.0.1"), PORTA)/' "$DEMO/pannello.py"

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

agent-browser open "http://$IP:$PORTE/" --timeout 30000 > /dev/null 2>&1 || true
sleep 2.5
agent-browser screenshot "$FOTO/1_home.png" 2>&1 | tail -1

agent-browser eval "document.querySelector('nav button[data-vista=messaggio]').click()" > /dev/null 2>&1 || true
sleep 1.5
agent-browser screenshot "$FOTO/2_messaggio.png" 2>&1 | tail -1

agent-browser eval "document.getElementById('b-anteprima').click()" > /dev/null 2>&1 || true
sleep 3
agent-browser screenshot "$FOTO/3_anteprima.png" 2>&1 | tail -1
agent-browser eval "document.getElementById('b-chiudi').click()" > /dev/null 2>&1 || true
sleep 1

agent-browser eval "document.querySelector('nav button[data-vista=clienti]').click()" > /dev/null 2>&1 || true
sleep 1.5
agent-browser screenshot "$FOTO/4_clienti.png" 2>&1 | tail -1

# NUOVA FOTO: modale NUOVO CLIENTE aperta e compilata
agent-browser eval "document.getElementById('b-nuovo-cliente').click()" > /dev/null 2>&1 || true
sleep 1
agent-browser eval "document.getElementById('c-nome').value='Luca Moretti'; document.getElementById('c-targa').value='QR567ST'; document.getElementById('c-email').value='luca.moretti@email.it'; document.getElementById('c-telefono').value='349 5550123'; document.getElementById('c-scadenza').value='03/2027'" > /dev/null 2>&1 || true
sleep 0.8
agent-browser screenshot "$FOTO/7_nuovo_cliente.png" 2>&1 | tail -1
agent-browser eval "document.getElementById('b-annulla-cliente').click()" > /dev/null 2>&1 || true
sleep 0.5

agent-browser eval "document.querySelector('nav button[data-vista=impostazioni]').click()" > /dev/null 2>&1 || true
sleep 1.5
agent-browser screenshot "$FOTO/5_impostazioni.png" 2>&1 | tail -1

agent-browser eval "document.querySelector('nav button[data-vista=registro]').click()" > /dev/null 2>&1 || true
sleep 1.5
agent-browser screenshot "$FOTO/6_registro.png" 2>&1 | tail -1

kill $PID 2>/dev/null || true
pkill -f "pannello.py" 2>/dev/null || true
echo "foto scattate:"
ls -la "$FOTO"
