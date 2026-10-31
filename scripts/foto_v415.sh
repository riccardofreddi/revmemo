#!/bin/bash
# Foto della v4.15: Home del pannello con il pulsante nuovo
# "CONTROLLA LE RISPOSTE STOP" nella card Azioni.
set -e
PORTE=8769
FOTO=/home/z/my-project/download/foto_pannello
DEMO=/tmp/demo_pannello_v415

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
echo "TEST pulsante nuovo visibile:"
agent-browser eval "const b=document.getElementById('b-risposte'); b ? 'OK: ' + b.textContent.trim() : 'MANCA'"
agent-browser eval "document.querySelector('#b-risposte').closest('.carta').scrollIntoView({block:'center'}); 'scorri'" > /dev/null 2>&1 || true
sleep 1
agent-browser screenshot "$FOTO/v415_1_home_pulsante_risposte.png" 2>&1 | tail -1

kill $PID 2>/dev/null || true
pkill -f "pannello.py" 2>/dev/null || true
echo "foto v4.15 scattata"
ls "$FOTO" | grep v415 || true
