#!/bin/bash
# Foto del pannello RevMemo v4.10: verifica visiva di scroll 10 in Home,
# menu VEDI 20/30/50 RIGHE (Clienti + Registro), label "Oggetto — 30 giorni".
set -e
IP=$(hostname -I | awk '{print $1}')
PORTE=8768
FOTO=/home/z/my-project/download/foto_pannello
DEMO=/tmp/demo_pannello_v410

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
# la demo resta in ascolto su 127.0.0.1: il browser apre quello

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

# 28 clienti demo in piu' (scadenze scalari) + 30 righe di registro
python3 - <<'EOF'
import json, urllib.request, datetime as dt
URL = "http://127.0.0.1:8768/api/cliente/nuovo"
oggi = dt.date.today()
nomi = ["Luca Moretti","Sara Conti","Marco Greco","Elena Fontana",
        "Davide Riva","Giulia Serra","Alberto Noce","Chiara Valli",
        "Franco Mansi","Laura Dotti","Sergio Balestri","Marta Guzzo",
        "Paolo Fumagalli","Silvia Ottavi","Roberto Lenti","Anna Peruzzi",
        "Nicola Baiardi","Elisa Roversi","Gino Spada","Carla Mucci",
        "Leo Bianchi","Mia Caruso","Tommaso Gelli","Irene Sarti",
        "Fabio Milani","Rosa Calo'","Sandro Pellei","Vera Naldini"]
for i, nome in enumerate(nomi):
    scad = (oggi + dt.timedelta(days=15 + i * 11)).strftime("%d/%m/%Y")
    corpo = json.dumps({"nome": nome, "email": f"cliente{i+1}@test.it",
                        "telefono": f"34{9}{i:02d} 555{1000+i}",
                        "targa": f"ZV{i:03d}QW", "scadenza": scad}).encode()
    req = urllib.request.Request(URL, data=corpo,
                                 headers={"Content-Type": "application/json"})
    try:
        urllib.request.urlopen(req, timeout=10).read()
    except Exception as e:
        print("salto", nome, e)
righe = []
for i in range(30):
    giorno = (dt.date(2026, 10, 2) + dt.timedelta(days=i)).isoformat()
    righe.append(f"{giorno};cliente{i+1}@test.it;DX{i:03d}RS;"
                 f"{[30,10,1][i%3]};{giorno};promemoria ({[30,10,1][i%3]} gg)")
with open("reminders_log.txt", "w", encoding="utf-8") as f:
    f.write("\n".join(righe) + "\n")
print("clienti demo aggiunti + log scritto")
EOF

agent-browser open "http://127.0.0.1:$PORTE/" --timeout 30000 > /dev/null 2>&1 || true
sleep 3
agent-browser screenshot "$FOTO/v410_1_home.png" 2>&1 | tail -1

agent-browser eval "document.querySelector('nav button[data-vista=clienti]').click()" > /dev/null 2>&1 || true
sleep 1.5
agent-browser screenshot "$FOTO/v410_2_clienti.png" 2>&1 | tail -1
# scroll DENTRO la tabella clienti: l'intestazione deve restare fissa
agent-browser eval "const c=document.getElementById('tabella-clienti'); c.scrollTop=520; 'ok'" > /dev/null 2>&1 || true
sleep 0.8
agent-browser screenshot "$FOTO/v410_3_clienti_scroll.png" 2>&1 | tail -1

agent-browser eval "document.querySelector('nav button[data-vista=registro]').click()" > /dev/null 2>&1 || true
sleep 1.5
agent-browser screenshot "$FOTO/v410_4_registro.png" 2>&1 | tail -1

agent-browser eval "document.querySelector('nav button[data-vista=messaggio]').click()" > /dev/null 2>&1 || true
sleep 1.5
agent-browser screenshot "$FOTO/v410_5_messaggio.png" 2>&1 | tail -1

kill $PID 2>/dev/null || true
pkill -f "pannello.py" 2>/dev/null || true
echo "foto v4.10 scattate"
ls "$FOTO" | grep v410
