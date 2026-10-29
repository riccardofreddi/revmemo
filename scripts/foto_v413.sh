#!/bin/bash
# Foto del pannello RevMemo v4.13: chiave di sessione (le API senza chiave
# rispondono 403), password mai inviata al browser (campo vuoto con nota
# "gia' salvata"), home con clienti demo e registro.
set -e
PORTE=8768
FOTO=/home/z/my-project/download/foto_pannello
DEMO=/tmp/demo_pannello_v413

mkdir -p "$DEMO" "$FOTO"
cp /home/z/my-project/download/RevMemo/pannello.py \
   /home/z/my-project/download/RevMemo/pannello.html \
   /home/z/my-project/download/RevMemo/revmemo.py \
   /home/z/my-project/download/RevMemo/controlla_pc.py \
   /home/z/my-project/download/RevMemo/template_email.html \
   /home/z/my-project/download/RevMemo/logo.png \
   /home/z/my-project/download/RevMemo/clienti_revisioni.xlsx \
   /home/z/my-project/download/RevMemo/config.ini "$DEMO/"
rm -f "$DEMO/reminders_log.txt" "$DEMO/revmemo.lock" "$DEMO/stop_gestiti.txt"
# password FINTA di prova: cosi' la foto delle Impostazioni mostra il
# campo vuoto con "gia' salvata" (la password non viaggia verso il browser)
sed -i 's/^password_app *= *$/password_app = PROVAFINTA123456/' "$DEMO/config.ini"

pkill -f "pannello.py" 2>/dev/null || true
sleep 1
cd "$DEMO"
PANNELLO_DEMO=1 PANNELLO_PORTA=$PORTE nohup python3 pannello.py > server.log 2>&1 &
PID=$!
for i in $(seq 1 30); do
  if curl -s -m 2 "http://127.0.0.1:$PORTE/api/ping" 2>/dev/null | grep -q revmemo; then break; fi
  sleep 0.5
done
echo "demo attiva (pid $PID)"

echo "TEST protezione: API senza chiave (deve dire 'errore'):"
curl -s "http://127.0.0.1:$PORTE/api/stato" | head -c 120; echo
TOKEN=$(curl -s "http://127.0.0.1:$PORTE/" | grep -o 'name="revmemo-token" content="[^"]*"' | cut -d'"' -f4)
echo "TEST chiave letta dalla pagina: ${#TOKEN} caratteri"

# 28 clienti demo + 30 righe di registro (come v4.10/v4.11, ora CON chiave)
export DEMO_TOKEN="$TOKEN"
python3 - <<'EOF'
import json, os, urllib.request, datetime as dt
URL = "http://127.0.0.1:8768/api/cliente/nuovo"
intest = {"Content-Type": "application/json",
          "X-RevMemo-Token": os.environ["DEMO_TOKEN"]}
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
    req = urllib.request.Request(URL, data=corpo, headers=intest)
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

# ---- 1) HOME: clienti demo + statistiche
echo "TEST home caricata:"; agent-browser eval "document.getElementById('tab-clienti') ? 'OK' : 'no'"
agent-browser screenshot "$FOTO/v413_1_home.png" 2>&1 | tail -1

# ---- 2) IMPOSTAZIONI: campo password VUOTO con nota "gia' salvata"
agent-browser eval "document.querySelector('nav button[data-vista=impostazioni]').click()" > /dev/null 2>&1 || true
sleep 1.5
echo "TEST campo password vuoto (non contiene la password vera):"
agent-browser eval "document.getElementById('i-password').value === '' ? 'OK: vuoto' : 'ERRORE: ha un valore'"
echo "TEST nota 'gia' salvata':"
agent-browser eval "document.getElementById('i-password').placeholder"
echo "TEST aiuto password:"
agent-browser eval "document.getElementById('aiuto-password').textContent"
agent-browser screenshot "$FOTO/v413_2_impostazioni_password_nascosta.png" 2>&1 | tail -1

# ---- 3) REGISTRO con le righe demo
agent-browser eval "document.querySelector('nav button[data-vista=registro]').click()" > /dev/null 2>&1 || true
sleep 1.5
agent-browser screenshot "$FOTO/v413_3_registro.png" 2>&1 | tail -1

kill $PID 2>/dev/null || true
pkill -f "pannello.py" 2>/dev/null || true
echo "foto v4.13 scattate"
ls "$FOTO" | grep v413 || true
