#!/bin/bash
# Foto del pannello RevMemo v4.11: istruzioni Registro con menu righe,
# opzione "Vedi 100 righe", finestra di conferma INTERNA (niente
# "127.0.0.1:8768 dice:" del browser).
set -e
PORTE=8768
FOTO=/home/z/my-project/download/foto_pannello
DEMO=/tmp/demo_pannello_v411

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

# 28 clienti demo + 30 righe di registro (come v4.10)
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

# ---- 1) REGISTRO: nuovo testo + menu con 100
agent-browser eval "document.querySelector('nav button[data-vista=registro]').click()" > /dev/null 2>&1 || true
sleep 1.5
echo "TEST menu 100 esiste:"; agent-browser eval "Array.from(document.getElementById('r-righe').options).map(o=>o.value).join(',')"
echo "TEST altezza registro a 20:"; agent-browser eval "document.getElementById('tabella-registro').style.maxHeight"
agent-browser eval "const s=document.getElementById('r-righe'); s.value='100'; s.dispatchEvent(new Event('change')); 'ok'" > /dev/null 2>&1 || true
sleep 0.5
echo "TEST altezza registro a 100 (deve salire):"; agent-browser eval "document.getElementById('tabella-registro').style.maxHeight"
echo "TEST scelta memorizzata:"; agent-browser eval "localStorage.getItem('revmemo_righe_registro')"
agent-browser screenshot "$FOTO/v411_1_registro_100.png" 2>&1 | tail -1

# ---- 2) CONFERMA INTERNA: INVIA ORA apre la finestra del pannello
agent-browser eval "document.querySelector('nav button[data-vista=home]').click()" > /dev/null 2>&1 || true
sleep 1.5
agent-browser eval "document.getElementById('b-invia').click()" > /dev/null 2>&1 || true
sleep 1
echo "TEST finestra conferma aperta:"; agent-browser eval "document.getElementById('velo-conferma').classList.contains('aperto')"
echo "TEST testo domanda:"; agent-browser eval "document.getElementById('conferma-testo').textContent"
agent-browser screenshot "$FOTO/v411_2_conferma.png" 2>&1 | tail -1
echo "TEST nessun '127.0.0.1' nel modal:"; agent-browser eval "document.getElementById('conferma-testo').textContent.includes('127.0.0.1') ? 'ERRORE: contiene indirizzo' : 'OK: solo la domanda'"

# ---- 3) "No, annulla" chiude e NON avvia niente
agent-browser eval "document.getElementById('conferma-no').click()" > /dev/null 2>&1 || true
sleep 0.5
echo "TEST finestra chiusa:"; agent-browser eval "document.getElementById('velo-conferma').classList.contains('aperto')"
echo "TEST nessun avvio (uscita nascosta):"; agent-browser eval "document.getElementById('uscita').style.display !== 'block' ? 'OK: non parte' : 'ERRORE: e partito!'"

kill $PID 2>/dev/null || true
pkill -f "pannello.py" 2>/dev/null || true
echo "foto v4.11 scattate"
ls "$FOTO" | grep v411 || true
