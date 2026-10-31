#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Aggiorna LEGGIMI.txt per la v4.15 (mantenendo i CRLF):
1) inserisce la sezione NOVITA' v4.15 sopra la NOVITA' v4.14;
2) "ultimi 7 giorni" -> "ultimi 35 giorni" (2 punti della sezione STOP);
3) aggiunge il pulsante CONTROLLA LE RISPOSTE STOP alla sezione 6.
Idempotente: se una modifica c'e' gia', non la rifà."""
import sys

PERCORSO = sys.argv[1] if len(sys.argv) > 1 else \
    "/home/z/my-project/revmemo-github/revmemo/LEGGIMI.txt"

SEZIONE = """NOVITA' v4.15 (ottobre 2026)
----------------------------------------------------------------------
- NESSUNA RICHIESTA STOP SI PERDE PIU': il controllo delle risposte
  guarda ora gli ultimi 35 giorni della casella email (prima 7). Anche
  se il computer resta spento per settimane, quando riparte il programma
  trova comunque le risposte STOP arrivate nel frattempo, PRIMA di
  inviare qualsiasi email.
- NUOVO PULSANTE "CONTROLLA LE RISPOSTE STOP" nel pannello: legge la
  casella email e registra gli STOP nel file clienti SENZA mandare
  niente. Serve raramente (l'attivita' pianificata giornaliera controlla
  gia' ogni mattina), ma dopo un lungo fermo puoi verificare con un
  click, quando vuoi, senza invii.
- La ricerca resta "di passaggio": le email dei clienti non risultano
  lette, e le risposte gia' gestite non vengono registrate due volte
  (riattivare un cliente a mano resta sicuro).

"""

NOTA_PULSANTE = """- In piu', dalla v4.15: nel pannello c'e' il pulsante CONTROLLA LE
  RISPOSTE STOP che fa questo controllo quando vuoi tu, senza inviare
  niente (comodo dopo un lungo fermo del computer)
"""

with open(PERCORSO, "rb") as f:
    grezzo = f.read()

cr_lf = b"\r\n" in grezzo
a_capo = "\r\n" if cr_lf else "\n"
testo = grezzo.decode("utf-8")
fatto = []

if "NOVITA' v4.15" not in testo:
    jalon = "NOVITA' v4.14"
    pos = testo.find(jalon)
    if pos < 0:
        print("ERRORE: non trovo il punto di inserimento (NOVITA' v4.14).")
        sys.exit(1)
    testo = testo[:pos] + SEZIONE.replace("\n", a_capo) + testo[pos:]
    fatto.append("sezione NOVITA' v4.15")

if "gli ultimi 7 giorni arrivate" in testo:
    testo = testo.replace("gli ultimi 7 giorni arrivate",
                          "gli ultimi 35 giorni arrivate")
    fatto.append("7->35 giorni (invio)")

if "la ricerca guarda gli ultimi 7 giorni" in testo:
    testo = testo.replace("la ricerca guarda gli ultimi 7 giorni",
                          "la ricerca guarda gli ultimi 35 giorni")
    fatto.append("7->35 giorni (ricerca)")

jalon_pulsante = "- Da quel momento quel cliente non riceve piu' niente"
if "RISPOSTE STOP che fa questo controllo" not in testo \
        and jalon_pulsante in testo:
    nota = NOTA_PULSANTE.replace("\n", a_capo)
    testo = testo.replace(jalon_pulsante + a_capo,
                          jalon_pulsante + a_capo + nota, 1)
    fatto.append("nota pulsante in sezione 6")

with open(PERCORSO, "wb") as f:
    f.write(testo.encode("utf-8"))
riepilogo = "; ".join(fatto) if fatto else "gia' tutto aggiornato"
print(f"{PERCORSO}: fatte {len(fatto)} modifiche ({riepilogo}) "
      f"[CRLF={'si' if cr_lf else 'no'}]")
