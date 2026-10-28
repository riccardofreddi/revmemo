# -*- coding: utf-8 -*-
"""v4.12: LEGGIMI.txt — ELIMINA cliente + scadenza nuova che riporta
'Revisione Fatta' a NO da sola (sezione CLIENTI)."""

PERCORSO = "/home/z/my-project/download/RevMemo/LEGGIMI.txt"

VECCHIO = (
    "                colonna), RIATTIVA fa ripartire chi aveva risposto STOP,\r\n"
    "                MODIFICA corregge i dati del cliente.\r\n"
)
NUOVO = (
    "                colonna), RIATTIVA fa ripartire chi aveva risposto STOP,\r\n"
    "                MODIFICA corregge i dati del cliente, ELIMINA lo toglie\r\n"
    "                dal file (conferma prima e copia di sicurezza\r\n"
    "                automatica: le righe sotto salgono da sole). Se SALVI\r\n"
    "                una scadenza NUOVA, Revisione Fatta torna NO da sola e\r\n"
    "                i promemoria ripartono per la data nuova.\r\n"
)

with open(PERCORSO, "rb") as f:
    testo = f.read().decode("utf-8")

fatti = 0
if VECCHIO in testo:
    testo = testo.replace(VECCHIO, NUOVO)
    fatti += 1

with open(PERCORSO, "wb") as f:
    f.write(testo.encode("utf-8"))

print("sostituzioni fatte:", fatti)
print("ELIMINA nel LEGGIMI:", "ELIMINA lo toglie" in testo)
