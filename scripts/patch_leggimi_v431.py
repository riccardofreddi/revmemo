# -*- coding: utf-8 -*-
"""v4.3.1 — corregge il paragrafo colori in LEGGIMI.txt (CRLF):
solo CELLE colorate, STOP mai nel menu a tendina."""
import io
import os

LEGGIMI = "/home/z/my-project/download/RevMemo/LEGGIMI.txt"

with io.open(LEGGIMI, "r", encoding="utf-8", newline="") as f:
    testo = f.read()


def sostituisci(testo, vecchio, nuovo, etichetta):
    quante = testo.count(vecchio)
    if quante != 1:
        raise SystemExit(f"ERRORE ({etichetta}): trovate {quante} occorrenze, "
                         f"attesa 1. Fragmento: {vecchio[:70]!r}")
    print(f"OK patch: {etichetta}")
    return testo.replace(vecchio, nuovo)


testo = sostituisci(
    testo,
    "Colori automatici: riga VERDE = cliente attivo, riga ROSSA = NO oppure\r\n"
    "STOP (la scritta STOP e' rosso scuro). I colori cambiano da soli quando\r\n"
    "scrivi SÌ o NO, senza rilanciare il programma.\r\n",
    "Colori automatici nelle CELLE di Revisione Fatta e Attivo:\r\n"
    "SÌ = verde, NO = rosso, STOP = rosso scuro. I colori cambiano da soli\r\n"
    "quando cambi SÌ o NO, senza rilanciare il programma. STOP non e' nel\r\n"
    "menu a tendina: lo scrive solo il programma quando un cliente lo\r\n"
    "chiede (cosi' non c'e' confusione).\r\n",
    "sezione 4: paragrafo colori corretto")

with io.open(LEGGIMI, "w", encoding="utf-8", newline="") as f:
    f.write(testo)

assert testo.count("\n") == testo.count("\r\n"), "CRLF rotto"
assert "riga VERDE" not in testo and "riga ROSSA" not in testo
print(f"Aggiornato: {LEGGIMI} (CRLF coerente, niente piu' 'riga VERDE/ROSSA')")
