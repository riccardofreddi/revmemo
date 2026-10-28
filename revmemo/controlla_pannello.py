#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
=====================================================================
 RevMemo - CONTROLLO DEL PANNELLO
=====================================================================
Quando il pannello non si apre: doppio click su controlla_pannello.bat
Controlla i punti tipici (cartella, python incluso, moduli, config,
file clienti, porte, blocco Windows, browser) e salva tutto nel file
assistenza_pannello.txt: basta mandare quel file, o una foto della
finestra, a chi ti assiste.
Non tocca niente: solo letture e una prova di scrittura innocua.
=====================================================================
"""
import configparser
import os
import socket
import sys

BASE = os.path.dirname(os.path.abspath(__file__))
# Python incluso = modalita' isolata (python312._pth): la cartella dello
# script non e' tra i percorsi di ricerca, quindi "import revmemo" non
# troverebbe il motore. La aggiungo io qui.
if BASE not in sys.path:
    sys.path.insert(0, BASE)
RISULTATO = os.path.join(BASE, "assistenza_pannello.txt")
PORTE = (8768, 8769)

righe = []
MOTORE_OK = None          # modulo revmemo, se importabile


def dico(testo=""):
    print(testo)
    righe.append(testo)


def esito(nome, ok, dettaglio=""):
    segno = "OK      " if ok else "PROBLEMA"
    riga = f"[{segno}] {nome}"
    if dettaglio:
        riga += f" -> {dettaglio}"
    dico(riga)
    return ok


def zona_bloccata(percorso):
    """True se il file ha il timbro 'arrivato da internet'
    (solo Windows: e' lo Zone.Identifier che fa comparire l'avviso)"""
    if os.name != "nt":
        return False
    try:
        with open(percorso + ":Zone.Identifier", "r", errors="ignore") as f:
            return bool(f.read().strip())
    except OSError:
        return False


def stato_porta(porta):
    """'libera' / 'pannello' (gia' aperto) / 'altro' (occupata)"""
    s = socket.socket()
    s.settimeout(1.5)
    try:
        s.connect(("127.0.0.1", porta))
    except OSError:
        return "libera"
    finally:
        s.close()
    try:
        import json
        import urllib.request
        with urllib.request.urlopen(
                f"http://127.0.0.1:{porta}/api/stato", timeout=3) as r:
            dati = json.loads(r.read().decode("utf-8"))
        if isinstance(dati, dict) and "nome_centro" in dati:
            return "pannello"
        return "altro"
    except Exception:
        return "altro"


def main():
    global MOTORE_OK
    tutto_ok = True

    dico("=" * 62)
    dico(" RevMemo - CONTROLLO DEL PANNELLO")
    dico("=" * 62)
    dico(f"Cartella: {BASE}")
    dico("")

    # 1. cartella scrivibile
    dico("1) La cartella del programma e' scrivibile")
    try:
        prova = os.path.join(BASE, "_prova_scrittura.tmp")
        with open(prova, "w") as f:
            f.write("ok")
        os.remove(prova)
        tutto_ok &= esito("posso scrivere nella cartella", True)
    except Exception as e:
        tutto_ok &= esito("posso scrivere nella cartella", False, str(e))

    # 2. python incluso
    dico("2) Python incluso nel pacchetto")
    presente = os.path.exists(os.path.join(BASE, "runtime", "python.exe"))
    tutto_ok &= esito("runtime\\python.exe presente", presente,
                      "" if presente else "manca: estrai TUTTO lo zip e riprova")
    dico(f"        versione: {sys.version.split()[0]} ({sys.platform})")

    # 3. openpyxl
    dico("3) Modulo per leggere l'Excel (openpyxl)")
    try:
        import openpyxl
        tutto_ok &= esito("openpyxl presente", True, openpyxl.__version__)
    except Exception as e:
        tutto_ok &= esito("openpyxl presente", False, str(e))

    # 4. motore
    dico("4) Motore revmemo.py")
    try:
        import revmemo
        MOTORE_OK = revmemo
        tutto_ok &= esito("revmemo importato", True,
                          getattr(revmemo, "VERSIONE", ""))
    except BaseException as e:
        tutto_ok &= esito("revmemo importato", False, str(e))

    # 5. config.ini
    dico("5) config.ini (impostazioni)")
    percorso_config = os.path.join(BASE, "config.ini")
    if os.path.exists(percorso_config):
        try:
            cfg = configparser.ConfigParser(interpolation=None)
            cfg.read(percorso_config, encoding="utf-8")
            tutto_ok &= esito("config.ini leggibile", True,
                              ", ".join(cfg.sections()))
        except Exception as e:
            tutto_ok &= esito("config.ini leggibile", False, str(e))
    else:
        esito("config.ini leggibile", True,
              "manca: si crea da solo al primo avvio")

    # 6. file clienti
    dico("6) File clienti (clienti_revisioni.xlsx)")
    percorso_xlsx = os.path.join(BASE, "clienti_revisioni.xlsx")
    if os.path.exists(percorso_xlsx):
        if MOTORE_OK is not None:
            try:
                clienti = MOTORE_OK.leggi_clienti(percorso_xlsx)
                tutto_ok &= esito("file clienti leggibile", True,
                                  f"{len(clienti)} clienti")
            except BaseException as e:
                tutto_ok &= esito("file clienti leggibile", False, str(e))
        else:
            tutto_ok &= esito("file clienti leggibile", False,
                              "saltato: il motore non si importa")
    else:
        esito("file clienti leggibile", True, "manca: si crea da solo")

    # 7. porte
    dico("7) Porte del pannello " + ", ".join(str(p) for p in PORTE))
    for porta in PORTE:
        stato = stato_porta(porta)
        if stato == "libera":
            esito(f"porta {porta}", True, "libera")
        elif stato == "pannello":
            esito(f"porta {porta}", True,
                  "qui c'e' GIA' il pannello: apri http://127.0.0.1:"
                  f"{porta}/ nel browser")
        else:
            tutto_ok &= esito(f"porta {porta}", False,
                              "occupata da un altro programma")

    # 8. blocco Windows
    dico("8) Blocco 'file arrivati da internet' di Windows")
    bloccati = []
    for nome in ("runtime\\python.exe", "pannello.py", "apri_pannello.bat",
                 "controlla_pannello.py"):
        percorso = os.path.join(BASE, nome)
        if os.path.exists(percorso) and zona_bloccata(percorso):
            bloccati.append(nome)
    if bloccati:
        esito("file bloccati da Windows", True,
              "trovati: " + ", ".join(bloccati)
              + " - apri_pannello.bat li sblocca da solo alla prima "
                "apertura (poi la finestra mostra il risultato)")
    else:
        esito("nessun file bloccato", True)

    # 9. browser
    dico("9) Browser predefinito")
    try:
        import webbrowser
        webbrowser.get()
        esito("browser trovato", True)
    except Exception:
        tutto_ok &= esito("browser trovato", False,
                          "nessun browser predefinito: apri tu "
                          "http://127.0.0.1:8768/ in Chrome o Edge")

    # 10. registro del pannello
    dico("10) Ultime righe del registro del pannello (pannello_log.txt)")
    percorso_log = os.path.join(BASE, "pannello_log.txt")
    if os.path.exists(percorso_log):
        try:
            with open(percorso_log, "r", encoding="utf-8",
                      errors="replace") as f:
                ultime = f.read().splitlines()[-8:]
            for u in ultime:
                dico("        " + u)
        except Exception as e:
            dico(f"        non leggibile: {e}")
    else:
        dico("        (il pannello non ha mai scritto il registro: "
             "non e' mai partito davvero)")

    dico("")
    dico("=" * 62)
    if tutto_ok:
        dico("RISULTATO: TUTTO OK")
        dico("Se il pannello ancora non si apre: chiudi ogni finestra")
        dico("nera e fai di nuovo doppio click su apri_pannello.bat,")
        dico("poi guarda quello che scrive la finestra.")
    else:
        dico("RISULTATO: ci sono PROBLEMI (guarda i [PROBLEMA] qui sopra)")
    dico("Questo controllo e' salvato nel file: assistenza_pannello.txt")
    dico("Manda quel file, o una foto di questa finestra, a chi ti assiste.")
    dico("=" * 62)

    try:
        with open(RISULTATO, "w", encoding="utf-8") as f:
            f.write("\n".join(righe) + "\n")
    except Exception as e:
        print(f"(non riesco a salvare {RISULTATO}: {e})")

    try:
        input("\nPremi INVIO per chiudere...")
    except Exception:
        pass


if __name__ == "__main__":
    try:
        main()
    except BaseException:
        import traceback
        print()
        print(traceback.format_exc())
        print("Fai una foto di questa finestra e inviala a chi ti assiste.")
        try:
            input("Premi INVIO per chiudere...")
        except Exception:
            pass
