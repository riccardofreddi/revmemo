# RevMemo — archivio completo del progetto

Copia di sicurezza di **RevMemo**: il programma per la gestione dei
promemoria di revisione per i centri revisioni. Pannello locale nel
browser, motore di invio email, file Excel dei clienti, tutto offline
sul computer del centro.

## Che cosa c'è dentro

| Cartella | Contenuto |
|---|---|
| `revmemo/` | Il pacchetto completo pronto da installare: pannello (`pannello.py` + `pannello.html`), motore (`revmemo.py`), runtime Windows Python 3.12.8 incorporato con openpyxl 3.1.5, guide (`LEGGIMI.txt`, `ISTRUZIONI_CENTRO.md`, `GUIDA_INSTALLATORE.md`) |
| `dist/` | I file ZIP consegnati: `RevMemo_v4.zip` = versione attuale (v4.12); `archivio/` = versioni precedenti |
| `scripts/` | Strumenti di sviluppo: suite di test `test_pannello.py` (151 controlli), script per ricostruire lo ZIP (`ricrea_zip_revmemo_v45.py`), verifiche HTML, harness fotografici |
| `scripts/revmemo_build/` | Asset per ricostruire il runtime incorporato (python_embed.zip + wheel openpyxl) |
| `foto/` | Schermate di verifica del pannello, versione per versione |
| `archivio/` | Progetti precedenti (`promemoria-revisioni`, la prima versione senza pannello) |
| `worklog.md` | Diario di sviluppo: ogni versione con richieste, scelte, test e verifiche |

## Versione attuale: v4.12 (2026-10-29)

- Salvando una **scadenza nuova**, la voce "Revisione fatta" torna **NO
  da sola** (avviso nel pannello) — niente più promemoria bloccati per errore
- Pulsante **🗑 Elimina** cliente accanto a Modifica: conferma interna
  nel pannello (nessun browser), backup automatico prima di cancellare,
  le righe sotto si riaggregano da sole
- Aiuto password in Impostazioni: **Libero/Virgilio = password normale
  della casella**, Gmail = Password per app di 16 caratteri
- STOP: nessuna modifica necessaria (verificato: ferma già tutto)
- Test: **151/151 verdi**

Storico completo delle versioni in `worklog.md`.

## Come si ricostruisce il pacchetto ZIP

```
python3 scripts/ricrea_zip_revmemo_v45.py
```

Lo script produce lo ZIP con 258 file (runtime incluso, dati clienti
esclusi). Prima di ricompilare, verificare che non ci siano cartelle
`__pycache__` nel pacchetto.

## Dati dei clienti

I file Excel con i clienti VERI **non sono in questo archivio**: stanno
solo sul computer del centro, con i backup automatici nella cartella
`backup`. Qui ci sono solo i dati d'esempio (Mario Rossi, Luigi Verdi e
compagnia). Anche la configurazione (`config.ini`) è il modello vuoto,
senza password.
