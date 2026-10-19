# AFFITTI – Lofta Milano

Pagina pubblica: https://rgispace-rgi.github.io/AFFITTI/

| File | Cosa fa |
|---|---|
| `index.html` | Tabella attuale: legge `AFFITTI.xlsx` da Dropbox (foglio `AFFITTI`, area di stampa) e si ricarica ogni 5 minuti. |
| `registro.html` | Registro modifiche: cronologia delle versioni, ricerca per unità/via, filtri (inserite, tolte, prezzo, metratura), elenco unità con data di inserimento/uscita, "foto" di ogni versione della tabella. |
| `scripts/snapshot.py` | Scarica l'Excel, estrae unità (nome, via, MQ P.T./sopp./cortile/tot, canone), confronta con l'ultima versione e salva in `data/`. |
| `.github/workflows/registro.yml` | Esegue `snapshot.py` ogni 15 minuti su GitHub Actions e fa commit dei nuovi dati. |

## Dati (`data/`)

- `indice.json` – elenco versioni con totali ed eventi (`iniziale`, `aggiunta`, `rimossa`, `modificata` con valori prima/dopo).
- `versioni/<AAAAMMGG-HHMMSS>.json` – copia completa della tabella di quella versione (UTC).
- `stato.json` – ultimo controllo (aggiornato almeno una volta al giorno: tiene attivo il job pianificato).

Una nuova versione nasce solo se cambia il contenuto: la cella `=ADESSO()` dell'intestazione è ignorata.

## Regole di lettura del foglio

- Riga 1 = intestazione, riga 2 = etichette; la prima riga della cella A2 è la via iniziale.
- Riga senza numeri con la prima parola MAIUSCOLA (es. `VIA MALAGA`) = nuova via.
- Ogni altra riga con un nome = unità (anche senza mq/canone). L'unità è riconosciuta dal nome
  (minuscolo, senza emoji/spazi doppi): se si rinomina, risulta tolta + inserita.
- Conta solo l'**area di stampa** del foglio: una riga aggiunta fuori area non viene vista
  (né dalla pagina né dal registro).

## Comandi

```bash
python scripts/snapshot.py                 # controllo manuale (scarica da Dropbox)
python scripts/snapshot.py --file a.xlsx   # prova con un file locale
```

Avvio manuale su GitHub: scheda **Actions → Registro modifiche AFFITTI → Run workflow**.
