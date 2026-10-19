# AFFITTI – Lofta Milano

Pagina pubblica: https://rgispace-rgi.github.io/AFFITTI/

| File | Cosa fa |
|---|---|
| `index.html` | Tabella attuale: legge `AFFITTI.xlsx` da Dropbox (foglio `AFFITTI`, area di stampa) e si ricarica ogni 5 minuti. |
| `registro.html` | Registro modifiche: cronologia delle versioni, ricerca per unità/via, filtri (inserite, tolte, prezzo, metratura), elenco unità con data di inserimento/uscita, "foto" di ogni versione della tabella. |
| `scripts/snapshot.py` | Scarica l'Excel, estrae unità (nome, via, MQ P.T./sopp./cortile/tot, canone), confronta con l'ultima versione e salva in `data/`. |
| `.github/workflows/registro.yml` | Esegue `snapshot.py` ogni 15 minuti su GitHub Actions e fa commit dei nuovi dati. |

## Storico da foto (16/09/2024 → 05/09/2026)

Le 81 foto `Scheda affitti GG.MM.AA.JPG` di `\\Nas-LoftAMilano\Dati\A\AFFITTI` sono state
trascritte in `storico/trascrizioni/` (una JSON per foto) e controllate contro la vecchia
estrazione `_ESTRAZIONE_loft_master.csv`: 1351/1357 unità identiche, le 6 differenze
verificate a vista sulla foto. Data = nome del file, ora = data di modifica del file.
Le foto (ricompresse) sono in `foto/` e si vedono nel registro con "Foto originale".

Note: nelle schede 03.02.26 e 09.02.26 l'intestazione dice per errore "VIA MALAGA":
la prima sezione è sempre trattata come VIA TUCIDIDE. Alcuni TOT stampati non sono
PT+SOPP+CORT (es. shed pt 6 c: 29) e sono riportati come nella foto.

Ricostruire lo storico (solo prima del primo controllo automatico):

```bash
python scripts/importa_foto.py --nas "//Nas-LoftAMilano/Dati/A/AFFITTI"
```

## Dati (`data/`)

- `indice.json` – elenco versioni con totali ed eventi (`iniziale`, `aggiunta`, `rimossa`, `modificata` con valori prima/dopo).
- `versioni/<AAAAMMGG-HHMMSS>.json` – copia completa della tabella di quella versione (UTC).
- `stato.json` – ultimo controllo (aggiornato almeno una volta al giorno: tiene attivo il job pianificato).

Una nuova versione nasce solo se cambia il contenuto: la cella `=ADESSO()` dell'intestazione è ignorata.

## Regole di lettura del foglio

- Riga 1 = intestazione, riga 2 = etichette; la prima riga della cella A2 è la via iniziale.
- Riga senza numeri con la prima parola MAIUSCOLA (es. `VIA MALAGA`) = nuova via.
- Ogni altra riga con un nome = unità (anche senza mq/canone). L'unità è riconosciuta dal nome
  (minuscolo, senza emoji/spazi, refusi noti in `ALIAS`): se si rinomina, risulta tolta + inserita.
- Una riga con più numeri è un **gruppo di unità uguali** (stessi mq e canone), scomposto in
  unità singole: se un numero sparisce, quell'unità è stata tolta (affittata).
  - spazio = elenco: `mergellina beach p1 17 21 22` → 17, 21, 22
  - `a/b` con b > a = intervallo: `palazzina 203/208` → 203…208
  - 3+ numeri o b < a = elenco: `super condor 2 / 5 / 6` → 2, 5, 6
  - lettera finale vale per tutti: `shed p1 16/18c` → 16c, 17c, 18c
  - trattino e nomi in `UNITE` = un solo loft: `stampaggio 4 - 5` = `stampaggio 4/5`
- Conta solo l'**area di stampa** del foglio: una riga aggiunta fuori area non viene vista
  (né dalla pagina né dal registro).

## Comandi

```bash
python scripts/snapshot.py                 # controllo manuale (scarica da Dropbox)
python scripts/snapshot.py --file a.xlsx   # prova con un file locale
```

Avvio manuale su GitHub: scheda **Actions → Registro modifiche AFFITTI → Run workflow**.
