"""Registro modifiche AFFITTI.

Scarica AFFITTI.xlsx da Dropbox, legge l'area di stampa del foglio AFFITTI,
la confronta con l'ultima versione salvata e, se e' cambiata, salva:

  data/versioni/<id>.json   copia completa della tabella (la "foto")
  data/indice.json          elenco versioni + eventi (aggiunte/rimosse/modificate)
  data/stato.json           ultimo controllo (aggiornato al massimo una volta al giorno
                            se non ci sono modifiche, per non creare commit inutili)

Uso:
  python scripts/snapshot.py                 # scarica da DROPBOX_URL (o link di default)
  python scripts/snapshot.py --file a.xlsx   # usa un file locale (prove)
"""

import argparse
import hashlib
import io
import json
import os
import re
import sys
import time
import urllib.request
from datetime import datetime, timezone

import openpyxl
from openpyxl.utils import column_index_from_string, range_boundaries

DEFAULT_URL = ("https://www.dropbox.com/scl/fi/q3ewoodokhsrd8g11l915/AFFITTI.xlsx"
               "?rlkey=47altne9kdaw9y5h8bm57lun5&st=qjbkqviq&dl=0")
SHEET = "AFFITTI"
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.environ.get("AFFITTI_DATA") or os.path.join(ROOT, "data")
CAMPI = ["pt", "sopp", "cort", "tot", "canone"]


# ---------------------------------------------------------------- download
def direct_link(url):
    u = url.strip().replace("www.dropbox.com", "dl.dropboxusercontent.com")
    if "dl=0" in u:
        u = u.replace("dl=0", "dl=1")
    elif "dl=1" not in u:
        u += ("&" if "?" in u else "?") + "dl=1"
    return u


def download(url, tentativi=3):
    err = None
    for i in range(tentativi):
        try:
            req = urllib.request.Request(direct_link(url), headers={"User-Agent": "affitti-registro"})
            with urllib.request.urlopen(req, timeout=60) as r:
                data = r.read()
            if data[:2] != b"PK":
                raise ValueError("risposta non e' un file xlsx")
            return data
        except Exception as e:  # noqa: BLE001
            err = e
            time.sleep(10 * (i + 1))
    raise SystemExit("Download fallito: %s" % err)


# ---------------------------------------------------------------- lettura
def valore(v):
    if isinstance(v, datetime):
        return {"data": v.strftime("%Y-%m-%dT%H:%M")}
    if isinstance(v, float) and v.is_integer():
        return int(v)
    if isinstance(v, str):
        return v.strip()
    return v


def area_stampa(ws):
    pa = ws.print_area
    if isinstance(pa, (list, tuple)):
        pa = pa[0] if pa else None
    if pa:
        ref = pa.split("!")[-1].replace("$", "")
        c0, r0, c1, r1 = range_boundaries(ref)
        return r0, r1, c0, c1
    # ripiego: colonne A-F fino alla prima riga vuota dopo l'intestazione
    r1 = 2
    while r1 < ws.max_row and any(ws.cell(r1 + 1, c).value is not None for c in range(1, 7)):
        r1 += 1
    return 1, r1, 1, 6


def colonne_nascoste(ws):
    nascoste = set()
    for key, dim in ws.column_dimensions.items():
        if dim.hidden:
            lo = dim.min or column_index_from_string(key)
            hi = dim.max or lo
            nascoste.update(range(lo, hi + 1))
    return nascoste


def leggi_tabella(xlsx_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes), data_only=True)
    if SHEET not in wb.sheetnames:
        raise SystemExit("Foglio %r non trovato" % SHEET)
    ws = wb[SHEET]
    r0, r1, c0, c1 = area_stampa(ws)
    hid_c = colonne_nascoste(ws)
    righe_vis = [r for r in range(r0, r1 + 1) if not ws.row_dimensions[r].hidden]
    col_vis = [c for c in range(c0, c1 + 1) if c not in hid_c]

    griglia = [[valore(ws.cell(r, c).value) for c in col_vis] for r in righe_vis]

    unioni = []
    for m in ws.merged_cells.ranges:
        rr = [i for i, r in enumerate(righe_vis) if m.min_row <= r <= m.max_row]
        cc = [i for i, c in enumerate(col_vis) if m.min_col <= c <= m.max_col]
        if rr and cc:
            unioni.append([rr[0], cc[0], rr[-1], cc[-1]])

    # data/ora di Excel (cella con =NOW() nell'intestazione)
    excel_ora = None
    for v in griglia[0] if griglia else []:
        if isinstance(v, dict) and "data" in v:
            excel_ora = v["data"]

    return {"griglia": griglia, "unioni": unioni}, excel_ora


# ---------------------------------------------------------------- unita
# errori di battitura visti nelle schede: stessa unita'
ALIAS = {"stocaggio": "stoccaggio", "montacarich": "montacarichi", "motacarichi": "montacarichi",
         "essecatura": "essiccatura", "merge": "mergellina", "mergell": "mergellina"}

# nomi con "/" che sono UN solo loft (unita' unite), non un gruppo (chiave senza spazi)
UNITE = {"stampaggio4/5"}
NUM = re.compile(r"^\d+(?:\.\d+)?[a-z]?$", re.I)


def espandi(nome):
    """Un nome con piu' numeri e' un gruppo di unita' uguali (stessi mq e canone):
      "mergellina beach p1 17 21 22" -> ... 17, ... 21, ... 22   (spazio = elenco)
      "palazzina 203/208"            -> 203, 204, ..., 208       (a/b con b > a = intervallo)
      "super condor 2 / 5 / 6"       -> 2, 5, 6                  (3+ numeri o b < a = elenco)
      "shed p1 16/18c"               -> 16c, 17c, 18c            (lettera finale vale per tutti)
    Il trattino ("stampaggio 4 - 5") e i nomi in UNITE restano una sola unita'.
    Restituisce la lista dei nomi delle singole unita'."""
    pulito = re.sub(r"\s+", " ", re.sub(r"[^\w\s/.\-]", " ", nome)).strip()
    if chiave(nome) in UNITE:
        return [pulito]
    tok = re.sub(r"\s*/\s*", " / ", pulito).split()
    i = next((j for j, t in enumerate(tok) if NUM.match(t)), None)
    if i is None or i == 0:
        return [pulito]
    base, resto = " ".join(tok[:i]), tok[i:]
    if any(t != "/" and not NUM.match(t) for t in resto):
        return [pulito]                       # es. trattino: unita' unite
    catene, prima = [], True                  # numeri legati da "/" formano una catena
    for t in resto:
        if t == "/":
            prima = False
            continue
        if prima or not catene:
            catene.append([t])
        else:
            catene[-1].append(t)
        prima = True
    numeri = [n for c in catene for n in c]
    if len(numeri) < 2:
        return [pulito]
    suff = re.sub(r"^[\d.]+", "", numeri[-1])
    if suff and not any(re.sub(r"^[\d.]+", "", n) for n in numeri[:-1]):
        numeri = [n + suff for n in numeri[:-1]] + [numeri[-1]]
        catene = [[n + suff if not re.sub(r"^[\d.]+", "", n) else n for n in c] for c in catene]
    out = []
    for c in catene:
        if len(c) == 2 and c[0].isdigit() == c[1].isdigit():
            a, b = (int(re.match(r"\d+", n).group()) for n in c)
            sa = re.sub(r"^\d+", "", c[0])
            if 0 < b - a <= 20 and "." not in c[0] + c[1]:
                out += ["%d%s" % (n, sa) for n in range(a, b + 1)]
                continue
        out += c
    return ["%s %s" % (base, n) for n in dict.fromkeys(out)]


def chiave(nome):
    """Chiave che riconosce la stessa unita': minuscolo, senza emoji/simboli,
    senza spazi (super condor 9 = supercondor 9), refusi noti corretti."""
    k = nome.lower()
    k = re.sub(r"[^\w\s/.\-]", " ", k)          # via emoji e simboli
    k = re.sub(r"(\d)\s*-\s*(\d)", r"\1/\2", k)  # "4 - 5" = "4/5" (unita' unite)
    k = " ".join(ALIAS.get(p, p) for p in k.split())
    return re.sub(r"\s+", "", k)


def aggiungi_unita(unita, viste, nome, sezione, dati, riga):
    """Aggiunge la riga della scheda come una o piu' unita' (gruppi scomposti)."""
    nomi = espandi(nome)
    for n in nomi:
        k = chiave(n)
        viste[k] = viste.get(k, 0) + 1
        if viste[k] > 1:
            k = "%s #%d" % (k, viste[k])
        u = {"key": k, "nome": n if len(nomi) > 1 else re.sub(r"\s+", " ", nome).strip(),
             "sezione": sezione, **dati, "riga": riga}
        if len(nomi) > 1:
            u["gruppo"] = re.sub(r"\s+", " ", nome).strip()
        unita.append(u)


def canon_sezione(testo):
    """Nome uniforme della via/sezione (senza numeri di telefono)."""
    t = str(testo or "").split("\n")[0].strip()
    u = t.upper()
    if "TUCIDIDE" in u:
        return "VIA TUCIDIDE"
    if "MALAGA" in u:
        return "VIA MALAGA"
    if "CASCATA" in u or "ROZZANO" in u:
        return "CASCATA (Rozzano)"
    return re.sub(r"\s*[\d/ ]{6,}$", "", t).strip() or "?"


def num(v):
    return v if isinstance(v, (int, float)) and not isinstance(v, bool) else None


def estrai_unita(griglia):
    """Riga 1 = intestazione, riga 2 = etichette (indirizzo nella prima cella).
    Poi: riga senza numeri con la prima parola MAIUSCOLA = nuova sezione
    (es. VIA MALAGA, CASCATA via ...);
    altrimenti e' un'unita' (anche se senza mq/canone)."""
    if len(griglia) < 2:
        return []
    sezione = canon_sezione(griglia[1][0])
    unita, viste = [], {}
    for i, riga in enumerate(griglia[2:], start=2):
        nome = riga[0]
        if not isinstance(nome, str) or not nome.strip():
            continue
        valori = [num(v) for v in riga[1:]]
        primo = re.sub(r"[^A-Za-zÀ-ÿ]", "", nome.split()[0])
        if all(v is None for v in valori) and len(primo) >= 2 and primo == primo.upper():
            sezione = canon_sezione(nome)
            continue
        dati = dict(zip(CAMPI, (valori + [None] * len(CAMPI))[:len(CAMPI)]))
        if dati["tot"] is None and any(dati[c] is not None for c in ("pt", "sopp", "cort")):
            dati["tot"] = sum(dati[c] or 0 for c in ("pt", "sopp", "cort"))
        aggiungi_unita(unita, viste, nome, sezione, dati, i)
    return unita


def confronta(prima, dopo):
    p = {u["key"]: u for u in prima}
    d = {u["key"]: u for u in dopo}
    eventi = []
    for u in dopo:
        if u["key"] not in p:
            eventi.append({"tipo": "aggiunta", **u})
            continue
        old = p[u["key"]]
        # il nome non si confronta: stessa chiave = stessa unita' anche se scritta diversa
        campi = {c: [old.get(c), u.get(c)] for c in CAMPI + ["sezione"] if old.get(c) != u.get(c)}
        if campi:
            eventi.append({"tipo": "modificata", **u, "campi": campi})
    for u in prima:
        if u["key"] not in d:
            eventi.append({"tipo": "rimossa", **u})
    return eventi


# ---------------------------------------------------------------- salvataggio
def leggi_json(path, default):
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        return default


def scrivi_json(path, obj, compatto=False):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8", newline="\n") as f:
        if compatto:
            json.dump(obj, f, ensure_ascii=False, separators=(",", ":"))
        else:
            json.dump(obj, f, ensure_ascii=False, indent=1)
        f.write("\n")


def impronta(tabella, unita):
    # la cella =NOW() cambia a ogni salvataggio: esclusa dal confronto
    g = [[None if isinstance(v, dict) else v for v in r] for r in tabella["griglia"]]
    s = json.dumps({"g": g, "m": tabella["unioni"], "u": unita}, ensure_ascii=False, sort_keys=True)
    return hashlib.sha256(s.encode("utf-8")).hexdigest()[:16]


def totali(unita):
    return {
        "n_unita": len(unita),
        "mq_tot": sum(u["tot"] or 0 for u in unita),
        "canone_tot": sum(u["canone"] or 0 for u in unita),
    }


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser()
    ap.add_argument("--file")
    ap.add_argument("--ora", help="data/ora UTC simulata, per le prove (YYYY-MM-DDTHH:MM:SS)")
    args = ap.parse_args()

    if args.file:
        with open(args.file, "rb") as f:
            xlsx = f.read()
    else:
        xlsx = download(os.environ.get("DROPBOX_URL") or DEFAULT_URL)

    adesso = (datetime.fromisoformat(args.ora).replace(tzinfo=timezone.utc) if args.ora
              else datetime.now(timezone.utc)).replace(microsecond=0)
    tabella, excel_ora = leggi_tabella(xlsx)
    unita = estrai_unita(tabella["griglia"])
    imp = impronta(tabella, unita)

    indice = leggi_json(os.path.join(DATA, "indice.json"), {"versioni": []})
    stato_path = os.path.join(DATA, "stato.json")
    stato = leggi_json(stato_path, {})
    rilevato = adesso.isoformat().replace("+00:00", "Z")

    ultima = indice["versioni"][-1] if indice["versioni"] else None
    if ultima and ultima["impronta"] == imp:
        # nessuna modifica: aggiorna stato.json solo una volta al giorno
        if (stato.get("ultimo_controllo") or "")[:10] != rilevato[:10]:
            stato["ultimo_controllo"] = rilevato
            scrivi_json(stato_path, stato)
            print("Nessuna modifica (stato giornaliero aggiornato).")
        else:
            print("Nessuna modifica.")
        return

    if ultima:
        prec = leggi_json(os.path.join(DATA, "versioni", ultima["id"] + ".json"), {})
        eventi = confronta(prec.get("unita", []), unita)
    else:
        eventi = [{"tipo": "iniziale", **u} for u in unita]

    vid = adesso.strftime("%Y%m%d-%H%M%S")
    meta = {"id": vid, "rilevato": rilevato, "excel": excel_ora, "impronta": imp,
            **totali(unita), "eventi": eventi}
    scrivi_json(os.path.join(DATA, "versioni", vid + ".json"),
                {"id": vid, "rilevato": rilevato, "excel": excel_ora, **tabella, "unita": unita},
                compatto=True)
    indice["versioni"].append(meta)
    scrivi_json(os.path.join(DATA, "indice.json"), indice)
    stato.update({"ultimo_controllo": rilevato, "ultima_versione": vid})
    scrivi_json(stato_path, stato)
    print("Nuova versione %s: %d eventi" % (vid, len(eventi)))


if __name__ == "__main__":
    main()
