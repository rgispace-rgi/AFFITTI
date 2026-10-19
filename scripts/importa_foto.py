"""Importa nello storico le foto "Scheda affitti GG.MM.AA.JPG" (prima del registro automatico).

Legge le trascrizioni in storico/trascrizioni/*.json (una per foto, fatte a mano/AI e
verificate), le ordina per data e ricostruisce data/ da zero con una versione per ogni
foto che cambia qualcosa. Copia anche le foto in foto/ per mostrarle nel registro.

Uso:
  python scripts/importa_foto.py --nas "\\\\Nas-LoftAMilano\\Dati\\A\\AFFITTI"

Va eseguito PRIMA del primo controllo automatico (rifiuta se in data/ ci sono gia'
versioni prese dall'Excel, salvo --forza che le mantiene in coda).
"""

import argparse
import glob
import json
import os
import re
import sys
from datetime import datetime, timezone
from zoneinfo import ZoneInfo

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import snapshot as S  # noqa: E402

ROMA = ZoneInfo("Europe/Rome")
TRASCR = os.path.join(S.ROOT, "storico", "trascrizioni")
FOTO = os.path.join(S.ROOT, "foto")


def campo_colonna(label):
    u = re.sub(r"\s+", "", label.upper())
    if "IVA" in u:
        return "canone_iva"
    if "ANNO" in u:
        return "dal2anno"
    if "CANONE" in u:
        return "canone"
    for k, pat in (("pt", "PT"), ("sopp", "SOPP"), ("cort", "CORT"), ("tot", "TOT")):
        if pat in u:
            return k
    return None


def numero(cella):
    t = str(cella or "").strip()
    if not t or "?" in t:
        return None
    t = t.replace("€", "").replace(".", "").replace(" ", "").replace(",", ".")
    try:
        f = float(t)
        return int(f) if f.is_integer() else f
    except ValueError:
        return None


def quando(tr, nas):
    """Data dal nome file; ora dalla data di modifica del file (o dalla cella in alto a destra)."""
    m = re.search(r"(\d{2})\.(\d{2})\.(\d{2})", tr["file"])
    g, me, a = int(m.group(1)), int(m.group(2)), 2000 + int(m.group(3))
    ora = None
    path = os.path.join(nas, tr["file"]) if nas else None
    if path and os.path.exists(path):
        mt = datetime.fromtimestamp(os.path.getmtime(path))
        if (mt.year, mt.month, mt.day) == (a, me, g):
            ora = (mt.hour, mt.minute)
    if ora is None:
        h = re.search(r"(\d{1,2})[.:](\d{2})\s*$", tr.get("header_right") or "")
        ora = (int(h.group(1)), int(h.group(2))) if h else (12, 0)
    return datetime(a, me, g, ora[0], ora[1], tzinfo=ROMA)


def converti(tr):
    cols = tr["columns"]
    campi = [campo_colonna(c) for c in cols]
    righe = tr["rows"]
    sezioni = [S.canon_sezione(r["text"]) for r in righe if r["type"] == "section"]
    prima = sezioni[0] if sezioni else "VIA TUCIDIDE"
    # intestazione sbagliata (es. "VIA MALAGA" in alto ma poi c'e' anche la sezione VIA MALAGA)
    if prima in sezioni[1:] or prima == "?":
        prima = "VIA TUCIDIDE"

    griglia = [
        ["VAI SU www.loftamilano.it"] + ["MQ" if c in ("pt", "sopp", "cort", "tot") else "" for c in campi],
        [righe[0]["text"] if righe and righe[0]["type"] == "section" else prima]
        + [re.sub(r"^MQ\s*", "", c) for c in cols],
    ]
    if tr.get("header_right"):
        griglia[0][-1] = tr["header_right"]
    unioni, unita, viste = [], [], {}
    sezione = prima
    for r in righe[1:] if righe and righe[0]["type"] == "section" else righe:
        i = len(griglia)
        if r["type"] == "unit":
            # descrizioni di icone messe dal trascrittore, es. "[flame icon]"
            r = {**r, "name": re.sub(r"\s*\[[^\]]*\]", "", r["name"]).strip()}
            if not r["name"]:
                continue
        if r["type"] == "section":
            sezione = S.canon_sezione(r["text"])
            griglia.append([r["text"]] + [None] * len(cols))
            unioni.append([i, 0, i, len(cols)])
            continue
        valori = [numero(c) for c in r["cells"]]
        griglia.append([r["name"].strip()] + valori)
        dati = {c: None for c in S.CAMPI}
        for c, v in zip(campi, valori):
            if c in dati:
                dati[c] = v
        if dati["tot"] is None and any(dati[c] is not None for c in ("pt", "sopp", "cort")):
            dati["tot"] = sum(dati[c] or 0 for c in ("pt", "sopp", "cort"))
        k = S.chiave(r["name"])
        viste[k] = viste.get(k, 0) + 1
        if viste[k] > 1:
            k = "%s #%d" % (k, viste[k])
        unita.append({"key": k, "nome": re.sub(r"\s+", " ", r["name"]).strip(), "sezione": sezione,
                      **dati, "riga": i})
    return {"griglia": griglia, "unioni": unioni}, unita


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser()
    ap.add_argument("--nas", help="cartella delle foto originali (per ora file e copia foto)")
    ap.add_argument("--forza", action="store_true")
    args = ap.parse_args()

    indice_path = os.path.join(S.DATA, "indice.json")
    vecchio = S.leggi_json(indice_path, {"versioni": []})
    live = [v for v in vecchio["versioni"] if v.get("fonte") != "foto"]
    if live and not args.forza:
        raise SystemExit("In data/ ci sono gia' %d versioni dall'Excel: usa --forza" % len(live))

    trs = []
    for p in glob.glob(os.path.join(TRASCR, "*.json")):
        with open(p, encoding="utf-8") as f:
            tr = json.load(f)
        trs.append((quando(tr, args.nas), tr))
    trs.sort(key=lambda x: x[0])

    # ricostruisce da zero le versioni da foto
    for v in vecchio["versioni"]:
        if v.get("fonte") == "foto":
            try:
                os.remove(os.path.join(S.DATA, "versioni", v["id"] + ".json"))
            except FileNotFoundError:
                pass
    os.makedirs(FOTO, exist_ok=True)

    versioni, prec, uguali = [], None, []
    for dt, tr in trs:
        tabella, unita = converti(tr)
        imp = S.impronta({"griglia": [[None]], "unioni": []}, [{k: v for k, v in u.items() if k != "riga"} for u in unita])
        if prec and prec["imp"] == imp:
            uguali.append(tr["file"])
            continue
        eventi = S.confronta(prec["unita"], unita) if prec else [{"tipo": "iniziale", **u} for u in unita]
        utc = dt.astimezone(timezone.utc)
        vid = utc.strftime("%Y%m%d-%H%M%S")
        foto_rel = "foto/%s.jpg" % vid
        if args.nas and os.path.exists(os.path.join(args.nas, tr["file"])):
            from PIL import Image
            with Image.open(os.path.join(args.nas, tr["file"])) as im:
                im.convert("RGB").save(os.path.join(S.ROOT, foto_rel), "JPEG", quality=85, optimize=True)
        rilevato = utc.isoformat().replace("+00:00", "Z")
        excel = dt.strftime("%Y-%m-%dT%H:%M")
        S.scrivi_json(os.path.join(S.DATA, "versioni", vid + ".json"),
                      {"id": vid, "rilevato": rilevato, "excel": excel, **tabella, "unita": unita,
                       "fonte": "foto", "file": tr["file"], "foto": foto_rel}, compatto=True)
        versioni.append({"id": vid, "rilevato": rilevato, "excel": excel, "impronta": imp,
                         "fonte": "foto", "file": tr["file"], "foto": foto_rel,
                         **S.totali(unita), "eventi": eventi})
        prec = {"imp": imp, "unita": unita}

    indice = {"versioni": versioni + live}
    S.scrivi_json(indice_path, indice)
    stato = S.leggi_json(os.path.join(S.DATA, "stato.json"), {})
    stato.setdefault("ultima_versione", versioni[-1]["id"] if versioni else None)
    S.scrivi_json(os.path.join(S.DATA, "stato.json"), stato)
    print("Foto lette: %d  versioni create: %d  identiche alla precedente: %d"
          % (len(trs), len(versioni), len(uguali)))


if __name__ == "__main__":
    main()
