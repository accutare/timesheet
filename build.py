#!/usr/bin/env python3
"""Junta os CSVs normalizados dos timesheets e gera o painel com os dados embutidos.

Uso:
    python3 build.py <pasta_com_csvs> <saida.html> [aliases.json] [notes.json]

Cada CSV tem o cabeçalho:
    staff,date,client,service,description,hours,kind,source
kind: client | internal | break | holiday  (break e holiday não entram nas horas trabalhadas)

aliases.json (opcional): {"nome como aparece": "nome padronizado"} para juntar clientes escritos de formas diferentes.
notes.json (opcional): lista de {"level": "warn|bad|muted", "title": "...", "detail": "...", "month": "YYYY-MM"}
"""
import csv
import json
import re
import sys
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

HERE = Path(__file__).resolve().parent
SUFFIXES = re.compile(r"\b(ltd|limited|llp|plc|uk|t/a|the)\b\.?", re.I)


def client_key(name: str) -> str:
    k = name.lower().replace("&", " and ")
    k = SUFFIXES.sub(" ", k)
    k = re.sub(r"[^a-z0-9à-ú]+", " ", k)
    return " ".join(k.split())


def main(src: str, out: str, aliases_path: str | None = None, notes_path: str | None = None) -> None:
    aliases = json.loads(Path(aliases_path).read_text()) if aliases_path else {}
    alias_keys = {client_key(k): v for k, v in aliases.items()}
    notes = json.loads(Path(notes_path).read_text()) if notes_path else []

    rows, skipped = [], Counter()
    for f in sorted(Path(src).glob("*.csv")):
        with f.open(newline="", encoding="utf-8") as fh:
            for r in csv.DictReader(fh):
                kind = (r.get("kind") or "").strip().lower()
                if kind in ("break", "holiday"):
                    skipped[kind] += 1
                    continue
                try:
                    hours = round(float(r["hours"]), 2)
                    date = datetime.strptime(r["date"].strip(), "%Y-%m-%d").date()
                except (ValueError, KeyError):
                    skipped["linha inválida"] += 1
                    continue
                if hours <= 0 or date.year != 2026:
                    skipped["fora de 2026 ou sem horas"] += 1
                    continue
                client = " ".join((r.get("client") or "").split())
                is_client = kind == "client" and client and client.lower() not in ("internal", "interno", "-", "—")
                rows.append({
                    "staff": r["staff"].strip(), "date": date.isoformat(),
                    "client": client if is_client else "Internal",
                    "service": (r.get("service") or "Other Client Work").strip() or "Other Client Work",
                    "hours": hours, "kind": "c" if is_client else "i",
                    "desc": " ".join((r.get("description") or "").split())[:120],
                })

    # Padroniza nomes de clientes: mesma chave -> grafia mais usada; aliases têm prioridade.
    by_key = defaultdict(Counter)
    for r in rows:
        if r["kind"] == "c":
            by_key[client_key(r["client"])][r["client"]] += r["hours"]
    canon = {}
    for k, names in by_key.items():
        canon[k] = alias_keys.get(k) or names.most_common(1)[0][0]
    for r in rows:
        if r["kind"] == "c":
            r["client"] = canon[client_key(r["client"])]

    staff = sorted({r["staff"] for r in rows})
    clients = sorted({r["client"] for r in rows}, key=str.lower)
    services = sorted({r["service"] for r in rows})
    si, ci, vi = ({v: i for i, v in enumerate(x)} for x in (staff, clients, services))
    data = {
        "generated": datetime.now().strftime("%d/%m/%Y"),
        "staff": staff, "clients": clients, "services": services,
        "rates": {}, "notes": notes,
        "rows": [[si[r["staff"]], r["date"], ci[r["client"]], vi[r["service"]], r["hours"], r["kind"], r["desc"]]
                 for r in sorted(rows, key=lambda r: (r["date"], r["staff"]))],
    }
    tpl = (HERE / "Sistema.html").read_text(encoding="utf-8")
    payload = json.dumps(data, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    Path(out).write_text(tpl.replace("/*DATA*/null", payload, 1), encoding="utf-8")

    print(f"{len(rows)} lançamentos · {len(staff)} pessoas · {len(clients)} clientes · ignorados: {dict(skipped)}")


if __name__ == "__main__":
    if len(sys.argv) < 3:
        sys.exit(__doc__)
    main(*sys.argv[1:5])
