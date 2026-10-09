"""Sauvegarde complète des deux bases Firestore (autopachat + autop-reception).

Chaque document est écrit sur une ligne JSON : {"path": "orders/abc", "data": {...}}.
Les collections recalculables (copies compactes, présence, diagnostic) sont ignorées
pour économiser le quota. Le workflow chiffre ensuite les fichiers avec un mot de passe.
Ce script ne fait QUE lire : il ne modifie et n'efface rien.
"""
import base64
import datetime
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from odoo_sync import firestore_client, reception_client  # noqa: E402

# Recalculables ou sans valeur métier : non sauvegardées
SKIP = {"presence", "diagReads", "dailyStats", "odooSync", "_bundles", "_probe"}
OUT = os.environ.get("BACKUP_DIR", os.path.join(HERE, "backup"))


def enc(v):
    if isinstance(v, datetime.datetime):
        return {"__ts": v.isoformat()}
    if isinstance(v, bytes):
        return {"__bytes": base64.b64encode(v).decode()}
    if hasattr(v, "latitude") and hasattr(v, "longitude"):
        return {"__geo": [v.latitude, v.longitude]}
    if hasattr(v, "path") and hasattr(v, "id"):
        return {"__ref": v.path}
    return str(v)


def dump(db, base_name):
    os.makedirs(OUT, exist_ok=True)
    total, counts = 0, {}
    for coll in db.collections():
        if coll.id in SKIP or coll.id.startswith("_deletions_"):
            continue
        n = 0
        with open(os.path.join(OUT, f"{base_name}__{coll.id}.jsonl"), "w", encoding="utf-8") as f:
            for doc in coll.stream():
                f.write(json.dumps({"path": doc.reference.path, "data": doc.to_dict()},
                                   ensure_ascii=False, default=enc) + "\n")
                n += 1
        counts[coll.id] = n
        total += n
    print(f"  ✓ {base_name} : {total} document(s) sauvegardé(s) — " +
          ", ".join(f"{k} {v}" for k, v in sorted(counts.items())))
    return total


def main():
    errors = 0
    # La base principale d'abord : sa connexion doit être la connexion « par défaut » de Firebase
    for name, get in (("autopachat", lambda: firestore_client()[0]), ("autop-reception", reception_client)):
        try:
            db = get()
            if db is None:
                print(f"  · {name} : clé absente, base ignorée")
                continue
            dump(db, name)
        except Exception as e:  # une base en erreur n'empêche pas l'autre
            errors += 1
            print(f"  ✗ {name} : {e}")
    with open(os.path.join(OUT, "_info.txt"), "w", encoding="utf-8") as f:
        f.write(f"Sauvegarde du {datetime.datetime.now(datetime.timezone.utc).isoformat()} (UTC)\n")
    sys.exit(1 if errors else 0)


if __name__ == "__main__":
    main()
