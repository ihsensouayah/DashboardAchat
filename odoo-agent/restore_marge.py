"""Restaure la collection margeData (Analyse Marge) depuis une sauvegarde chiffrée.

Usage (dans GitHub Actions, après déchiffrement de la sauvegarde) :
    python restore_marge.py --backup-dir DIR --safety-dir DIR2            # simulation : n'écrit rien
    python restore_marge.py --backup-dir DIR --safety-dir DIR2 --apply    # restauration réelle

Avant toute écriture, l'état ACTUEL de margeData est copié dans --safety-dir (copie de sécurité,
chiffrée ensuite par le workflow) : la restauration elle-même peut donc être annulée.
"""
import argparse
import collections
import datetime
import glob
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import odoo_sync  # noqa: E402
from backup import enc  # noqa: E402

COLL = "margeData"


def dec(v):
    if isinstance(v, dict):
        if set(v) == {"__ts"}:
            return datetime.datetime.fromisoformat(v["__ts"])
        return {k: dec(x) for k, x in v.items()}
    if isinstance(v, list):
        return [dec(x) for x in v]
    return v


def by_month(docs):
    c = collections.Counter(str((d or {}).get("mois")) for d in docs.values())
    return ", ".join(f"{k}: {v}" for k, v in sorted(c.items()))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--backup-dir", required=True)
    ap.add_argument("--safety-dir", required=True)
    ap.add_argument("--apply", action="store_true")
    a = ap.parse_args()

    files = glob.glob(os.path.join(a.backup_dir, "**", f"autopachat__{COLL}.jsonl"), recursive=True)
    if not files:
        sys.exit(f"❌ Fichier autopachat__{COLL}.jsonl introuvable dans la sauvegarde.")
    backup = {}
    with open(files[0], encoding="utf-8") as f:
        for line in f:
            if line.strip():
                e = json.loads(line)
                backup[e["path"].split("/", 1)[1]] = dec(e["data"])
    if len(backup) < 100:
        sys.exit(f"❌ Sauvegarde suspecte : seulement {len(backup)} ligne(s). Restauration refusée.")

    db, _ = odoo_sync.firestore_client()
    current = {d.id: d.to_dict() for d in db.collection(COLL).stream()}

    os.makedirs(a.safety_dir, exist_ok=True)
    with open(os.path.join(a.safety_dir, f"avant-restauration__{COLL}.jsonl"), "w", encoding="utf-8") as f:
        for i, d in current.items():
            f.write(json.dumps({"path": f"{COLL}/{i}", "data": d}, ensure_ascii=False, default=enc) + "\n")

    to_delete = [i for i in current if i not in backup]
    print(f"Sauvegarde : {len(backup)} ligne(s) — {by_month(backup)}")
    print(f"Actuel     : {len(current)} ligne(s) — {by_month(current)}")
    print(f"À supprimer (absentes de la sauvegarde) : {len(to_delete)} — à (re)écrire : {len(backup)}")

    if not a.apply:
        print("ℹ️  SIMULATION : rien n'a été modifié. Relancer avec « appliquer » coché pour restaurer.")
        return

    ops, batch = 0, db.batch()
    for i in to_delete:
        batch.delete(db.collection(COLL).document(i)); ops += 1
        if ops >= 400:
            batch.commit(); batch, ops = db.batch(), 0
    for i, d in backup.items():
        batch.set(db.collection(COLL).document(i), d); ops += 1
        if ops >= 400:
            batch.commit(); batch, ops = db.batch(), 0
    if ops:
        batch.commit()
    print(f"✓ Restauration faite : {len(to_delete)} supprimée(s), {len(backup)} réécrite(s)")

    odoo_sync.BUNDLE_COLLECTIONS = [COLL]
    odoo_sync.build_bundles(db, force=True)
    print("✓ Copie compacte Analyse Marge reconstruite — demander à chacun de faire Ctrl + F5")


if __name__ == "__main__":
    main()
