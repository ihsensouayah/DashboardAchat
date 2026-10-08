#!/usr/bin/env python3
"""
Agent de synchronisation Odoo -> Firestore (AUTOP Tunisie).

Lit les modèles Odoo définis dans config.json (XML-RPC, clé API), produit des
lignes au MÊME format que les exports Excel Odoo (mêmes en-têtes de colonnes),
et les dépose dans Firestore :

    odooSync/{feed}                 -> méta : syncId, syncedAt, rowCount, chunks, hash
    odooSync/{feed}/chunks/{0..n}   -> { rows: [...] }

Le dashboard applique ensuite ces lignes avec ses fonctions d'import
existantes (mêmes règles métier que l'import manuel du fichier Excel).

Usage :
    python odoo_sync.py                 # synchronise tous les flux actifs
    python odoo_sync.py --feed reception_bc
    python odoo_sync.py --dry-run       # affiche les lignes sans écrire dans Firestore
    python odoo_sync.py --discover project.project   # liste les champs (nom technique <-> libellé)

Variables d'environnement (secrets GitHub) :
    ODOO_URL, ODOO_DB, ODOO_USER, ODOO_API_KEY
    (Odoo 13 n'a pas de clés API : ODOO_API_KEY contient alors le mot de passe de l'utilisateur)
    FIREBASE_SERVICE_ACCOUNT   (contenu JSON du compte de service Firebase)
"""
import argparse
import ast
import hashlib
import json
import os
import re
import sys
import unicodedata
import uuid
import xmlrpc.client
from datetime import datetime, timedelta, timezone
from zoneinfo import ZoneInfo

HERE = os.path.dirname(os.path.abspath(__file__))
TZ = ZoneInfo("Africa/Tunis")
CHUNK_ROWS = 400


def norm(s):
    s = unicodedata.normalize("NFD", str(s or ""))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return " ".join(s.lower().split())


def env(name):
    v = os.environ.get(name, "").strip()
    if not v:
        sys.exit(f"❌ Variable d'environnement manquante : {name}")
    return v


# ----------------------------------------------------------------- Odoo
class Odoo:
    def __init__(self):
        self.url = env("ODOO_URL").rstrip("/")
        self.db = env("ODOO_DB")
        user = env("ODOO_USER")
        self.key = env("ODOO_API_KEY")
        common = xmlrpc.client.ServerProxy(f"{self.url}/xmlrpc/2/common", allow_none=True)
        self.uid = common.authenticate(self.db, user, self.key, {})
        if not self.uid:
            sys.exit("❌ Authentification Odoo refusée (vérifier ODOO_DB / ODOO_USER / ODOO_API_KEY).")
        self.models = xmlrpc.client.ServerProxy(f"{self.url}/xmlrpc/2/object", allow_none=True)
        self.ctx = {"lang": "fr_FR", "tz": "Africa/Tunis"}
        self._fields = {}

    def call(self, model, method, args, kw=None):
        kw = dict(kw or {})
        kw.setdefault("context", self.ctx)
        return self.models.execute_kw(self.db, self.uid, self.key, model, method, args, kw)

    def fields(self, model):
        if model not in self._fields:
            self._fields[model] = self.call(model, "fields_get", [], {"attributes": ["string", "type", "relation", "selection"]})
        return self._fields[model]

    def resolve(self, model, segment):
        """Nom technique ou libellé (tel qu'il apparaît dans l'export Excel) -> nom technique."""
        f = self.fields(model)
        if segment in f:
            return segment
        target = norm(segment)
        hits = [k for k, v in f.items() if norm(v.get("string")) == target]
        if len(hits) == 1:
            return hits[0]
        if len(hits) > 1:
            stored = [h for h in hits if not h.startswith("x_studio_")] or hits
            return sorted(stored)[0]
        raise KeyError(f"Champ « {segment} » introuvable sur {model}. Lancez : python odoo_sync.py --discover {model}")


def fmt_value(odoo, model, fname, value, raw=False):
    meta = odoo.fields(model)[fname]
    t = meta["type"]
    if value is False or value is None:
        return "" if t not in ("float", "monetary", "integer") else None
    if t == "many2one":
        return value[1] if isinstance(value, (list, tuple)) else value
    if t == "selection":
        if raw:
            return value
        return dict(meta.get("selection") or []).get(value, value)
    if t == "datetime":
        dt = datetime.strptime(value, "%Y-%m-%d %H:%M:%S").replace(tzinfo=timezone.utc).astimezone(TZ)
        return dt.strftime("%Y-%m-%d %H:%M:%S")
    return value


TILE_MODELS = ["ks_dashboard_ninja.item"]  # tuiles du « Dashboard Manager » (Ksolves Dashboard Ninja)


def tile_domain(odoo, tile_name):
    """Lit une tuile du Dashboard Manager et renvoie (modèle, filtre) : l'agent applique
    exactement le même filtre que la tuile affichée dans Odoo."""
    last_err = None
    for tm in TILE_MODELS:
        try:
            f = odoo.fields(tm)
        except Exception as e:
            last_err = e
            continue
        want = [x for x in ("name", "ks_domain", "ks_model_id", "ks_model_name") if x in f]
        recs = odoo.call(tm, "search_read", [[["name", "=", tile_name]]], {"fields": want, "limit": 5})
        if not recs:
            recs = odoo.call(tm, "search_read", [[["name", "ilike", tile_name]]], {"fields": want, "limit": 5})
        if not recs:
            raise KeyError(f"Tuile « {tile_name} » introuvable dans {tm}")
        r = recs[0]
        model = r.get("ks_model_name") or ""
        if not model and r.get("ks_model_id"):
            model = odoo.call("ir.model", "read", [[r["ks_model_id"][0]]], {"fields": ["model"]})[0]["model"]
        raw = (r.get("ks_domain") or "[]").strip() or "[]"
        src = raw.replace('"%UID"', str(odoo.uid)).replace("'%UID'", str(odoo.uid)).replace("%UID", str(odoo.uid))
        try:
            dom = ast.literal_eval(src)
        except Exception:
            raise ValueError(f"Filtre de la tuile non lisible automatiquement : {raw}")
        print(f"  ↳ tuile « {r.get('name')} » : modèle {model}, filtre {dom}")
        return model, [list(x) if isinstance(x, (list, tuple)) else x for x in dom]
    raise RuntimeError(f"Module de tableaux de bord introuvable ({last_err})")


def read_rows(odoo, feed):
    model = feed["model"]
    domain = list(feed.get("domain", []))
    if feed.get("domain_from_tile"):
        tile_model, tile_dom = tile_domain(odoo, feed["domain_from_tile"])
        model = tile_model or model
        domain = tile_dom + domain
    if feed.get("since"):
        # Date de départ fixe (heure de Tunis) : seuls les enregistrements à partir de ce jour
        start = datetime.strptime(feed["since"], "%Y-%m-%d").replace(tzinfo=TZ).astimezone(timezone.utc)
        field = feed.get("date_field", "create_date")
        ftype = odoo.fields(model).get(field, {}).get("type")
        domain.append([field, ">=", feed["since"] if ftype == "date" else start.strftime("%Y-%m-%d %H:%M:%S")])
    elif feed.get("days_back"):
        since = (datetime.now(timezone.utc) - timedelta(days=int(feed["days_back"]))).strftime("%Y-%m-%d %H:%M:%S")
        domain.append([feed.get("date_field", "create_date"), ">=", since])

    # Résolution des colonnes : "source" = chemin séparé par "/" (libellés ou noms techniques)
    cols = []
    for c in feed["columns"]:
        if c.get("compute"):
            cols.append({**c, "chain": []})
            continue
        parts = c["source"].split("/")
        first = odoo.resolve(model, parts[0])
        chain, cur_model, cur_field = [first], model, first
        for seg in parts[1:]:
            rel = odoo.fields(cur_model)[cur_field].get("relation")
            if not rel:
                raise KeyError(f"« {cur_field} » n'est pas un champ relationnel ({c['source']})")
            cur_model = rel
            cur_field = odoo.resolve(cur_model, seg)
            chain.append(cur_field)
        cols.append({**c, "chain": chain})

    top_fields = sorted({c["chain"][0] for c in cols if c["chain"]})
    ids = odoo.call(model, "search", [domain], {"order": feed.get("order", "id desc"), "limit": feed.get("limit", 0)})
    recs = []
    for i in range(0, len(ids), 500):
        recs += odoo.call(model, "read", [ids[i:i + 500]], {"fields": top_fields})

    # Champs relationnels imbriqués : on lit les enregistrements liés par lots
    related_cache = {}  # (model, id) -> record

    def fetch_related(rel_model, rel_ids, fields):
        todo = [i for i in set(rel_ids) if (rel_model, i) not in related_cache]
        for i in range(0, len(todo), 500):
            for r in odoo.call(rel_model, "read", [todo[i:i + 500]], {"fields": fields}):
                related_cache[(rel_model, r["id"])] = r

    computed = {}
    for c in cols:
        if c.get("compute"):
            computed[c["header"]] = COMPUTES[c["compute"]](odoo, recs)
    for c in cols:
        if len(c["chain"]) < 2:
            continue
        cur_model, values = model, [r.get(c["chain"][0]) for r in recs]
        for depth, fname in enumerate(c["chain"][:-1]):
            rel = odoo.fields(cur_model)[fname]["relation"]
            nxt = c["chain"][depth + 1]
            rel_ids = [v[0] for v in values if isinstance(v, (list, tuple)) and v]
            fetch_related(rel, rel_ids, [nxt])
            values = [related_cache.get((rel, v[0]), {}).get(nxt) if isinstance(v, (list, tuple)) and v else False for v in values]
            cur_model = rel
        c["_values"], c["_model"] = values, cur_model

    rows = []
    for idx, r in enumerate(recs):
        row = {}
        for c in cols:
            if c.get("compute"):
                row[c["header"]] = computed[c["header"]].get(r["id"], "")
                continue
            if len(c["chain"]) == 1:
                v = fmt_value(odoo, model, c["chain"][0], r.get(c["chain"][0]), c.get("raw"))
            else:
                v = fmt_value(odoo, c["_model"], c["chain"][-1], c["_values"][idx], c.get("raw"))
            row[c["header"]] = v
        rows.append(row)
    return rows


# ---------------------------------------------------- colonnes calculées
def po_receipt_status(odoo, recs):
    """Statut de réception d'un bon de commande : pending / partial / full.
    Odoo 16+ a le champ receipt_status ; en Odoo 13 on le calcule à partir des
    lignes (quantité reçue vs quantité commandée)."""
    ids = [r["id"] for r in recs]
    if not ids:
        return {}
    if "receipt_status" in odoo.fields("purchase.order"):
        out = {}
        for i in range(0, len(ids), 500):
            for r in odoo.call("purchase.order", "read", [ids[i:i + 500]], {"fields": ["receipt_status"]}):
                out[r["id"]] = r["receipt_status"] or ""
        return out
    lf = odoo.fields("purchase.order.line")
    fields = ["order_id", "product_qty", "qty_received"] + (["display_type"] if "display_type" in lf else [])
    agg = {}  # order_id -> [commandé, reçu]
    for i in range(0, len(ids), 200):
        lines = odoo.call("purchase.order.line", "search_read", [[["order_id", "in", ids[i:i + 200]]]], {"fields": fields})
        for l in lines:
            if l.get("display_type") or not l.get("product_qty"):
                continue
            a = agg.setdefault(l["order_id"][0], [0.0, 0.0])
            a[0] += l["product_qty"]
            a[1] += min(l["qty_received"] or 0.0, l["product_qty"])
    out = {}
    for oid in ids:
        ordered, received = agg.get(oid, [0.0, 0.0])
        if ordered <= 0:
            out[oid] = ""
        elif received <= 0:
            out[oid] = "pending"
        elif received + 1e-6 >= ordered:
            out[oid] = "full"
        else:
            out[oid] = "partial"
    return out


COMPUTES = {"po_receipt_status": po_receipt_status}


# ------------------------------------------------------------- Firestore
def firestore_client():
    import firebase_admin
    from firebase_admin import credentials, firestore
    sa_file = os.environ.get("FIREBASE_SERVICE_ACCOUNT_FILE", "").strip()
    sa = json.load(open(sa_file, encoding="utf-8")) if sa_file else json.loads(env("FIREBASE_SERVICE_ACCOUNT"))
    if not firebase_admin._apps:
        firebase_admin.initialize_app(credentials.Certificate(sa))
    return firestore.client(), firestore


def publish(db, fs, feed_name, rows, force=False):
    payload = json.dumps(rows, ensure_ascii=False, sort_keys=True, default=str)
    digest = hashlib.sha256(payload.encode()).hexdigest()
    meta_ref = db.collection("odooSync").document(feed_name)
    meta = meta_ref.get()
    if meta.exists and meta.to_dict().get("hash") == digest and not force:
        meta_ref.update({"checkedAt": fs.SERVER_TIMESTAMP})
        print(f"  = {feed_name} : aucune modification dans Odoo ({len(rows)} lignes)")
        return
    chunks = [rows[i:i + CHUNK_ROWS] for i in range(0, len(rows), CHUNK_ROWS)] or [[]]
    old_chunks = (meta.to_dict() or {}).get("chunks", 0) if meta.exists else 0
    batch = db.batch()
    for i, ch in enumerate(chunks):
        batch.set(meta_ref.collection("chunks").document(str(i)), {"rows": ch})
    for i in range(len(chunks), old_chunks):
        batch.delete(meta_ref.collection("chunks").document(str(i)))
    batch.set(meta_ref, {
        "syncId": uuid.uuid4().hex,
        "hash": digest,
        "rowCount": len(rows),
        "chunks": len(chunks),
        "syncedAt": fs.SERVER_TIMESTAMP,
        "checkedAt": fs.SERVER_TIMESTAMP,
        "source": "odoo-agent",
    })
    batch.commit()
    print(f"  ✓ {feed_name} : {len(rows)} lignes publiées ({len(chunks)} bloc(s))")


# ------------------------------------------- Réception BC : écriture directe
# Écrit les bons de commande Odoo directement dans la base dédiée à Réception BC
# (projet Firebase autop-reception, collection receptionBC, 1 document par BC).
# Règles (ajout seulement) :
#   - un BC absent est AJOUTÉ avec le statut de réception calculé dans Odoo ;
#   - un BC déjà suivi n'est mis à jour QUE s'il n'a jamais été modifié à la main
#     (manual absent) et que son statut est « non » ou « partiel » : il passe alors au
#     statut plus avancé indiqué par Odoo (partiel / recu). Rien d'autre n'est touché
#     (contrôle magasin, commentaire, corbeille, choix manuels).
RECEPTION_EXCLUDED = ["mehdi hajjaji", "saifallah chaouachi", "amine ben omrane"]
RECEPTION_RANK = {"non": 0, "partiel": 1, "recu": 2}


def reception_client():
    raw = os.environ.get("FIREBASE_RECEPTION_SERVICE_ACCOUNT", "").strip()
    if not raw:
        return None
    import firebase_admin
    from firebase_admin import credentials, firestore
    try:
        app = firebase_admin.get_app("reception")
    except ValueError:
        app = firebase_admin.initialize_app(credentials.Certificate(json.loads(raw)), name="reception")
    return firestore.client(app)


def _rbc_status(row):
    if str(row.get("État", "")).strip() == "cancel":
        return "ferme"
    return {"pending": "non", "partial": "partiel", "full": "recu"}.get(str(row.get("Statut réception", "")).strip())


def reception_direct(rows):
    rdb = reception_client()
    if rdb is None:
        print("  · reception_bc direct : clé FIREBASE_RECEPTION_SERVICE_ACCOUNT absente, étape ignorée")
        return
    today = datetime.now(TZ).strftime("%Y-%m-%d")
    coll = rdb.collection("receptionBC")
    # On ne lit QUE les BC présents dans l'export Odoo (quelques centaines), pas toute la collection :
    # chaque passage coûte ~1 lecture par BC exporté, quel que soit l'historique stocké.
    ids = sorted({re.sub(r"[^A-Za-z0-9_-]", "-", str(r.get("Référence commande", "")).strip())
                  for r in rows if str(r.get("Référence commande", "")).strip().upper().startswith("PO/")})
    existing = {}
    for i in range(0, len(ids), 300):
        for d in rdb.get_all([coll.document(x) for x in ids[i:i + 300]], field_paths=["status", "manual", "deleted"]):
            if d.exists:
                existing[d.id] = d.to_dict() or {}
    batch, ops, added, updated = rdb.batch(), 0, 0, 0
    for r in rows:
        ref = str(r.get("Référence commande", "")).strip()
        m = re.match(r"(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})", str(r.get("Date de la commande", "")))
        if not ref.upper().startswith("PO/") or not m:
            continue
        resp = str(r.get("Responsable achats", "") or "")
        if any(n in norm(resp) for n in RECEPTION_EXCLUDED):
            continue
        doc_id = re.sub(r"[^A-Za-z0-9_-]", "-", ref)
        st = _rbc_status(r)
        cur = existing.get(doc_id)
        if cur is None:
            total = r.get("Total")
            batch.set(coll.document(doc_id), {
                "day": f"{m[1]}-{m[2]}-{m[3]}", "deleted": False, "ref": ref, "time": f"{m[4]}:{m[5]}",
                "fournisseur": str(r.get("Fournisseur", "") or ""), "responsable": resp,
                "origine": str(r.get("Document d'origine", "") or ""),
                "total": round(float(total), 3) if isinstance(total, (int, float)) else None,
                "status": st or "non", "recu": today if st and st != "non" else None,
                "note": "", "manual": None, "mag": None, "magDate": None,
            })
            added += 1
        elif (not cur.get("deleted") and not cur.get("manual") and st in ("partiel", "recu")
              and RECEPTION_RANK.get(cur.get("status") or "non", 9) < RECEPTION_RANK[st]):
            batch.update(coll.document(doc_id), {"status": st, "recu": today})
            updated += 1
        else:
            continue
        ops += 1
        if ops >= 400:
            batch.commit(); batch, ops = rdb.batch(), 0
    if ops:
        batch.commit()
    print(f"  ✓ reception_bc direct (autop-reception) : {added} BC ajouté(s), {updated} statut(s) mis à jour depuis Odoo")


# ------------------------------------------------- copies compactes (quota)
# Le dashboard lit ces copies (quelques documents) au lieu de relire des milliers de
# documents à chaque ouverture, puis n'écoute que les changements (champ _syncAt).
BUNDLE_COLLECTIONS = ["orders", "repairDossiers", "controleMarge", "margeData"]
BUNDLE_EVERY_HOURS = 6
BUNDLE_CHUNK_BYTES = 700_000


def _enc(v):
    if isinstance(v, datetime):
        ts = v.timestamp()
        sec = int(ts // 1)
        return {"__ts": [sec, int(round((ts - sec) * 1e9))]}
    if isinstance(v, dict):
        return {k: _enc(x) for k, x in v.items()}
    if isinstance(v, (list, tuple)):
        return [_enc(x) for x in v]
    if hasattr(v, "latitude") and hasattr(v, "longitude"):
        return {"latitude": v.latitude, "longitude": v.longitude}
    if hasattr(v, "path") and hasattr(v, "id") and not isinstance(v, str):
        return v.path
    if isinstance(v, bytes):
        return None
    return v


def build_bundles(db, force=False):
    for name in BUNDLE_COLLECTIONS:
        meta_ref = db.collection("_bundles").document(name)
        meta = meta_ref.get()
        old = meta.to_dict() if meta.exists else {}
        built = old.get("builtAt")
        if not force and built and (datetime.now(timezone.utc) - built).total_seconds() < BUNDLE_EVERY_HOURS * 3600:
            print(f"  = copie {name} : récente, pas de reconstruction")
            continue
        built_at = datetime.now(timezone.utc) - timedelta(minutes=2)  # marge : tout ce qui change pendant la lecture sera aussi dans les « changements »
        docs = [{"id": d.id, "d": _enc(d.to_dict())} for d in db.collection(name).stream()]
        chunks, cur, size = [], [], 2
        for e in docs:
            s = len(json.dumps(e, ensure_ascii=False, default=str).encode()) + 1
            if cur and size + s > BUNDLE_CHUNK_BYTES:
                chunks.append(cur); cur, size = [], 2
            cur.append(e); size += s
        chunks.append(cur)
        batch = db.batch()
        for i, ch in enumerate(chunks):
            batch.set(meta_ref.collection("chunks").document(str(i)), {"json": json.dumps(ch, ensure_ascii=False, default=str)})
        for i in range(len(chunks), int(old.get("chunks", 0) or 0)):
            batch.delete(meta_ref.collection("chunks").document(str(i)))
        batch.commit()
        meta_ref.set({"builtAt": built_at, "chunks": len(chunks), "count": len(docs)})
        # Traces de suppression antérieures à la copie : devenues inutiles
        n_del = 0
        for d in db.collection("_deletions_" + name).where("at", "<", built_at).stream():
            d.reference.delete(); n_del += 1
        print(f"  ✓ copie {name} : {len(docs)} documents en {len(chunks)} bloc(s)" + (f", {n_del} trace(s) nettoyée(s)" if n_del else ""))


# ------------------------------------------------------------------ main
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--feed", action="append")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--force", action="store_true")
    ap.add_argument("--discover", metavar="MODEL")
    ap.add_argument("--bundles-only", action="store_true", help="reconstruit seulement les copies compactes")
    ap.add_argument("--force-bundles", action="store_true")
    ap.add_argument("--reception-only", action="store_true",
                    help="met à jour seulement Réception BC (base autop-reception), sans toucher la base principale")
    a = ap.parse_args()
    if a.reception_only:
        odoo = Odoo()
        print(f"✅ Connecté à Odoo {odoo.url} (base {odoo.db}, uid {odoo.uid}) — mode Réception BC seule")
        cfg = json.load(open(os.path.join(HERE, "config.json"), encoding="utf-8"))
        rows = read_rows(odoo, cfg["feeds"]["reception_bc"])
        print(f"  • reception_bc : {len(rows)} ligne(s) lue(s) dans Odoo")
        reception_direct(rows)
        return

    errors = 0
    db = fs = None
    if not a.dry_run and not a.discover:
        db, fs = firestore_client()
        # Copies compactes d'abord : elles ne dépendent pas d'Odoo
        try:
            build_bundles(db, force=a.force_bundles)
        except Exception as e:
            errors += 1
            print(f"  ✗ copies compactes : {e}")
        if a.bundles_only:
            sys.exit(1 if errors else 0)

    odoo = Odoo()
    print(f"✅ Connecté à Odoo {odoo.url} (base {odoo.db}, uid {odoo.uid})")

    if a.discover:
        for k, v in sorted(odoo.fields(a.discover).items(), key=lambda kv: kv[1].get("string", "")):
            rel = f" -> {v['relation']}" if v.get("relation") else ""
            print(f"  {v.get('string', ''):<40} {k:<35} {v['type']}{rel}")
        return

    cfg = json.load(open(os.path.join(HERE, "config.json"), encoding="utf-8"))
    feeds = {k: v for k, v in cfg["feeds"].items() if v.get("enabled", True) and (not a.feed or k in a.feed)}
    for name, feed in feeds.items():
        try:
            rows = read_rows(odoo, feed)
            print(f"  • {name} : {len(rows)} ligne(s) lue(s) dans Odoo")
            if a.dry_run:
                for r in rows[:5]:
                    print("     ", json.dumps(r, ensure_ascii=False, default=str))
            else:
                # Réception BC a sa propre base (autop-reception) : on l'alimente D'ABORD,
                # pour qu'elle reste à jour même si la base principale a atteint son quota.
                if name == "reception_bc":
                    try:
                        reception_direct(rows)
                    except Exception as e:
                        errors += 1
                        print(f"  ✗ reception_bc direct (autop-reception) : {e}")
                publish(db, fs, name, rows, a.force)
        except Exception as e:  # un flux en erreur n'empêche pas les autres
            errors += 1
            print(f"  ✗ {name} : {e}")
    sys.exit(1 if errors else 0)


if __name__ == "__main__":
    main()
