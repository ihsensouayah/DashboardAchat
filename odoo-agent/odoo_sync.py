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


# ------------------------------------------------------------------ main
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--feed", action="append")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--force", action="store_true")
    ap.add_argument("--discover", metavar="MODEL")
    a = ap.parse_args()

    odoo = Odoo()
    print(f"✅ Connecté à Odoo {odoo.url} (base {odoo.db}, uid {odoo.uid})")

    if a.discover:
        for k, v in sorted(odoo.fields(a.discover).items(), key=lambda kv: kv[1].get("string", "")):
            rel = f" -> {v['relation']}" if v.get("relation") else ""
            print(f"  {v.get('string', ''):<40} {k:<35} {v['type']}{rel}")
        return

    cfg = json.load(open(os.path.join(HERE, "config.json"), encoding="utf-8"))
    feeds = {k: v for k, v in cfg["feeds"].items() if v.get("enabled", True) and (not a.feed or k in a.feed)}
    db = fs = None
    if not a.dry_run:
        db, fs = firestore_client()

    errors = 0
    for name, feed in feeds.items():
        try:
            rows = read_rows(odoo, feed)
            print(f"  • {name} : {len(rows)} ligne(s) lue(s) dans Odoo")
            if a.dry_run:
                print(f"  • {name} : {len(rows)} lignes")
                for r in rows[:5]:
                    print("     ", json.dumps(r, ensure_ascii=False, default=str))
            else:
                publish(db, fs, name, rows, a.force)
        except Exception as e:  # un flux en erreur n'empêche pas les autres
            errors += 1
            print(f"  ✗ {name} : {e}")
    sys.exit(1 if errors else 0)


if __name__ == "__main__":
    main()

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


def read_rows(odoo, feed):
    model = feed["model"]
    domain = list(feed.get("domain", []))
    if feed.get("days_back"):
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


# ------------------------------------------------------------------ main
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--feed", action="append")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--force", action="store_true")
    ap.add_argument("--discover", metavar="MODEL")
    a = ap.parse_args()

    odoo = Odoo()
    print(f"✅ Connecté à Odoo {odoo.url} (base {odoo.db}, uid {odoo.uid})")

    if a.discover:
        for k, v in sorted(odoo.fields(a.discover).items(), key=lambda kv: kv[1].get("string", "")):
            rel = f" -> {v['relation']}" if v.get("relation") else ""
            print(f"  {v.get('string', ''):<40} {k:<35} {v['type']}{rel}")
        return

    cfg = json.load(open(os.path.join(HERE, "config.json"), encoding="utf-8"))
    feeds = {k: v for k, v in cfg["feeds"].items() if v.get("enabled", True) and (not a.feed or k in a.feed)}
    db = fs = None
    if not a.dry_run:
        db, fs = firestore_client()

    errors = 0
    for name, feed in feeds.items():
        try:
            rows = read_rows(odoo, feed)
            if a.dry_run:
                print(f"  • {name} : {len(rows)} lignes")
                for r in rows[:5]:
                    print("     ", json.dumps(r, ensure_ascii=False, default=str))
            else:
                publish(db, fs, name, rows, a.force)
        except Exception as e:  # un flux en erreur n'empêche pas les autres
            errors += 1
            print(f"  ✗ {name} : {e}")
    sys.exit(1 if errors else 0)


if __name__ == "__main__":
    main()
