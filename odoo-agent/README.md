# Agent de synchro Odoo → Dashboards AUTOP

Toutes les 30 min (lun–sam, 7h–19h), l'agent lit Odoo et met à jour les dashboards,
sans import Excel manuel.

```
Odoo ──(API, lecture seule)──> GitHub Actions (odoo_sync.py) ──> Firestore odooSync/*
                                                                     │
                     index.html (Réparation, Facturation) <──────────┤ mêmes fonctions d'import
                     reception-bc (Réception BC)          <──────────┘ que les fichiers Excel
```

| Flux | Modèle Odoo | Dashboard | État |
|---|---|---|---|
| `reception_bc` | `purchase.order` (60 derniers jours) | Réception BC : nouvelles commandes + statut de réception automatique | **actif** |
| `reparation` | `project.project` | Réparation (= bouton « Importer Odoo ») | à activer après test |
| `facturation` | à confirmer | Fin réparation (= import facturation) | à configurer |

## 0. Odoo 13 : points spécifiques
- **Pas de clé API en Odoo 13** : `ODOO_API_KEY` contient le **mot de passe** de l'utilisateur. Un compte dédié (créé par l'admin) est donc fortement recommandé ; si le mot de passe change, mettre à jour le secret.
- Le statut de réception est calculé à partir des lignes du BC (quantité reçue / commandée).
- **Odoo doit être joignable depuis Internet** pour GitHub Actions. Test : ouvrir l'adresse Odoo sur un téléphone en 4G (Wi-Fi coupé).
  - Si ça s'ouvre → GitHub Actions (étapes ci-dessous).
  - Si ça ne s'ouvre pas (Odoo seulement sur le réseau du bureau) → **mode PC du bureau** (voir en bas).

## 1. Côté Odoo : créer l'accès de l'agent
1. Créer un utilisateur dédié, par ex. **« Agent Dashboard »**, avec les droits **lecture seule** sur Achats et Projet.
2. Odoo 13 : utiliser le mot de passe de ce compte comme `ODOO_API_KEY`.

⚠️ Ne pas envoyer la clé ni le mot de passe par chat ou par mail : ils vont seulement dans les secrets GitHub (étape 3).

## 2. Côté Firebase : compte de service
Console Firebase → projet `autopachat` → ⚙️ Paramètres → *Comptes de service* → **Générer une nouvelle clé privée** (fichier JSON).

## 3. Côté GitHub (dépôt AshZehiri1)
1. Copier le dossier `odoo-agent/` et le fichier `.github/workflows/odoo-sync.yml` à la racine du dépôt.
2. *Settings → Secrets and variables → Actions → New repository secret* :

| Secret | Valeur |
|---|---|
| `ODOO_URL` | ex. `https://autop.odoo.com` |
| `ODOO_DB` | nom de la base (visible dans *Paramètres → Activer le mode développeur*, ou la page `/web/database/selector`) |
| `ODOO_USER` | login de l'utilisateur « Agent Dashboard » |
| `ODOO_API_KEY` | le mot de passe du compte (Odoo 13) |
| `FIREBASE_SERVICE_ACCOUNT` | tout le contenu du fichier JSON de l'étape 2 |

3. Onglet *Actions* → « Synchro Odoo -> Dashboards » → **Run workflow** pour un premier test.

## 4. Activer le flux Réparation
En local (avec les mêmes variables d'environnement) :
```
python odoo_sync.py --discover project.project   # liste libellé <-> nom technique des champs
python odoo_sync.py --feed reparation --dry-run  # affiche les lignes sans rien écrire
```
Mettre dans `domain` le même filtre que celui de l'export Excel habituel (par ex. `[["stage_id.name","!=","Clôturé"]]`),
puis passer `"enabled": true` dans `config.json`.

## Règles de fonctionnement
- **Réception BC** : un statut choisi à la main dans la liste déroulante n'est jamais écrasé par Odoo.
  Correspondance : `pending` → Non réceptionnée, `partial` → Partielle, `full` → Réceptionnée, commande annulée → Commande fermée.
- **Réparation** : mêmes règles que l'import manuel (un statut modifié à la main reste tel quel ; le devis validé et les heures allouées ne sont jamais modifiés).
- L'agent ne publie rien si Odoo n'a pas changé ; un seul navigateur applique chaque synchro (pas de doublon).
- Aucune écriture dans Odoo : l'agent ne fait que lire.
- GitHub suspend les tâches planifiées d'un dépôt resté 60 jours sans commit : il suffit alors de le réactiver dans l'onglet *Actions*.

## Mode PC du bureau (Odoo non joignable depuis Internet)
1. Sur un PC du bureau allumé en journée : installer Python 3.12 (python.org, cocher « Add to PATH »), puis `pip install firebase-admin`.
2. Copier le dossier `odoo-agent` sur ce PC (ex. `C:\odoo-agent`).
3. Renommer `secrets.env.exemple` en `secrets.env` et le remplir ; mettre le JSON Firebase à côté sous le nom `firebase.json`.
4. Double-cliquer `run_local.bat` → vérifier `sync.log`.
5. *Planificateur de tâches Windows → Créer une tâche de base* → Quotidien → répéter toutes les 30 min pendant 12 h → action : `C:\odoo-agent\run_local.bat`.
6. Dans ce cas, ne pas créer le workflow GitHub (ou le désactiver dans l'onglet Actions).
