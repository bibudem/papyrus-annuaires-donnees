# Synchronisation Profs → DSpace

Pipeline en deux étapes pour extraire les professeurs du fichier liste_personnel
(export RH à largeur fixe), croiser leur ORCID via un fichier Excel, puis
les synchroniser dans un dépôt DSpace 7 (création/mise à jour, liaison à
leur OrgUnit, et suivi des départs).

```
fichier liste_personnel (.txt)  ┐
                                 ├─▶ extraire_professeurs.py ─▶ tableau / CSV
fichier ORCID (.xlsx)           ┘
                                             │
                                             ▼
                                 synchro_profs_dspace.py ─▶ DSpace (API REST)
```

## Prérequis

```bash
pip install requests openpyxl python-dotenv
```

(`python-dotenv` est optionnel — seulement nécessaire pour le chargement
automatique du fichier `.env`, voir ci-dessous.)

Les deux scripts doivent être dans le **même dossier** (`synchro_profs_dspace.py`
importe des fonctions de `extraire_professeurs.py`).

### Configuration sensible (`.env`)

Les identifiants et l'URL DSpace peuvent être mis dans un fichier `.env`
plutôt que tapés en ligne de commande (évite qu'ils traînent dans
l'historique shell ou la liste des processus) :

```bash
cp .env.example .env
# puis éditer .env avec les vraies valeurs
```

```
DSPACE_BASE_URL=http://localhost:8080/server/api
DSPACE_USER=dspace
DSPACE_PASSWORD=...
DSPACE_COMMUNITY=1acd99a0-6ffb-42f8-a261-30b96f3f2405
```

`.env` est dans `.gitignore` — ne jamais le committer. Les options en ligne
de commande (`--base-url`, `--user`, `--password`, `--community`) restent
toujours prioritaires si fournies, `.env` ne sert que de valeur par défaut.

---

## 1. `extraire_professeurs.py`

Lit le fichier liste_personnel (texte à largeur fixe, sans séparateur), garde
uniquement les lignes dont le titre correspond à un poste de professeur, et
affiche un tableau **Nom / Prénom / Courriel / CodeUnite / Statut**. Ajoute
l'ORCID si un fichier Excel est fourni.

### Usage

```bash
# Tableau simple
python extraire_professeurs.py liste_personnel.txt

# Avec ORCID (croisé par courriel)
python extraire_professeurs.py liste_personnel.txt --excel annuaire.xlsx

# Écrit aussi un fichier CSV (utile pour Excel, évite les problèmes d'affichage)
python extraire_professeurs.py liste_personnel.txt --excel annuaire.xlsx --csv resultat.csv

# Mode diagnostic : affiche les champs extraits + une règle de positions,
# pour valider/ajuster le format si jamais la structure du fichier change
python extraire_professeurs.py liste_personnel.txt --diagnostic
```

### Fichier Excel attendu

Une colonne dont l'en-tête contient « Courriel » et une autre contenant
« ORCID » (trouvées par mot-clé, peu importe leur position/lettre de
colonne). Si un courriel a plusieurs ORCID différents dans le fichier
(erreur de saisie), le script avertit et garde le premier trouvé.

**Seul l'ORCID est effectivement récupéré.** La colonne Courriel ne sert
que de clé de correspondance avec le fichier liste_personnel (le courriel
affiché dans le tableau final vient du fichier liste_personnel, pas de
l'Excel). Les autres colonnes du fichier Excel (Fonction, DescUnite, etc.)
sont ignorées, même si elles sont présentes.

### Ce qui est configurable

Tout en haut du fichier, section `CONFIGURATION` : positions des champs
(`CHAMP_NOM_PRENOM`, `ZONE_COURRIEL`), liste des titres reconnus comme
« professeur » (`TITRES_PROFESSEUR`), mots-clés des colonnes Excel, etc.
C'est le seul endroit à modifier si le format du fichier liste_personnel change.

---

## 2. `synchro_profs_dspace.py`

Prend la liste de professeurs (même logique d'extraction que ci-dessus) et
la synchronise dans DSpace 7 via l'API REST :

- **Crée** un item Person s'il n'existe pas déjà (recherché par courriel),
  avec `UdeM.statut = Actif`.
- **Met à jour** les champs qui ont changé sur un item existant (rien n'est
  réécrit inutilement).
- **Lie** chaque prof à son OrgUnit (trouvé via son CodeUnite) par une vraie
  relation DSpace (`/core/relationships`) — pas une métadonnée bricolée.
- **Marque les départs** : tout item Person déjà dans DSpace mais absent du
  fichier traité est marqué `UdeM.statut = Inactif` — jamais retiré ni
  supprimé, juste signalé pour révision (voir plus bas).
- **Idempotent** : relancer le script ne crée jamais de doublons ni ne
  re-marque inutilement ce qui est déjà à jour, donc une interruption en
  cours de route n'est jamais dangereuse.

### Usage

```bash
# 1. Mode simulation (RIEN n'est écrit) : montre ce qui serait fait
python synchro_profs_dspace.py liste_personnel.txt --excel annuaire.xlsx

# 2. Petit test réel sur 5 profs seulement
python synchro_profs_dspace.py liste_personnel.txt --excel annuaire.xlsx --apply --limit 5

# 3. Une fois validé, sur la liste complète
python synchro_profs_dspace.py liste_personnel.txt --excel annuaire.xlsx --apply
```

⚠️ **Par défaut, rien n'est écrit dans DSpace** tant que `--apply` n'est pas
précisé. Toujours tester avec `--limit` avant un run complet.

### Options principales

| Option | Défaut | Description |
|---|---|---|
| `--excel` | *(requis)* | Fichier Excel des ORCID |
| `--apply` | désactivé | Applique réellement les changements |
| `--limit N` | tous | Ne traite que les N premiers profs (pour tester) |
| `--base-url` | `http://localhost:8080/server/api` | URL de l'API REST DSpace |
| `--user` / `--password` | `dspace` / `dspace` | Identifiants admin |
| `--community` | UUID configuré | Communauté où chercher la collection Person |
| `--collection` | *(auto)* | UUID de la collection cible (court-circuite `--community`) |
| `--log-file` | horodaté | Chemin du fichier `.log` |
| `--verbose` | désactivé | Affiche le détail ligne par ligne en console |

### Résolution de la collection

Le script cherche automatiquement la collection dans la communauté donnée.
S'il y en a **plusieurs**, il les liste et s'arrête — il faut alors relancer
avec `--collection <uuid>` pour confirmer explicitement laquelle utiliser
(volontaire, pour éviter un dépôt dans la mauvaise collection).

### Suivi des départs (`UdeM.statut`)

Après avoir traité tous les profs du fichier, le script parcourt **tous**
les items Person de la collection cible (via l'API Discovery, `scope` +
pagination) et compare leur courriel à la liste actuelle :

- **Présent dans le fichier** → `UdeM.statut = Actif` (déjà fait pendant le
  traitement normal, création ou mise à jour).
- **Absent du fichier, mais présent dans DSpace** → `UdeM.statut = Inactif`.
  Rien n'est retiré ni supprimé — c'est un signal à réviser manuellement
  (le prof peut être parti, ou simplement absent de ce run par erreur).
- **Déjà `Inactif`** → jamais retouché inutilement.

Cette étape fonctionne aussi en mode simulation (affiche combien **seraient**
marqués, sans rien écrire).

⚠️ Le script marque volontairement en deux passes séparées (jamais tout le
monde à `Inactif` d'un coup avant de retraiter la liste) : si le run est
interrompu en cours de route, aucun prof actif ne se retrouve dans une
fenêtre où il semble à tort marqué `Inactif`.

### Journalisation

Chaque run écrit un fichier `synchro_profs_dspace_AAAAMMJJ_HHMMSS.log` (ou
le chemin donné via `--log-file`) qui garde **tout le détail** (créations,
mises à jour, liaisons, avertissements) — utile pour auditer après coup un
run de plusieurs milliers d'entrées. La console, elle, reste compacte par
défaut :

```
3000 professeur(s) | collection 9317ddbc-... | mode APPLICATION RÉELLE

100/3000 (3%) — 80c 15m 5i 0e — ~45 min restantes
...
Résumé : 2850 créé(s), 120 mis à jour, 25 inchangé(s), 5 erreur(s)
  12 prof(s) absent(s) de la liste_personnel marqué(s) Inactif (présents dans DSpace mais plus dans le fichier)
  3 liaison(s) OrgUnit échouée(s) (probablement déjà liées)
```

(`c`/`m`/`i`/`e` = créés / mis à jour / inchangés / erreurs)

Ajouter `--verbose` pour retrouver le détail ligne par ligne en direct dans
la console (par défaut, seul le fichier `.log` le garde).

### Robustesse (gros volumes)

- **Reconnexion automatique** si le jeton d'authentification expire en
  cours de route (401 → reconnexion → nouvel essai).
- **Nouvelle tentative avec délai croissant** (2s, 4s, 8s) sur erreur réseau
  ou erreur serveur transitoire (502/503/504).
- Un échec de **liaison OrgUnit** n'interrompt jamais la création/mise à
  jour de la fiche Person — seule une vraie erreur sur la fiche elle-même
  compte comme erreur.

### Ce qui est configurable

Tout en haut du fichier, section `CONFIGURATION` :

- Identifiants/URL par défaut
- Noms des champs de métadonnées DSpace (`CHAMP_NOM`, `CHAMP_ORCID`,
  `CHAMP_CODEUNITE`, `CHAMP_STATUT`, etc.) — à ajuster si ton schéma de
  métadonnées diffère
- `VALEUR_STATUT_ACTIF` / `VALEUR_STATUT_INACTIF` — les deux valeurs
  utilisées pour `UdeM.statut`
- `NOM_RELATION_ORGUNIT` — nom du relationshipType DSpace pour le lien
  Person↔OrgUnit (résolu dynamiquement via `/core/relationshiptypes`, pas
  besoin de connaître l'ID ni l'ordre gauche/droite à l'avance)
- Paramètres réseau (`NB_TENTATIVES_MAX`, `DELAI_BACKOFF_SECONDES`)

**Note** : le script ne définit plus `dc.title` (titre affiché) sur les
fiches créées — ce champ n'est pas nécessaire au bon fonctionnement du lien
Person↔OrgUnit, qui repose uniquement sur l'UUID.

---

## Dépannage

**`ModuleNotFoundError: No module named 'requests'`**
→ `pip install requests openpyxl`

**`UnicodeDecodeError` à la lecture du fichier liste_personnel**
→ Le script essaie déjà UTF-8, cp1252 et latin-1 dans l'ordre ; si aucun ne
fonctionne, le fichier a un encodage encore différent — le préciser.

**`Échec de connexion DSpace (405) : Only POST is allowed for login requests`**
→ Généralement causé par une redirection HTTP suivie automatiquement (et
convertie en GET). Déjà géré dans le script (`allow_redirects=False` +
gestion manuelle), mais si l'erreur revient, vérifier que `--base-url`
pointe bien vers `.../server/api` exactement.

**`Plusieurs collections trouvées`**
→ Normal si la communauté contient plus d'une collection ; relancer avec
`--collection <uuid>` (voir la liste affichée).

**Erreur 500 générique sur `/core/relationships`**
→ La réponse REST ne contient jamais le vrai détail d'un 500 ; regarder le
`dspace.log` côté serveur (⚠️ en UTC, pas l'heure locale) autour de l'horaire
de l'erreur. Voir aussi la piste « liaison déjà existante » dans le résumé
du script.

**`Impossible de vérifier les départs : 404 ... /core/collections/{uuid}/items`**
→ Cet endpoint n'existe pas dans l'API REST DSpace 7 ; le script liste les
items d'une collection via l'API Discovery (`scope=<uuid>`) à la place. Si
l'erreur revient malgré tout, vérifier que la collection existe bien et que
le compte utilisé a les droits de lecture dessus.
