# Synchronisation Profs → Papyrus

Ces scripts prennent la liste des professeurs du fichier liste_personnel,
y ajoutent leur ORCID à partir d'un fichier Excel, puis mettent Papyrus à
jour : ils créent les fiches qui manquent, corrigent celles qui ont changé,
relient chaque prof à son unité et repèrent ceux qui sont partis.

```
fichier liste_personnel (.txt)  ┐
                                 ├─▶ extraire_professeurs.py ─▶ tableau / CSV
fichier ORCID (.xlsx)           ┘
                                             │
                                             ▼
                                 synchro_profs_papyrus.py ─▶ Papyrus (API REST DSpace)
```

## Installation

```bash
pip install requests openpyxl python-dotenv
```

Organisation des fichiers :

```
papyrus-annuaires-donnees/
├── synchro_profs_papyrus.py
├── enrichir_orcid_papyrus.py
├── photos_papyrus.py
├── .env.example   # modèle pour créer votre .env
├── outils/
│   └── extraire_professeurs.py
├── data/          # fichiers reçus (liste_personnel .txt, ORCID .xlsx)
└── logs/          # journaux de chaque exécution (créé tout seul)
```

Le script retrouve lui-même ses dossiers (`outils/`, `logs/`, `.env`) : vous
pouvez le lancer depuis n'importe quel dossier.

### Le fichier `.env` (identifiants)

L'adresse de Papyrus, le compte à utiliser et les fichiers à lire se mettent
dans un fichier `.env`, pour ne pas avoir à les taper dans la commande :

```bash
cp .env.example .env
# puis ouvrez .env et mettez les vraies valeurs
```

```
DSPACE_BASE_URL=http://localhost:8080/server/api
DSPACE_USER=dspace
DSPACE_PASSWORD=...
DSPACE_COMMUNITY=1acd99a0-6ffb-42f8-a261-30b96f3f2405
DSPACE_COLLECTION=
FICHIER_PERSONNEL=data/synchro_PERSONNEL_*.txt
FICHIER_ORCID=data/SynchroORCID_*.xlsx
```

- `DSPACE_COLLECTION` (facultatif) : l'UUID de la collection Person. Vide,
  le script la cherche dans la communauté ; s'il en trouve plusieurs, il les
  affiche et s'arrête — copiez alors l'UUID de la bonne ici.
- `FICHIER_PERSONNEL` / `FICHIER_ORCID` : le chemin d'un fichier, ou un motif
  avec `*` — le script prend alors le plus récent (d'après la date dans le
  nom) et écrit dans le journal lequel il a choisi. Un chemin relatif part du
  dossier du script.
- Ne mettez jamais `.env` dans Git (il est déjà exclu par `.gitignore`).
- Ce que vous donnez dans la commande (fichiers, `--base-url`, `--user`,
  `--password`, `--community`, `--collection`) passe avant le `.env`.
- Avec `--apply`, le script **refuse de démarrer** s'il ne trouve pas
  l'adresse, l'utilisateur et le mot de passe (ni dans `.env`, ni dans la
  commande). Il n'écrit jamais dans Papyrus avec des valeurs par défaut.

---

## 1. `extraire_professeurs.py` — voir la liste sans rien changer

Ce script lit le fichier liste_personnel, garde seulement les professeurs et
affiche un tableau **Nom / Prénom / Courriel / CodeUnite / Statut**. Si vous
donnez le fichier Excel, il ajoute la colonne ORCID. Il ne touche jamais à
Papyrus.

> **Statut**, ici, c'est le titre du poste (ex. « Professeur(e) titulaire »).
> Dans Papyrus, il va dans `person.jobTitle`. Ce n'est pas la même chose que
> `UdeM.statut` (`Actif` / `Inactif`), expliqué plus bas.

```bash
# Le tableau
python outils/extraire_professeurs.py data/liste_PERSONNEL.txt

# Avec les ORCID
python outils/extraire_professeurs.py data/liste_PERSONNEL.txt --excel data/SynchroORCID.xlsx

# En plus, enregistrer le résultat dans un fichier CSV (pratique pour Excel)
python outils/extraire_professeurs.py data/liste_PERSONNEL.txt --excel data/SynchroORCID.xlsx --csv resultat.csv

# Vérifier que le fichier est bien lu (utile si son format change)
python outils/extraire_professeurs.py data/liste_PERSONNEL_20261001.txt --diagnostic
```

`--csv` et `--diagnostic` existent seulement dans ce script, pas dans
`liste_profs_papyrus.py`.

### Le fichier Excel

Il lui faut une colonne dont le titre contient « Courriel » et une autre dont
le titre contient « ORCID », peu importe où elles sont placées.

- Le courriel sert seulement à retrouver le bon prof. On ne prend que l'ORCID ;
  les autres colonnes sont ignorées.
- Une cellule qui ne contient pas un vrai ORCID (ex. « N/A », « à venir »)
  est ignorée, et le script indique combien il en a trouvé.
- Si un même courriel a plusieurs ORCID différents, le script vous prévient
  et garde le premier.

### Si le format du fichier change

Tout se règle en haut du fichier, dans la section `CONFIGURATION` : position
des champs (`CHAMP_NOM_PRENOM`, `ZONE_COURRIEL`), liste des titres qui
comptent comme « professeur » (`TITRES_PROFESSEUR`), noms des colonnes Excel.

---

## 2. `synchro_profs_papyrus.py` — mettre Papyrus à jour

Pour chaque prof de la liste, le script :

- **crée** sa fiche Person si elle n'existe pas encore (on cherche par
  courriel), avec `UdeM.statut = Actif` ;
- **met à jour** seulement ce qui a changé (nom, prénom, poste, ORCID,
  statut). Une case vide dans le fichier n'efface jamais une
  valeur déjà dans Papyrus ;
- pour le **poste** (`person.jobTitle`) : quand le titre change et que le
  nouveau est un titre de rang (Professeur.e adjoint.e, agrégé.e, titulaire,
  etc. — liste `FONCTIONS_DE_RANG`), il **remplace** l'ancien titre de rang,
  par exemple après une promotion. Les titres saisis à la main, hors de cette
  liste, sont **gardés**. Un autre titre de la Synchro (ex. Vice-doyen.ne)
  remplace le premier titre de la fiche, comme avant ;
- **le relie à son unité** (OrgUnit, trouvée par son CodeUnite). Si le prof
  a changé d'unité, le lien vers la nouvelle unité est ajouté, puis celui
  vers l'ancienne est retiré ;
- **repère les départs** : à la fin, toute fiche de Papyrus qui n'est plus
  dans la liste passe à `UdeM.statut = Inactif` et perd son lien vers son unité.

Vous pouvez relancer le script autant de fois que vous voulez : il ne crée
pas de doublons et ne refait pas ce qui est déjà à jour. Si un run est coupé
en plein milieu, il suffit de le relancer.

### Comment le lancer

Avec le `.env` rempli, il n'y a rien d'autre à donner :

```bash
# 1. Simulation : montre ce qui serait fait, sans RIEN écrire
python synchro_profs_papyrus.py

# 2. Petit essai réel sur 5 profs
python synchro_profs_papyrus.py --apply --limit 5

# 3. Toute la liste
python synchro_profs_papyrus.py --apply

# Utiliser d'autres fichiers que ceux du .env, pour une fois
python synchro_profs_papyrus.py data/autre_PERSONNEL.txt --excel data/autre_ORCID.xlsx
```

**Sans `--apply`, rien n'est écrit dans Papyrus.** Faites toujours une
simulation avant.

### Options

| Option | À quoi ça sert |
|---|---|
| `fichier_txt` | Le fichier liste_personnel (remplace `FICHIER_PERSONNEL` du `.env`) |
| `--excel` | Le fichier Excel des ORCID (remplace `FICHIER_ORCID` du `.env`) |
| `--apply` | Écrit vraiment dans Papyrus (sans cette option : simulation) |
| `--limit N` | Ne traite que les N premiers profs, pour tester (les départs ne sont pas vérifiés) |
| `--ignorer-seuils` | Désactive les garde-fous, ex. pour le premier chargement |
| `--verbose` | Affiche tout le détail à l'écran |

Les autres options servent rarement : `--base-url`, `--user`, `--password`,
`--community` et `--collection` remplacent les valeurs du `.env`, et `--log-file` change
l'endroit du journal. `python synchro_profs_papyrus.py --help` les liste toutes.

### Les départs (`UdeM.statut`)

Une fois la liste traitée, le script regarde toutes les fiches Person de la
collection :

- le prof est dans le fichier → `Actif` ;
- le prof n'est plus dans le fichier → `Inactif`, et son lien vers son unité
  (OrgUnit) est retiré. La fiche elle-même n'est jamais supprimée : c'est un
  signal à vérifier (il est peut-être parti, ou il manque par erreur dans le
  fichier). S'il revient dans la liste, il repasse `Actif` et son lien est recréé ;
- déjà `Inactif` → le statut ne change pas, mais un lien vers une unité qui
  resterait est retiré.

Les départs ne sont pas vérifiés quand on utilise `--limit` (la liste est
incomplète, sinon tous les autres profs passeraient `Inactif`), ni quand le
run a été arrêté avant la fin.

---

## 3. `enrichir_orcid_papyrus.py` — mots-clés et biographie ORCID

Ce script complète les fiches Person de Papyrus avec le profil ORCID public
de chaque prof :

| Dans ORCID | Dans Papyrus |
|---|---|
| Mots-clés (*Keywords*) | `dc.subject.discipline`, une valeur par mot-clé |
| Biographie | `UdeM.DonneesORCID` |

Il traite toutes les fiches de la collection qui ont un ORCID
(`person.identifier.orcid`, rempli par `synchro_profs_papyrus.py`). Il ne lit
ni le fichier liste_personnel ni le fichier Excel. Les fiches `Inactif` sont
sautées, sauf avec `--inclure-inactifs`.

- **Champ vide** dans Papyrus : il est rempli avec la valeur ORCID.
- **Champ déjà rempli et différent** : il est remplacé. ORCID est la seule
  source pour ces deux champs : un mot-clé ajouté à la main dans
  `dc.subject.discipline` est remplacé au prochain run.
- **Rien dans ORCID** (pas de mots-clés, pas de biographie, ou profil privé) :
  la fiche n'est pas touchée. Le script n'efface jamais rien.
- Un mot-clé ORCID qui en regroupe plusieurs (« écologie, climat; eau ») est
  découpé sur `,`, `;` et le retour à la ligne. Les doublons sont retirés.
  Des mots-clés séparés seulement par des espaces ne peuvent pas être découpés :
  ils restent en une seule valeur.

```bash
# Simulation (rien n'est écrit) ; --verbose montre les mots-clés trouvés
python enrichir_orcid_papyrus.py --verbose

# Petit essai réel, puis toute la collection
python enrichir_orcid_papyrus.py --apply --limit 5
python enrichir_orcid_papyrus.py --apply
```

Il utilise le même `.env` que la synchronisation, et les mêmes options
`--apply`, `--limit`, `--verbose`, `--log-file`, `--base-url`, `--user`,
`--password`, `--community` et `--collection`. Il partage aussi son verrou :
les deux scripts ne peuvent pas écrire en même temps. Comptez quelques minutes
pour toute la collection (un appel ORCID par fiche). Une fois par semaine
suffit : les profils ORCID changent peu.

À la fin, le résumé liste les ORCID introuvables, désactivés ou fusionnés
dans un autre profil : l'ORCID de ces fiches est à vérifier.

Facultatif : avec des identifiants « Public API » ORCID dans le `.env`
(`ORCID_CLIENT_ID`, `ORCID_CLIENT_SECRET`), le script s'authentifie et a des
limites d'appel plus généreuses. Sans eux, l'accès anonyme suffit.

---

## 4. `photos_papyrus.py` — photos des fiches

Ce script met la photo officielle de chaque chercheur sur sa fiche Person,
à partir de l'archive `data/chercheurs.zip` (pas besoin de la décompresser) :

```
chercheurs/<codesim>/info.txt                   nom, prenom, courriel, codesim, idsadvr
chercheurs/<codesim>/photoOfficielle.jpg        la photo envoyée dans Papyrus
chercheurs/<codesim>/photoOfficiellePetite.jpg  ignorée (trop petite)
```

- Le dossier est rattaché à la fiche **par le courriel** de `info.txt`. Sans
  courriel, ou si aucune fiche n'a ce courriel, le dossier est sauté. Le
  résumé liste les dossiers sans courriel et ceux dont **le courriel diffère
  de celui de la fiche au même nom** (photo non mise, courriel à corriger) ;
  pour les chercheurs qui n'ont pas de fiche dans Papyrus, il ne donne que le total.
- Un dossier **sans courriel** est détaillé dans le résumé seulement si une
  fiche Papyrus porte le même nom (son courriel est affiché, pour compléter
  `info.txt`). Sinon, le résumé ne donne que le total.
- Avec `--csv-logs`, les dossiers **sans fiche dans Papyrus** sont aussi écrits
  dans `logs/photos_sans_fiche_<date>.csv` (codesim, nom, prénom, courriel,
  idsadvr), à ouvrir dans Excel : `python photos_papyrus.py --apply --csv-logs`.
- **Pas de photo** sur la fiche : elle est ajoutée. **Même photo** : rien n'est
  envoyé. **Autre image** (ancienne photo, ou image ajoutée à la main) : elle
  est remplacée par la photo officielle.
- La photo va dans le bundle `ORIGINAL`, avec une copie dans `THUMBNAIL` : elle
  s'affiche tout de suite, sans attendre la tâche DSpace `filter-media`. La
  nouvelle photo est envoyée avant de retirer l'ancienne.
- Les fiches `Inactif` sont sautées, sauf avec `--inclure-inactifs`. Un nom ou
  prénom différent de `info.txt` est signalé, sans bloquer.

```bash
python photos_papyrus.py                       # simulation
python photos_papyrus.py --apply --limit 1     # essai réel sur 1 fiche
python photos_papyrus.py --apply
```

L'archive se règle dans le `.env` (`FICHIER_PHOTOS`, motif `*` accepté, la plus
récente est prise) ou se donne dans la commande. Mêmes options, même `.env` et
même verrou que les autres scripts. Le premier run, qui ajoute presque toutes
les photos, prend 20 à 30 minutes ; les suivants, quelques minutes.

---

## Bon à savoir

Le script peut tourner seul (par exemple chaque nuit). Pour éviter les dégâts, il s'arrête de lui-même si :

| Situation | Ce que fait le script |
|---|---|
| Moins de **1500 profs** dans le fichier (fichier vide ou coupé) | S'arrête sans rien écrire |
| Plus de **100 créations** dans le même run (risque de doublons) | S'arrête après la 100ᵉ |
| Plus de **5 %** de la collection à passer `Inactif` d'un coup | Ne marque personne `Inactif` |
| **50 erreurs de suite** (Papyrus en panne) | Abandonne le run |

- Pour le **premier chargement** ou un test sur un petit fichier, ajoutez `--ignorer-seuils`.
- Un fichier liste_personnel de plus de **2 jours** ou ORCID de plus de **30 jours** donne un avertissement.
- **Journal** : tout le détail est dans `logs/` ; `--verbose` l'affiche aussi à l'écran.
- **Codes de sortie** : `0` OK · `1` run abandonné · `2` terminé, à vérifier dans le journal · `3` une autre synchronisation tournait déjà.
- **Liens retirés** : quand un prof change d'unité, le lien vers l'ancienne est retiré ; quand il passe `Inactif`, tous ses liens vers une unité sont retirés — même ceux ajoutés à la main. Pour l'éviter : `RETIRER_ANCIENNES_UNITES = False` ou `RETIRER_UNITE_SI_INACTIF = False`.

---

**Auteur :** Natalia Jabinschi
