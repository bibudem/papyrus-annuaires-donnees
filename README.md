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
papyrus-liste-prof/
├── synchro_profs_papyrus.py
├── .env.example   # modèle pour créer votre .env
├── outils/
│   └── extraire_professeurs.py
├── data/          # fichiers reçus (liste_personnel .txt, ORCID .xlsx)
└── logs/          # journaux de chaque exécution (créé tout seul)
```

Le script retrouve lui-même ses dossiers (`outils/`, `logs/`, `.env`) : vous
pouvez le lancer depuis n'importe quel dossier.

### Le fichier `.env` (identifiants)

L'adresse de Papyrus et le compte à utiliser se mettent dans un fichier
`.env`, pour ne pas avoir à taper le mot de passe dans la commande :

```bash
cp .env.example .env
# puis ouvrez .env et mettez les vraies valeurs
```

```
DSPACE_BASE_URL=http://localhost:8080/server/api
DSPACE_USER=dspace
DSPACE_PASSWORD=...
DSPACE_COMMUNITY=1acd99a0-6ffb-42f8-a261-30b96f3f2405
```

- Ne mettez jamais `.env` dans Git (il est déjà exclu par `.gitignore`).
- Si vous donnez `--base-url`, `--user`, `--password` ou `--community` dans
  la commande, ces valeurs passent avant celles du `.env`.
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
  CodeUnite, statut). Une case vide dans le fichier n'efface jamais une
  valeur déjà dans Papyrus ;
- **le relie à son unité** (OrgUnit, trouvée par son CodeUnite). Si le prof
  a changé d'unité, le lien vers la nouvelle unité est ajouté, puis celui
  vers l'ancienne est retiré ;
- **repère les départs** : à la fin, toute fiche de Papyrus qui n'est plus
  dans la liste passe à `UdeM.statut = Inactif`.

Vous pouvez relancer le script autant de fois que vous voulez : il ne crée
pas de doublons et ne refait pas ce qui est déjà à jour. Si un run est coupé
en plein milieu, il suffit de le relancer.

### Comment le lancer

```bash
# 1. Simulation : montre ce qui serait fait, sans RIEN écrire
python synchro_profs_papyrus.py data/liste_PERSONNEL.txt --excel data/SynchroORCID.xlsx

# 2. Petit essai réel sur 5 profs
python synchro_profs_papyrus.py data/liste_PERSONNEL.txt --excel data/SynchroORCID.xlsx --apply --limit 5

# 3. Toute la liste
python synchro_profs_papyrus.py data/liste_PERSONNEL.txt --excel data/SynchroORCID.xlsx --apply

# Même chose, mais en laissant le script choisir les fichiers les plus récents de data/
python synchro_profs_papyrus.py data --excel data --apply
```

**Sans `--apply`, rien n'est écrit dans Papyrus.** Faites toujours une
simulation avant.

**Donner un dossier au lieu d'un fichier** : le script prend le
`liste_PERSONNEL_*.txt` et le `SynchroORCID_*.xlsx` les plus récents
(d'après la date dans leur nom) et écrit dans le journal lesquels il a
choisis. C'est la façon conseillée pour l'exécution automatique.

### Options

| Option | À quoi ça sert |
|---|---|
| `fichier_txt` | Le fichier liste_personnel, ou le dossier `data` (le plus récent est choisi) |
| `--excel` | Le fichier Excel des ORCID, ou le dossier `data` (le plus récent est choisi) |
| `--apply` | Écrit vraiment dans Papyrus (sans cette option : simulation) |
| `--limit N` | Ne traite que les N premiers profs, pour tester (les départs ne sont pas vérifiés) |
| `--ignorer-seuils` | Désactive les garde-fous, ex. pour le premier chargement |
| `--collection` | À donner seulement si le script trouve plusieurs collections |
| `--verbose` | Affiche tout le détail à l'écran |

Les autres options servent rarement : `--base-url`, `--user`, `--password`
et `--community` remplacent les valeurs du `.env`, et `--log-file` change
l'endroit du journal. `python synchro_profs_papyrus.py --help` les liste toutes.

### La collection

Le script cherche tout seul la collection dans la communauté. S'il en trouve
**plusieurs**, il les affiche et s'arrête : relancez avec
`--collection <uuid>` pour dire laquelle utiliser. C'est voulu, pour ne
jamais écrire dans la mauvaise collection.

### Les départs (`UdeM.statut`)

Une fois la liste traitée, le script regarde toutes les fiches Person de la
collection :

- le prof est dans le fichier → `Actif` ;
- le prof n'est plus dans le fichier → `Inactif`. Sa fiche n'est ni retirée
  ni supprimée : c'est juste un signal à vérifier (il est peut-être parti,
  ou il manque par erreur dans le fichier) ;
- déjà `Inactif` → on n'y touche pas.

Les départs ne sont pas vérifiés quand on utilise `--limit` (la liste est
incomplète, sinon tous les autres profs passeraient `Inactif`), ni quand le
run a été arrêté avant la fin.

---

## Sécurités, journal et points à connaître

Le script peut tourner seul (par exemple chaque nuit). Il s'arrête de lui-même dans ces cas :

| Situation | Ce que fait le script |
|---|---|
| Moins de **1500 profs** dans le fichier (fichier vide ou coupé) | S'arrête sans rien écrire |
| Plus de **100 créations** dans le même run (index de recherche vide, risque de doublons) | S'arrête après la 100ᵉ |
| Plus de **5 %** de la collection à passer `Inactif` d'un coup (fichier incomplet) | Ne marque personne `Inactif` |
| **50 erreurs de suite** (Papyrus en panne) | Abandonne le run |
| Une autre synchronisation `--apply` tourne déjà | Ne démarre pas |
| Papyrus ne répond plus | N'attend jamais plus de 2 minutes |

- `--ignorer-seuils` désactive les trois premières règles. C'est nécessaire pour le **premier chargement** ou pour un test sur un petit fichier. Les chiffres se changent en haut de `synchro_profs_papyrus.py`.
- Si la connexion expire, le script se reconnecte tout seul. Une lecture qui échoue est réessayée deux fois ; une écriture douteuse ne l'est pas (pas de doublon), c'est le run suivant qui la refera.
- Un fichier liste_personnel de plus de **2 jours** ou un fichier ORCID de plus de **30 jours** donne un avertissement.
- **Journal** : chaque exécution écrit tout le détail dans `logs/` (fiches créées ou modifiées, liens, erreurs). À l'écran, on ne voit que l'avancement et le résumé ; ajoutez `--verbose` pour tout voir. Dans l'avancement, `c` / `m` / `i` / `e` = créés / mis à jour / inchangés / erreurs.
- **Codes de sortie** : `0` tout va bien · `1` rien n'a été fait ou run abandonné · `2` terminé, mais à vérifier dans le journal · `3` une autre synchronisation tournait déjà.
- **Double affiliation** : quand un prof change d'unité, le lien vers l'ancienne est retiré, y compris un lien ajouté à la main. Pour l'éviter, mettez `RETIRER_ANCIENNES_UNITES = False`.
- **Erreur 500 sur `/core/relationships`** : regardez le `dspace.log` sur le serveur, à l'heure de l'erreur (attention, il est en heure UTC).

---

**Auteur :** Natalia Jabinschi
