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

## 2. `liste_profs_papyrus.py` — mettre Papyrus à jour

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
python liste_profs_papyrus.py data/liste_PERSONNEL.txt --excel data/SynchroORCID.xlsx

# 2. Petit essai réel sur 5 profs
python liste_profs_papyrus.py data/liste_PERSONNEL.txt --excel data/SynchroORCID.xlsx --apply --limit 5

# 3. Toute la liste
python liste_profs_papyrus.py data/liste_PERSONNEL.txt --excel data/SynchroORCID.xlsx --apply

# Même chose, mais en laissant le script choisir les fichiers les plus récents de data/
python liste_profs_papyrus.py data --excel data --apply
```

**Sans `--apply`, rien n'est écrit dans Papyrus.** Faites toujours une
simulation avant.

**Donner un dossier au lieu d'un fichier** : le script prend le
`liste_PERSONNEL_*.txt` et le `SynchroORCID_*.xlsx` les plus récents
(d'après la date dans leur nom) et écrit dans le journal lesquels il a
choisis. C'est la façon conseillée pour l'exécution automatique.

### Options

| Option | Par défaut | À quoi ça sert |
|---|---|---|
| `fichier_txt` | *(obligatoire)* | Le fichier liste_personnel, ou un dossier (le plus récent est choisi) |
| `--excel` | *(obligatoire)* | Le fichier Excel des ORCID, ou un dossier (le plus récent est choisi) |
| `--apply` | non (simulation) | Écrit vraiment dans Papyrus |
| `--limit N` | tous | Ne traite que les N premiers profs, pour tester. Les départs ne sont alors **pas** vérifiés |
| `--ignorer-seuils` | non | Désactive les garde-fous (voir plus bas), ex. pour le tout premier chargement |
| `--base-url` | valeur du `.env` | Adresse de l'API Papyrus (finit par `/server/api`) |
| `--user` / `--password` | valeurs du `.env` | Compte administrateur |
| `--community` | valeur du `.env` | Communauté où se trouve la collection Person |
| `--collection` | trouvée toute seule | Identifiant (UUID) de la collection à utiliser |
| `--log-file` | `logs/liste_profs_papyrus_AAAAMMJJ_HHMMSS.log` | Où écrire le journal |
| `--verbose` | non | Affiche tout le détail à l'écran, pas seulement dans le journal |

`python liste_profs_papyrus.py --help` affiche la liste complète.

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

## Les garde-fous

Le script est fait pour tourner seul, par exemple chaque nuit. Pour qu'un
mauvais fichier ou une panne ne fasse pas de dégâts, il s'arrête de
lui-même dans ces cas :

| Situation | Ce que fait le script | Cause probable |
|---|---|---|
| Moins de **1500 profs** dans le fichier | S'arrête avant d'écrire quoi que ce soit | Fichier vide, coupé ou au mauvais format |
| Plus de **100 créations** dans le même run | S'arrête après la 100ᵉ | L'index de recherche de Papyrus est vide ou en reconstruction : les profs existants ne sont pas retrouvés, le script créerait des doublons |
| Plus de **5 %** de la collection à passer `Inactif` d'un coup | Ne marque personne `Inactif` | Fichier incomplet |
| **50 erreurs de suite** | Abandonne le run | Papyrus est en panne |
| Une autre synchronisation `--apply` tourne déjà | Ne démarre pas | Le run précédent n'est pas fini |
| Papyrus ne répond plus | N'attend jamais une réponse plus de 2 minutes (au lieu d'attendre sans fin) | Serveur bloqué |

Les trois premiers peuvent être désactivés avec `--ignorer-seuils`. C'est
nécessaire pour le **tout premier chargement** (plus de 100 fiches à
créer) ou pour un test sur un petit fichier. Les chiffres se changent en
haut de `liste_profs_papyrus.py` (`MIN_PROFESSEURS_ATTENDUS`,
`SEUIL_MAX_CREATIONS`, `SEUIL_MAX_DEPARTS_POURCENT`,
`MAX_ERREURS_CONSECUTIVES`).

Autres protections :

- si la connexion expire en cours de route, le script se reconnecte tout seul ;
- une lecture qui échoue est réessayée deux fois (après 2 s, puis 4 s). Une
  écriture, elle, n'est pas réessayée si on ne sait pas si elle a été faite,
  pour ne pas risquer de doublon : le run suivant s'en chargera ;
- un fichier liste_personnel de plus de **2 jours** ou un fichier ORCID de
  plus de **30 jours** donne un avertissement (export pas reçu ?). Le run se
  fait quand même, mais finit avec le code 2.

## Exécution automatique

Exemple de commande à mettre dans le Planificateur de tâches Windows (ou cron) :

```bash
python C:\chemin\vers\papyrus-liste-prof\liste_profs_papyrus.py C:\chemin\vers\papyrus-liste-prof\data --excel C:\chemin\vers\papyrus-liste-prof\data --apply
```

Avant de l'activer, faites-le tourner quelques jours **sans `--apply`** et
lisez les journaux.

Le code de sortie indique au planificateur si tout s'est bien passé :

| Code | Signification |
|---|---|
| `0` | Tout s'est bien passé |
| `1` | Rien n'a été fait, ou le run a été abandonné (connexion impossible, fichier introuvable, fichier trop petit, Papyrus en panne…) |
| `2` | Le run est allé au bout, mais il y a quelque chose à regarder dans le journal (erreurs sur certains profs, garde-fou déclenché, fichier trop vieux) |
| `3` | Une autre synchronisation était déjà en cours |

## Le journal

Chaque exécution écrit un fichier dans `logs/`
(`liste_profs_papyrus_AAAAMMJJ_HHMMSS.log`) avec **tout le détail** :
fiches créées, modifiées, liens ajoutés ou retirés, avertissements et
erreurs (avec leur cause). À l'écran, on ne voit que l'avancement et le
résumé :

```
2034 professeur(s) | collection 9317ddbc-... | mode APPLICATION RÉELLE

100/2034 (5%) — 2c 15m 83i 0e — ~30 min restantes
...
Résumé : 20 créé(s), 150 mis à jour, 1860 inchangé(s), 4 erreur(s)
  12 prof(s) absent(s) de la liste_personnel marqué(s) Inactif (présents dans DSpace mais plus dans le fichier)
  3 lien(s) vers une ancienne OrgUnit retiré(s) (changement d'unité)
  2 CodeUnite sans OrgUnit correspondant dans DSpace.
```

`c` / `m` / `i` / `e` = créés / mis à jour / inchangés / erreurs. En
simulation, le résumé dit « à marquer » ou « à retirer » au lieu de
« marqué(s) » ou « retiré(s) ».

Ajoutez `--verbose` pour voir tout le détail à l'écran pendant que le
script tourne.

## À savoir

- **Double affiliation** : quand un prof change d'unité, le script retire
  son lien vers l'ancienne unité. Il retirerait aussi un lien vers une
  deuxième unité ajouté à la main. Si ça vous pose problème, mettez
  `RETIRER_ANCIENNES_UNITES = False` en haut de `liste_profs_papyrus.py` :
  le script ne retirera plus jamais de lien.
- **Erreur 500 sur `/core/relationships`** : Papyrus ne renvoie pas le
  détail de l'erreur. Il faut regarder le `dspace.log` sur le serveur, à
  l'heure de l'erreur (attention, ce journal est en heure UTC).

---

**Auteur :** Natalia Jabinschi
