#!/usr/bin/env python3
"""
Synchronise les professeurs (extraits par extraire_professeurs.py) avec un
dépôt DSpace 7 via son API REST : crée les items "Person" manquants, met à
jour les existants (ORCID, CodeUnite, etc.), et les lie à leur OrgUnit par
une vraie relation DSpace — le tout par correspondance de courriel.

Usage:
    python synchro_profs_papyrus.py fichier.txt --excel annuaire.xlsx
    python synchro_profs_papyrus.py fichier.txt --excel annuaire.xlsx --apply

Par défaut le script tourne en mode SIMULATION (affiche ce qu'il ferait sans
rien écrire dans DSpace). Ajoute --apply pour appliquer réellement les
créations/mises à jour.

Conçu pour tourner sur de gros volumes (plusieurs milliers d'entrées) :
reconnexion automatique si le jeton expire en cours de route, nouvelle
tentative avec délai croissant sur les erreurs réseau/serveur transitoires,
journal détaillé dans un fichier .log, et affichage périodique de la
progression. Le script est idempotent : le relancer ne crée jamais de
doublons, donc une interruption n'est jamais dangereuse, juste à refaire.

Dépendances : requests, openpyxl (pip install requests openpyxl si absents).
python-dotenv optionnel (pip install python-dotenv) pour charger un fichier
.env automatiquement — voir .env.example pour les variables reconnues.
Doit être exécuté depuis la racine du projet (extraire_professeurs.py est
importé depuis outils/, voir DOSSIER_OUTILS ci-dessous).
"""

import argparse
import logging
import glob
import os
import re
import sys
import time
from datetime import datetime

import requests

# Dossier du script : sert de base à tous les chemins internes (.env, outils/,
# logs/, verrou) pour que le script marche quel que soit le dossier courant
# (ex. Planificateur de tâches Windows, qui démarre souvent dans System32).
DOSSIER_SCRIPT = os.path.dirname(os.path.abspath(__file__))

try:
    from dotenv import load_dotenv
    if not load_dotenv(os.path.join(DOSSIER_SCRIPT, ".env")):
        print("⚠ Aucun fichier .env trouvé (ou vide) — utilisation des valeurs par défaut "
              "codées en dur (voir .env.example pour en créer un).", file=sys.stderr)
except ImportError:
    print("⚠ python-dotenv n'est pas installé dans CET environnement Python — le fichier "
          ".env ne sera PAS chargé, les valeurs par défaut codées en dur seront utilisées. "
          "Installe-le avec : pip install python-dotenv", file=sys.stderr)

# extraire_professeurs.py vit dans outils/ (pas un package, juste un dossier
# de scripts) : on l'ajoute au chemin de recherche des modules pour pouvoir
# l'importer normalement, peu importe le dossier depuis lequel ce script est
# lancé.
DOSSIER_OUTILS = os.path.join(DOSSIER_SCRIPT, "outils")
sys.path.insert(0, DOSSIER_OUTILS)

from extraire_professeurs import (
    charger_orcid_depuis_excel,
    extraire_infos,
    filtrer_professeurs,
    lire_lignes,
    normaliser_courriel,
)

# =============================================================================
# CONFIGURATION
# Tout ce qui est propre à l'instance DSpace, à son modèle de données, ou au
# comportement du script sur de gros volumes est regroupé ici.
# =============================================================================

# --- Connexion à l'instance DSpace (surchargeables par --base-url/--user/--password) ---
# Valeurs par défaut lues depuis les variables d'environnement (voir .env.example) ;
# la ligne de commande (--base-url/--user/--password/--community) a toujours priorité.
BASE_URL_PAR_DEFAUT = os.environ.get("DSPACE_BASE_URL", "http://localhost:8080/server/api")
UTILISATEUR_PAR_DEFAUT = os.environ.get("DSPACE_USER", "dspace")
MOT_DE_PASSE_PAR_DEFAUT = os.environ.get("DSPACE_PASSWORD", "dspace")
# Paramètres de connexion qui doivent venir du .env (ou de la ligne de commande)
# en mode --apply : on refuse d'écrire avec les valeurs de secours codées en dur
# ci-dessus (signe que le .env n'a pas été chargé).
CONNEXION_OBLIGATOIRE = {"base_url": "DSPACE_BASE_URL", "user": "DSPACE_USER", "password": "DSPACE_PASSWORD"}

# --- Emplacement de dépôt (surchargeable par --community/--collection) ---
# UUID de la communauté où chercher automatiquement la collection Person.
COMMUNAUTE_PAR_DEFAUT = os.environ.get("DSPACE_COMMUNITY", "1acd99a0-6ffb-42f8-a261-30b96f3f2405")
# UUID de la collection Person à utiliser directement (facultatif) : si rempli,
# la communauté n'est plus consultée. Utile quand elle contient plusieurs collections.
COLLECTION_PAR_DEFAUT = os.environ.get("DSPACE_COLLECTION") or None

# --- Fichiers d'entrée ---
# Chaque fichier peut être donné comme un chemin de fichier, un dossier ou un
# motif avec * (ex. data/synchro_PERSONNEL_*.txt). Pour un dossier ou un motif,
# le fichier le plus récent est choisi (d'après la date AAAAMMJJ dans son nom,
# sinon sa date de modification). Pour un dossier, on y cherche le motif ci-dessous.
MOTIF_FICHIER_PERSONNEL = "synchro_PERSONNEL_*.txt"
MOTIF_FICHIER_ORCID = "SynchroORCID_*.xlsx"
# Fichiers utilisés quand la commande n'en donne pas : FICHIER_PERSONNEL et
# FICHIER_ORCID du .env, sinon les motifs ci-dessus dans data/. Un chemin
# relatif part du dossier du script (pas du dossier courant).
FICHIER_PERSONNEL_PAR_DEFAUT = os.environ.get("FICHIER_PERSONNEL") or os.path.join("data", MOTIF_FICHIER_PERSONNEL)
FICHIER_ORCID_PAR_DEFAUT = os.environ.get("FICHIER_ORCID") or os.path.join("data", MOTIF_FICHIER_ORCID)
# Au-delà de cet âge (en jours), le fichier est signalé comme périmé (export
# non reçu ?) : le run se fait quand même, mais finit avec le code de sortie 2.
AGE_MAX_FICHIER_PERSONNEL_JOURS = 2
AGE_MAX_FICHIER_ORCID_JOURS = 30

# --- Noms des champs de métadonnées DSpace pour un item "Person" ---
CHAMP_TYPE_ENTITE = "dspace.entity.type"
VALEUR_TYPE_ENTITE = "Person"
CHAMP_NOM = "person.familyName"
CHAMP_PRENOM = "person.givenName"
CHAMP_COURRIEL = "person.email"
CHAMP_FONCTION = "person.jobTitle"
CHAMP_ORCID = "person.identifier.orcid"
CHAMP_CODEUNITE = "organization.identifier.UdeM"

# --- Statut Actif/Inactif (présence dans la liste_personnel actuelle) ---
# "Actif" pour tout prof présent dans le fichier traité ; "Inactif" pour un
# item Person déjà dans DSpace mais absent du fichier (départ probable) —
# jamais retiré/supprimé automatiquement, juste marqué pour révision.
CHAMP_STATUT = "UdeM.statut"
VALEUR_STATUT_ACTIF = "Actif"
VALEUR_STATUT_INACTIF = "Inactif"

# --- Relation Person <-> OrgUnit ---
# Nom du relationshipType DSpace (tel qu'il apparaît en leftwardType ou
# rightwardType dans /core/relationshiptypes) décrivant le lien Person-OrgUnit.
# Créée via l'API dédiée /core/relationships (PAS une métadonnée) : c'est
# cette API qui fait apparaître le champ "relation.isOrgUnitOfPerson" —
# généré automatiquement et affiché en lecture seule sur la fiche Person.
NOM_RELATION_ORGUNIT = "isOrgUnitOfPerson"
VALEUR_TYPE_ENTITE_ORGUNIT = "OrgUnit"
# Prof qui change d'unité : retirer le lien vers l'ancienne OrgUnit (après
# création du nouveau). ATTENTION : retire aussi tout lien de ce type ajouté
# à la main vers une autre unité (ex. double affiliation) — mettre False pour
# ne jamais retirer de lien.
RETIRER_ANCIENNES_UNITES = True

# --- Authentification DSpace (protocole JWT + double cookie CSRF) ---
COOKIE_CSRF = "DSPACE-XSRF-COOKIE"
EN_TETE_CSRF = "X-XSRF-TOKEN"
EN_TETE_AUTORISATION = "Authorization"

# --- Recherche d'un item existant ---
TAILLE_PAGE_RECHERCHE = 10  # nb de résultats examinés pour trouver une correspondance de courriel

# --- Codes HTTP considérés comme un succès pour une création/mise à jour ---
CODES_HTTP_SUCCES = (200, 201)

# --- Robustesse réseau (utile sur de gros volumes qui tournent longtemps) ---
NB_TENTATIVES_MAX = 3            # nb total d'essais avant d'abandonner un appel
DELAI_BACKOFF_SECONDES = 2       # délai de base, doublé à chaque nouvel essai (2s, 4s, 8s...)
CODES_HTTP_A_REESSAYER = {502, 503, 504}  # erreurs serveur considérées comme transitoires
# Délai max (connexion, lecture) en secondes pour chaque appel HTTP : sans ça,
# un serveur qui ne répond plus bloquerait le script indéfiniment.
TIMEOUT_HTTP = (10, 120)

# --- Garde-fous (exécution automatique sans surveillance) ---
# Protègent contre un fichier d'entrée vide/tronqué ou un index de recherche
# DSpace vide (ex. réindexation en cours). Désactivables avec --ignorer-seuils
# (ex. premier chargement complet, ou tests sur un petit fichier).
MIN_PROFESSEURS_ATTENDUS = 1500    # moins de profs dans le fichier -> arrêt avant toute écriture
SEUIL_MAX_CREATIONS = 100          # créations max par run -> au-delà, arrêt (doublons probables)
SEUIL_MAX_DEPARTS_POURCENT = 5     # % max de la collection marquable Inactif en un run
# Erreurs d'affilée avant d'abandonner le run (DSpace probablement en panne) :
# évite de passer des heures à échouer prof par prof. Indépendant de --ignorer-seuils.
MAX_ERREURS_CONSECUTIVES = 50

# --- Verrou : empêche deux synchronisations (--apply) de tourner en même temps ---
FICHIER_VERROU = os.path.join(DOSSIER_SCRIPT, ".synchro_profs_papyrus.lock")

# --- Codes de sortie (pour le planificateur de tâches / la supervision) ---
CODE_SORTIE_OK = 0
CODE_SORTIE_ERREUR_FATALE = 1       # rien ou presque n'a été fait (connexion, fichier, garde-fou...)
CODE_SORTIE_ERREURS_PARTIELLES = 2  # run terminé, mais avec des erreurs à examiner dans le .log
CODE_SORTIE_DEJA_EN_COURS = 3       # une autre synchronisation tient déjà le verrou

# --- Journalisation et suivi de progression ---
INTERVALLE_PROGRES = 100  # affiche un point d'étape tous les N professeurs traités
DOSSIER_LOGS = os.path.join(DOSSIER_SCRIPT, "logs")  # dossier des .log par défaut (créé si absent)


# =============================================================================
# JOURNALISATION
# Toutes les sorties du script passent par ce logger : elles apparaissent à
# la fois dans la console et dans un fichier .log horodaté, utile pour
# auditer après coup un run de plusieurs milliers d'entrées.
# =============================================================================

LOG = logging.getLogger("synchro_profs_papyrus")


def configurer_journalisation(chemin_log, verbose=False):
    """Le fichier .log garde TOUJOURS le détail complet (utile pour auditer
    après coup un run de plusieurs milliers d'entrées). La console, elle,
    n'affiche par défaut que les points d'étape et le résumé — le détail
    ligne par ligne (créations, mises à jour, liaisons) n'apparaît à l'écran
    qu'avec --verbose."""
    LOG.setLevel(logging.DEBUG)

    dossier = os.path.dirname(chemin_log)
    if dossier:
        os.makedirs(dossier, exist_ok=True)

    gestionnaire_fichier = logging.FileHandler(chemin_log, encoding="utf-8")
    gestionnaire_fichier.setLevel(logging.DEBUG)
    gestionnaire_fichier.setFormatter(logging.Formatter("%(asctime)s  %(message)s", "%Y-%m-%d %H:%M:%S"))
    LOG.addHandler(gestionnaire_fichier)

    gestionnaire_console = logging.StreamHandler()
    gestionnaire_console.setLevel(logging.DEBUG if verbose else logging.INFO)
    gestionnaire_console.setFormatter(logging.Formatter("%(message)s"))
    LOG.addHandler(gestionnaire_console)

    LOG.debug(f"Journal détaillé écrit dans : {chemin_log}")


# =============================================================================
# CLIENT REST DSPACE
# Encapsule l'authentification et les appels API. Aucune logique métier
# (quels champs comparer, etc.) ne devrait se retrouver ici.
# =============================================================================

class DSpaceClient:
    """Client REST pour DSpace 7 : authentification (avec reconnexion
    automatique si le jeton expire), recherche d'un item Person par courriel,
    création, mise à jour, et gestion des relations Person<->OrgUnit."""

    def __init__(self, base_url):
        self.base_url = base_url.rstrip("/")
        self.session = requests.Session()
        self._identifiants = None  # (utilisateur, mot_de_passe), mémorisés pour reconnexion auto
        self._cache_orgunits = {}  # code_unite -> (uuid, nom) ou None si introuvable
        self._type_relation_orgunit = None  # (id, entite_gauche, entite_droite), résolu une fois

    def _jeton_csrf(self):
        return self.session.cookies.get(COOKIE_CSRF)

    def _requete(self, methode, url, csrf=False, _reessai_auth=True, **kwargs):
        """Point de passage unique pour tous les appels HTTP :
        - reconnexion automatique si le jeton a expiré (401) ;
        - nouvelle tentative avec délai croissant sur erreur réseau ou
          erreur serveur transitoire (502/503/504) ;
        - csrf=True : (ré)injecte un jeton CSRF à jour à chaque tentative
          (nécessaire après une reconnexion, qui en génère un nouveau).

        Une requête d'écriture (POST/PATCH) n'est relancée que si elle n'a
        jamais atteint le serveur (ConnectTimeout) : après une réponse perdue
        (délai de lecture dépassé, 502/504...), le serveur a pu l'appliquer
        quand même, et la relancer risquerait de créer un doublon. Elle compte
        alors comme une erreur, et le run suivant la rattrapera."""
        lecture_seule = methode.upper() == "GET"
        derniere_erreur = None
        for tentative in range(1, NB_TENTATIVES_MAX + 1):
            appel_kwargs = dict(kwargs)
            appel_kwargs.setdefault("timeout", TIMEOUT_HTTP)
            if csrf:
                en_tetes = dict(appel_kwargs.get("headers", {}))
                en_tetes[EN_TETE_CSRF] = self._jeton_csrf()
                appel_kwargs["headers"] = en_tetes

            try:
                r = self.session.request(methode, url, **appel_kwargs)
            except requests.exceptions.RequestException as e:
                derniere_erreur = e
                peut_reessayer = lecture_seule or isinstance(e, requests.exceptions.ConnectTimeout)
                if peut_reessayer and tentative < NB_TENTATIVES_MAX:
                    delai = DELAI_BACKOFF_SECONDES * (2 ** (tentative - 1))
                    LOG.warning(f"    (connexion échouée, nouvel essai dans {delai}s : {e})")
                    time.sleep(delai)
                    continue
                raise

            if r.status_code == 401 and _reessai_auth and self._identifiants:
                LOG.info("    (jeton expiré, reconnexion automatique...)")
                self._authentifier(*self._identifiants)
                return self._requete(methode, url, csrf=csrf, _reessai_auth=False, **kwargs)

            if r.status_code in CODES_HTTP_A_REESSAYER and lecture_seule and tentative < NB_TENTATIVES_MAX:
                delai = DELAI_BACKOFF_SECONDES * (2 ** (tentative - 1))
                LOG.warning(f"    (erreur serveur {r.status_code}, nouvel essai dans {delai}s)")
                time.sleep(delai)
                continue

            return r

        raise derniere_erreur

    def se_connecter(self, utilisateur, mot_de_passe):
        """Authentifie la session et mémorise les identifiants pour pouvoir
        se reconnecter automatiquement si le jeton expire en cours de route."""
        self._identifiants = (utilisateur, mot_de_passe)
        self._authentifier(utilisateur, mot_de_passe)

    def _authentifier(self, utilisateur, mot_de_passe):
        """Récupère le cookie CSRF, puis échange les identifiants contre un
        jeton JWT (conservé dans les en-têtes de session).

        allow_redirects=False : si le serveur répond par une redirection
        (301/302), 'requests' la suit par défaut en repassant la requête en
        GET, ce qui casse la connexion (DSpace exige explicitement un POST
        sur /authn/login). On détecte donc la redirection nous-mêmes et on
        renvoie le POST vers la nouvelle URL en gardant la méthode."""
        self.session.get(f"{self.base_url}/authn/status", timeout=TIMEOUT_HTTP).raise_for_status()

        url = f"{self.base_url}/authn/login"
        for _ in range(5):  # au cas où plusieurs redirections s'enchaînent
            r = self.session.post(
                url,
                data={"user": utilisateur, "password": mot_de_passe},
                headers={EN_TETE_CSRF: self._jeton_csrf()},
                allow_redirects=False,
                timeout=TIMEOUT_HTTP,
            )
            if r.status_code in (301, 302, 307, 308) and "Location" in r.headers:
                url = r.headers["Location"]
                continue
            break

        if r.status_code != 200:
            raise RuntimeError(f"Échec de connexion DSpace ({r.status_code}) : {r.text[:300]}")

        jeton_auth = r.headers.get(EN_TETE_AUTORISATION)
        if not jeton_auth:
            raise RuntimeError("Connexion réussie mais aucun jeton d'autorisation reçu.")
        self.session.headers[EN_TETE_AUTORISATION] = jeton_auth

    def resoudre_collection(self, communaute_uuid):
        """Trouve la collection à utiliser dans une communauté. S'il y en a
        plusieurs, les liste et demande à l'appelant de préciser --collection."""
        r = self._requete("GET", f"{self.base_url}/core/communities/{communaute_uuid}/collections")
        r.raise_for_status()
        collections = r.json().get("_embedded", {}).get("collections", [])

        if not collections:
            raise RuntimeError(f"Aucune collection trouvée dans la communauté {communaute_uuid}.")
        if len(collections) == 1:
            return collections[0]["uuid"]

        LOG.info("Plusieurs collections trouvées — précise --collection avec l'UUID voulu :")
        for c in collections:
            LOG.info(f"  {c['uuid']}  {c.get('name')}")
        sys.exit(1)

    def chercher_par_courriel(self, courriel):
        """Cherche un item existant dont CHAMP_COURRIEL correspond exactement
        (après normalisation) au courriel donné."""
        r = self._requete(
            "GET", f"{self.base_url}/discover/search/objects",
            params={"query": f'"{courriel}"', "dsoType": "item", "size": TAILLE_PAGE_RECHERCHE},
        )
        r.raise_for_status()
        objets = (r.json().get("_embedded", {}).get("searchResult", {})
                  .get("_embedded", {}).get("objects", []))

        cible = normaliser_courriel(courriel)
        for obj in objets:
            item = obj.get("_embedded", {}).get("indexableObject", {})
            valeurs = item.get("metadata", {}).get(CHAMP_COURRIEL, [])
            if any(normaliser_courriel(v.get("value", "")) == cible for v in valeurs):
                return item
        return None

    def chercher_orgunit_par_code(self, code_unite):
        """Cherche l'item OrgUnit dont CHAMP_CODEUNITE correspond exactement au
        code donné, et retourne (uuid, code) ou None si introuvable. Résultat
        mis en cache : plusieurs profs partagent souvent le même CodeUnite.
        Le 'code' sert juste à des fins d'affichage dans les logs (le lien
        lui-même ne repose que sur l'UUID)."""
        if code_unite in self._cache_orgunits:
            return self._cache_orgunits[code_unite]

        r = self._requete(
            "GET", f"{self.base_url}/discover/search/objects",
            params={"query": f'"{code_unite}"', "dsoType": "item", "size": TAILLE_PAGE_RECHERCHE},
        )
        r.raise_for_status()
        objets = (r.json().get("_embedded", {}).get("searchResult", {})
                  .get("_embedded", {}).get("objects", []))

        resultat = None
        cible = code_unite.strip()
        for obj in objets:
            item = obj.get("_embedded", {}).get("indexableObject", {})
            metadonnees = item.get("metadata", {})
            type_entite = metadonnees.get(CHAMP_TYPE_ENTITE, [{}])[0].get("value")
            valeurs_code = [v.get("value", "").strip() for v in metadonnees.get(CHAMP_CODEUNITE, [])]
            if type_entite == VALEUR_TYPE_ENTITE_ORGUNIT and cible in valeurs_code:
                resultat = (item["uuid"], cible)
                break

        self._cache_orgunits[code_unite] = resultat
        return resultat

    def _label_type_entite(self, href):
        """Suit un lien _links.leftType/rightType.href et retourne le 'label'
        de l'EntityType (ex: 'Person', 'OrgUnit')."""
        r = self._requete("GET", href)
        r.raise_for_status()
        return r.json().get("label")

    def resoudre_type_relation(self, nom_relation):
        """Trouve, dans /core/relationshiptypes, le type dont leftwardType OU
        rightwardType correspond à nom_relation, et retourne
        (type_id, entite_gauche, entite_droite). Mis en cache (au plus 3 appels
        API pour toute la durée du script : la liste + les 2 EntityType).

        On suit les liens _links.leftType/rightType.href directement (plutôt
        que de compter sur le paramètre 'embed', pas toujours honoré selon la
        configuration/version) pour être sûr d'obtenir le vrai label."""
        if self._type_relation_orgunit is not None:
            return self._type_relation_orgunit

        r = self._requete("GET", f"{self.base_url}/core/relationshiptypes", params={"size": 100})
        r.raise_for_status()
        types = r.json().get("_embedded", {}).get("relationshiptypes", [])

        for t in types:
            if nom_relation in (t.get("leftwardType"), t.get("rightwardType")):
                liens = t.get("_links", {})
                href_gauche = liens.get("leftType", {}).get("href")
                href_droite = liens.get("rightType", {}).get("href")
                gauche = self._label_type_entite(href_gauche) if href_gauche else None
                droite = self._label_type_entite(href_droite) if href_droite else None
                self._type_relation_orgunit = (t["id"], gauche, droite)
                return self._type_relation_orgunit

        raise RuntimeError(f"Type de relation '{nom_relation}' introuvable dans DSpace "
                            f"(vérifier /core/relationshiptypes).")

    def lister_relations(self, item_uuid):
        """Retourne les relations de item_uuid sous forme de dicts
        {id, autre_uuid, nom_type} : autre_uuid = l'item à l'autre bout (None
        si l'API ne l'a pas inclus), nom_type = (leftwardType, rightwardType)
        du relationshipType (None si non inclus)."""
        r = self._requete(
            "GET", f"{self.base_url}/core/items/{item_uuid}/relationships",
            params={"embed": "leftItem,rightItem,relationshipType", "size": 100},
        )
        r.raise_for_status()
        resultat = []
        for relation in r.json().get("_embedded", {}).get("relationships", []):
            embed = relation.get("_embedded", {})
            gauche = embed.get("leftItem", {}).get("uuid")
            droite = embed.get("rightItem", {}).get("uuid")
            type_relation = embed.get("relationshipType") or {}
            resultat.append({
                "id": relation.get("id"),
                "autre_uuid": droite if gauche == item_uuid else gauche,
                "nom_type": ((type_relation.get("leftwardType"), type_relation.get("rightwardType"))
                             if type_relation else None),
            })
        return resultat

    def a_deja_une_relation_avec(self, item_uuid, autre_item_uuid):
        """Vérifie si item_uuid a déjà une relation (peu importe le type) avec
        autre_item_uuid, pour éviter de créer un doublon."""
        return any(rel["autre_uuid"] == autre_item_uuid for rel in self.lister_relations(item_uuid))

    def supprimer_relation(self, relation_id):
        r = self._requete("DELETE", f"{self.base_url}/core/relationships/{relation_id}", csrf=True)
        if r.status_code not in (200, 204):
            raise RuntimeError(f"Échec de suppression de la relation {relation_id} ({r.status_code}) : {r.text[:300]}")

    def creer_relation(self, type_id, item_gauche_uuid, item_droit_uuid):
        """Crée la relation via l'API dédiée (Content-Type: text/uri-list,
        un item par ligne : le premier est l'item 'gauche', le second le
        'droit', selon leftType/rightType du relationshipType)."""
        corps = f"{self.base_url}/core/items/{item_gauche_uuid}\n{self.base_url}/core/items/{item_droit_uuid}"
        r = self._requete(
            "POST", f"{self.base_url}/core/relationships",
            params={"relationshipType": type_id}, data=corps,
            headers={"Content-Type": "text/uri-list"}, csrf=True,
        )
        if r.status_code not in CODES_HTTP_SUCCES:
            raise RuntimeError(
                f"Échec de création de la relation ({r.status_code}) : {r.text[:300]}\n"
                f"    Requête envoyée : POST {self.base_url}/core/relationships?relationshipType={type_id}\n"
                f"    Corps (text/uri-list) :\n"
                f"      {self.base_url}/core/items/{item_gauche_uuid}\n"
                f"      {self.base_url}/core/items/{item_droit_uuid}"
            )
        return r.json()

    def creer_personne(self, collection_uuid, metadonnees, titre_affichage):
        r = self._requete(
            "POST", f"{self.base_url}/core/items",
            params={"owningCollection": collection_uuid},
            json={
                "name": titre_affichage,
                "metadata": metadonnees,
                "inArchive": True,
                "discoverable": True,
                "withdrawn": False,
                "type": "item",
            },
            headers={"Content-Type": "application/json"}, csrf=True,
        )
        if r.status_code not in CODES_HTTP_SUCCES:
            raise RuntimeError(f"Échec de création ({r.status_code}) : {r.text[:300]}")
        return r.json()

    def mettre_a_jour_personne(self, item_uuid, operations):
        r = self._requete(
            "PATCH", f"{self.base_url}/core/items/{item_uuid}",
            json=operations, headers={"Content-Type": "application/json"}, csrf=True,
        )
        if r.status_code not in CODES_HTTP_SUCCES:
            raise RuntimeError(f"Échec de mise à jour ({r.status_code}) : {r.text[:300]}")
        return r.json()

    def lister_items_de_la_collection(self, collection_uuid, taille_page=100):
        """Retourne TOUS les items d'une collection (pagine automatiquement),
        via l'API Discovery (scope=collection_uuid) — /core/collections/{uuid}/items
        n'existe pas dans l'API REST DSpace 7. Utilisé pour détecter les profs
        présents dans DSpace mais absents de la liste_personnel actuelle."""
        items = []
        page = 0
        while True:
            r = self._requete(
                "GET", f"{self.base_url}/discover/search/objects",
                params={"scope": collection_uuid, "dsoType": "item", "size": taille_page, "page": page},
            )
            r.raise_for_status()
            resultat_recherche = r.json().get("_embedded", {}).get("searchResult", {})
            objets = resultat_recherche.get("_embedded", {}).get("objects", [])
            for obj in objets:
                item = obj.get("_embedded", {}).get("indexableObject", {})
                if item:
                    items.append(item)

            info_page = resultat_recherche.get("page", {})
            if page + 1 >= info_page.get("totalPages", 1):
                break
            page += 1
        return items


# =============================================================================
# CONSTRUCTION DES MÉTADONNÉES / OPÉRATIONS JSON PATCH
# =============================================================================

def construire_metadonnees(nom, prenom, courriel, orcid, statut, code_unite):
    """Construit le dict de métadonnées DSpace pour la création d'un item Person.
    Ne contient PAS la relation OrgUnit : elle est créée séparément via
    l'API /core/relationships (voir lier_orgunit)."""
    metadonnees = {
        CHAMP_TYPE_ENTITE: [{"value": VALEUR_TYPE_ENTITE}],
        CHAMP_NOM: [{"value": nom}],
        CHAMP_PRENOM: [{"value": prenom}],
        CHAMP_COURRIEL: [{"value": courriel}],
        CHAMP_STATUT: [{"value": VALEUR_STATUT_ACTIF}],
    }
    if statut:
        metadonnees[CHAMP_FONCTION] = [{"value": statut}]
    if orcid:
        metadonnees[CHAMP_ORCID] = [{"value": orcid}]
    if code_unite:
        metadonnees[CHAMP_CODEUNITE] = [{"value": code_unite}]
    return metadonnees


def operation_maj_champ(item_existant, champ, valeur):
    """Construit une opération JSON Patch pour ajouter ou remplacer un champ
    d'un item existant, ou None si la valeur est déjà à jour."""
    valeurs_actuelles = item_existant.get("metadata", {}).get(champ, [])
    if not valeurs_actuelles:
        return {"op": "add", "path": f"/metadata/{champ}", "value": [{"value": valeur}]}
    if valeurs_actuelles[0].get("value") != valeur:
        return {"op": "replace", "path": f"/metadata/{champ}/0", "value": {"value": valeur}}
    return None


def operations_maj_personne(item_existant, nom, prenom, statut, orcid, code_unite):
    """Liste toutes les opérations JSON Patch nécessaires pour mettre à jour
    un item Person existant (seuls les champs différents génèrent une opération).
    La relation OrgUnit n'est PAS gérée ici, voir lier_orgunit."""
    champs_a_verifier = [
        (CHAMP_NOM, nom),
        (CHAMP_PRENOM, prenom),
        (CHAMP_FONCTION, statut),
        (CHAMP_ORCID, orcid),
        (CHAMP_CODEUNITE, code_unite),
        (CHAMP_STATUT, VALEUR_STATUT_ACTIF),  # présent dans le fichier -> toujours Actif
    ]
    operations = []
    for champ, valeur in champs_a_verifier:
        if not valeur:
            continue
        operation = operation_maj_champ(item_existant, champ, valeur)
        if operation:
            operations.append(operation)
    return operations


def retirer_anciennes_unites(client, person_uuid, orgunit_uuid, relations, appliquer):
    """Retire les relations NOM_RELATION_ORGUNIT de la personne qui pointent
    vers une AUTRE OrgUnit que orgunit_uuid (prof qui a changé d'unité).
    Une relation n'est retirée que si son type ET l'item à l'autre bout sont
    connus avec certitude — dans le doute, on n'y touche pas. Retourne le
    nombre de relations retirées (ou qui le seraient, en simulation)."""
    if not RETIRER_ANCIENNES_UNITES:
        return 0
    anciennes = [rel for rel in relations
                 if rel["nom_type"] and NOM_RELATION_ORGUNIT in rel["nom_type"]
                 and rel["autre_uuid"] and rel["autre_uuid"] != orgunit_uuid]
    for rel in anciennes:
        LOG.debug(f"  ↳ retrait du lien vers l'ancienne OrgUnit {rel['autre_uuid']} (relation {rel['id']})")
        if appliquer:
            client.supprimer_relation(rel["id"])
    return len(anciennes)


def lier_orgunit(client, person_uuid, orgunit_uuid, appliquer):
    """Crée la relation Person<->OrgUnit si elle n'existe pas déjà, puis
    retire les liens vers une ancienne OrgUnit (voir retirer_anciennes_unites).
    Le nouveau lien est créé AVANT le retrait de l'ancien : la fiche n'est
    jamais sans unité. En simulation, seules des lectures sont faites.
    Retourne (statut, nb_retirees) où statut vaut 'lie' ou 'deja_lie' ;
    lève une exception en cas d'erreur."""
    relations = client.lister_relations(person_uuid)
    if any(rel["autre_uuid"] == orgunit_uuid for rel in relations):
        return "deja_lie", retirer_anciennes_unites(client, person_uuid, orgunit_uuid, relations, appliquer)

    if not appliquer:
        return "lie", retirer_anciennes_unites(client, person_uuid, orgunit_uuid, relations, appliquer)

    type_id, entite_gauche, entite_droite = client.resoudre_type_relation(NOM_RELATION_ORGUNIT)
    if entite_gauche == VALEUR_TYPE_ENTITE_ORGUNIT:
        item_gauche, item_droit = orgunit_uuid, person_uuid
    elif entite_droite == VALEUR_TYPE_ENTITE_ORGUNIT:
        item_gauche, item_droit = person_uuid, orgunit_uuid
    else:
        raise RuntimeError(f"Le type de relation '{NOM_RELATION_ORGUNIT}' ne concerne pas "
                            f"un {VALEUR_TYPE_ENTITE_ORGUNIT} (gauche={entite_gauche}, droite={entite_droite}).")

    try:
        client.creer_relation(type_id, item_gauche, item_droit)
    except Exception:
        # La création a échoué : avant de considérer ça comme une vraie
        # erreur, on revérifie si le lien existe déjà (ex: créé entre-temps,
        # ou un délai d'indexation côté serveur a empêché la première
        # vérification de le voir). Si c'est le cas -> pas une erreur, on
        # voulait ce lien et il y est déjà. Sinon -> vraie erreur, on la
        # laisse remonter telle quelle pour ne rien cacher.
        if not client.a_deja_une_relation_avec(person_uuid, orgunit_uuid):
            raise
    return "lie", retirer_anciennes_unites(client, person_uuid, orgunit_uuid, relations, appliquer)


def marquer_departs(client, collection_uuid, courriels_actifs, appliquer, ignorer_seuils=False):
    """Parcourt tous les items Person de la collection et marque
    CHAMP_STATUT='Inactif' sur ceux dont le courriel n'est PAS dans
    courriels_actifs (= absent de la liste_personnel actuelle). Ne touche
    jamais aux items déjà marqués Inactif, et ne retire/supprime jamais
    rien — juste une métadonnée, à réviser manuellement au besoin.

    Garde-fou : si plus de SEUIL_MAX_DEPARTS_POURCENT % de la collection
    devrait passer Inactif d'un coup, rien n'est marqué (fichier d'entrée
    probablement incomplet), sauf avec ignorer_seuils.

    Retourne (nb_departs, bloque, nb_erreurs) : nb_departs = items marqués
    (ou qui le seraient, en simulation / si bloqué)."""
    items = client.lister_items_de_la_collection(collection_uuid)
    a_marquer = []
    for item in items:
        valeurs_courriel = item.get("metadata", {}).get(CHAMP_COURRIEL, [])
        courriel_item = normaliser_courriel(valeurs_courriel[0].get("value")) if valeurs_courriel else ""
        if courriel_item in courriels_actifs:
            continue  # présent dans la liste actuelle -> déjà traité comme Actif ailleurs

        operation = operation_maj_champ(item, CHAMP_STATUT, VALEUR_STATUT_INACTIF)
        if not operation:
            continue  # déjà Inactif

        a_marquer.append((item, courriel_item, operation))

    pourcentage = 100 * len(a_marquer) / len(items) if items else 0
    if pourcentage > SEUIL_MAX_DEPARTS_POURCENT and not ignorer_seuils:
        LOG.error(f"GARDE-FOU : {len(a_marquer)} départ(s) détecté(s) sur {len(items)} item(s) "
                   f"({pourcentage:.0f}% > {SEUIL_MAX_DEPARTS_POURCENT}%) — aucun n'est marqué Inactif. "
                   f"Vérifier le fichier d'entrée ; si ces départs sont réels, relancer avec --ignorer-seuils.")
        return len(a_marquer), True, 0

    nb_erreurs = 0
    for item, courriel_item, operation in a_marquer:
        LOG.debug(f"[INACTIF] <{courriel_item or 'sans courriel'}> ({item.get('uuid')})")
        if appliquer:
            try:
                client.mettre_a_jour_personne(item["uuid"], [operation])
            except Exception as e:
                LOG.error(f"[ERREUR] marquage Inactif de <{courriel_item}> : {e}")
                nb_erreurs += 1

    return len(a_marquer), False, nb_erreurs


# =============================================================================
# SYNCHRONISATION D'UN PROFESSEUR (création ou mise à jour)
# =============================================================================

def synchroniser_professeur(client, collection_uuid, professeur, orcid, appliquer):
    """Crée ou met à jour un professeur dans DSpace, et s'assure qu'il est lié
    à son OrgUnit (relation créée séparément des métadonnées). Retourne
    (resultat, lien_echoue, orgunit_introuvable, nb_unites_retirees) où
    resultat est l'une des chaînes 'cree', 'maj', 'inchange'. Un échec de
    LIAISON OrgUnit n'interrompt
    jamais la création/mise à jour de la fiche Person elle-même ; le détail
    (créations, liaisons, avertissements) va dans le fichier .log mais pas la
    console par défaut (voir --verbose) — seule une vraie erreur sur la fiche
    Person lève une exception et s'affiche toujours."""
    nom, prenom, courriel, code_unite, statut = professeur
    item_existant = client.chercher_par_courriel(courriel)

    orgunit = client.chercher_orgunit_par_code(code_unite) if code_unite else None
    orgunit_introuvable = bool(code_unite and orgunit is None)
    if orgunit_introuvable:
        LOG.debug(f"  ⚠ Aucun OrgUnit trouvé pour le CodeUnite '{code_unite}' ({nom}, {prenom})")

    if item_existant is None:
        LOG.debug(f"[NOUVEAU] {nom}, {prenom} <{courriel}> ORCID={orcid or '—'}")
        person_uuid = None
        if appliquer:
            metadonnees = construire_metadonnees(nom, prenom, courriel, orcid, statut, code_unite)
            person_uuid = client.creer_personne(collection_uuid, metadonnees, f"{nom}, {prenom}")["uuid"]
        lien_echoue = False
        if orgunit and person_uuid:
            try:
                statut_lien, _ = lier_orgunit(client, person_uuid, orgunit[0], appliquer)
                if statut_lien == "lie":
                    LOG.debug(f"  ↳ lié à l'OrgUnit '{orgunit[1]}'")
            except Exception as e:
                LOG.debug(f"  ⚠ liaison OrgUnit '{orgunit[1]}' échouée ({nom}, {prenom}) : {e}")
                lien_echoue = True
        return "cree", lien_echoue, orgunit_introuvable, 0

    operations = operations_maj_personne(item_existant, nom, prenom, statut, orcid, code_unite)
    if operations:
        LOG.debug(f"[MISE À JOUR] {nom}, {prenom} <{courriel}> — {len(operations)} champ(s) à modifier")
        if appliquer:
            client.mettre_a_jour_personne(item_existant["uuid"], operations)

    resultat = "maj" if operations else "inchange"

    lien_echoue = False
    nb_retirees = 0
    if orgunit:
        try:
            statut_lien, nb_retirees = lier_orgunit(client, item_existant["uuid"], orgunit[0], appliquer)
            if statut_lien == "lie":
                LOG.debug(f"  ↳ lié à l'OrgUnit '{orgunit[1]}' ({nom}, {prenom})")
            if statut_lien == "lie" or nb_retirees:
                resultat = "maj"  # une relation a changé, même si aucune métadonnée n'a changé
        except Exception as e:
            LOG.debug(f"  ⚠ liaison OrgUnit '{orgunit[1]}' échouée ({nom}, {prenom}) : {e}")
            lien_echoue = True

    return resultat, lien_echoue, orgunit_introuvable, nb_retirees


# =============================================================================
# POINT D'ENTRÉE
# =============================================================================

def analyser_arguments():
    parser = argparse.ArgumentParser(
        description="Synchronise les professeurs extraits avec un dépôt DSpace 7 (crée/met à jour "
                     "les items Person par correspondance de courriel)."
    )
    parser.add_argument("fichier_txt", nargs="?",
                         help="Fichier Synchro (.txt à largeur fixe), dossier ou motif avec * (le plus "
                              "récent est choisi). Défaut : FICHIER_PERSONNEL du .env")
    parser.add_argument("--excel", metavar="FICHIER.xlsx",
                         help="Fichier Excel contenant les ORCID (colonnes 'Courriel' et 'ORCID'), dossier "
                              "ou motif avec * (le plus récent est choisi). Défaut : FICHIER_ORCID du .env")
    parser.add_argument("--base-url", help="URL de base de l'API REST DSpace (défaut : DSPACE_BASE_URL du .env)")
    parser.add_argument("--user", help="Utilisateur admin DSpace (défaut : DSPACE_USER du .env)")
    parser.add_argument("--password", help="Mot de passe admin DSpace (défaut : DSPACE_PASSWORD du .env)")
    parser.add_argument("--community", default=COMMUNAUTE_PAR_DEFAUT,
                         help="UUID de la communauté où chercher/créer la collection Person")
    parser.add_argument("--collection", default=COLLECTION_PAR_DEFAUT,
                         help="UUID de la collection cible, défaut : DSPACE_COLLECTION du .env (évite la résolution "
                                              "automatique via --community)")
    parser.add_argument("--apply", action="store_true",
                         help="Applique réellement les changements (sinon, mode simulation)")
    parser.add_argument("--limit", type=int, metavar="N",
                         help="Ne traite que les N premiers professeurs trouvés (utile pour tester) ; "
                              "l'étape de marquage des départs est alors sautée")
    parser.add_argument("--ignorer-seuils", action="store_true",
                         help=f"Désactive les garde-fous (min. {MIN_PROFESSEURS_ATTENDUS} profs dans le "
                              f"fichier, max. {SEUIL_MAX_CREATIONS} créations, max. "
                              f"{SEUIL_MAX_DEPARTS_POURCENT}%% de départs) — ex. premier chargement complet")
    parser.add_argument("--log-file", metavar="FICHIER.log",
                         help="Chemin du fichier de journal (par défaut : horodaté, dans logs/)")
    parser.add_argument("--verbose", action="store_true",
                         help="Affiche le détail ligne par ligne dans la console (par défaut, "
                              "seul le fichier .log le garde ; la console reste compacte)")
    args = parser.parse_args()

    # Paramètres de connexion venus ni de la ligne de commande ni du .env : le
    # script retombe sur les valeurs codées en dur (refusé en --apply, voir executer).
    args.connexion_par_defaut = [variable for option, variable in CONNEXION_OBLIGATOIRE.items()
                                 if getattr(args, option) is None and variable not in os.environ]
    args.base_url = args.base_url or BASE_URL_PAR_DEFAUT
    args.user = args.user or UTILISATEUR_PAR_DEFAUT
    args.password = args.password or MOT_DE_PASSE_PAR_DEFAUT

    # Fichiers absents de la commande : ceux du .env, relatifs au dossier du script.
    args.fichier_txt = args.fichier_txt or os.path.join(DOSSIER_SCRIPT, FICHIER_PERSONNEL_PAR_DEFAUT)
    args.excel = args.excel or os.path.join(DOSSIER_SCRIPT, FICHIER_ORCID_PAR_DEFAUT)
    return args


def acquerir_verrou(chemin):
    """Pose un verrou exclusif au niveau du système d'exploitation sur
    'chemin' et retourne le fichier ouvert (à garder ouvert pendant tout le
    run), ou None si une autre synchronisation le tient déjà. Le système
    libère ce verrou tout seul si le processus meurt (plantage, kill) : pas
    de verrou orphelin à nettoyer à la main."""
    f = open(chemin, "a+")
    try:
        if os.name == "nt":
            import msvcrt
            f.seek(0)
            msvcrt.locking(f.fileno(), msvcrt.LK_NBLCK, 1)
        else:
            import fcntl
            fcntl.flock(f, fcntl.LOCK_EX | fcntl.LOCK_NB)
    except OSError:
        f.close()
        return None
    return f


def date_du_fichier(chemin):
    """Date d'un fichier d'entrée : celle écrite dans son nom (AAAAMMJJ, ex.
    synchro_PERSONNEL_20261001.txt) si présente, sinon sa date de modification."""
    m = re.search(r"(\d{8})", os.path.basename(chemin))
    if m:
        try:
            return datetime.strptime(m.group(1), "%Y%m%d")
        except ValueError:
            pass
    return datetime.fromtimestamp(os.path.getmtime(chemin))


def resoudre_fichier_entree(chemin, motif):
    """Retourne 'chemin' tel quel si c'est un fichier ; si c'est un motif avec
    * (ou un dossier, où l'on cherche 'motif'), retourne le fichier
    correspondant le plus récent."""
    if os.path.isdir(chemin):
        chemin = os.path.join(chemin, motif)
    elif not glob.has_magic(chemin):
        return chemin
    candidats = [f for f in glob.glob(chemin) if os.path.isfile(f)]
    if not candidats:
        raise FileNotFoundError(f"aucun fichier ne correspond à '{chemin}'")
    return max(candidats, key=date_du_fichier)


def verifier_fraicheur(chemin, age_max_jours, description):
    """Avertit si le fichier est plus vieux que age_max_jours. Retourne True
    s'il est périmé."""
    age = (datetime.now() - date_du_fichier(chemin)).days
    if age > age_max_jours:
        LOG.warning(f"⚠ Fichier {description} périmé : '{chemin}' date de {age} jour(s) "
                     f"(maximum {age_max_jours}) — export non reçu ?")
        return True
    return False


def executer(args):
    """Déroule toute la synchronisation et retourne un code de sortie
    (CODE_SORTIE_*)."""
    if args.apply and args.connexion_par_defaut:
        LOG.error(f"Mode --apply refusé : {', '.join(args.connexion_par_defaut)} absent(s) du .env "
                   f"et de la ligne de commande (les valeurs de secours codées en dur ne sont "
                   f"jamais utilisées pour écrire). Vérifier que {os.path.join(DOSSIER_SCRIPT, '.env')} "
                   f"existe et que python-dotenv est installé.")
        return CODE_SORTIE_ERREUR_FATALE

    args.fichier_txt = resoudre_fichier_entree(args.fichier_txt, MOTIF_FICHIER_PERSONNEL)
    args.excel = resoudre_fichier_entree(args.excel, MOTIF_FICHIER_ORCID)
    LOG.info(f"Fichiers : {args.fichier_txt} | {args.excel}")
    fichiers_perimes = verifier_fraicheur(args.fichier_txt, AGE_MAX_FICHIER_PERSONNEL_JOURS, "liste_personnel")
    fichiers_perimes |= verifier_fraicheur(args.excel, AGE_MAX_FICHIER_ORCID_JOURS, "ORCID")

    lignes = lire_lignes(args.fichier_txt)
    professeurs_bruts = filtrer_professeurs(lignes)
    professeurs = [extraire_infos(ligne) + [statut] for ligne, statut in professeurs_bruts]
    orcid_par_courriel = charger_orcid_depuis_excel(args.excel)

    # Vérifié sur le fichier complet (avant --limit) : un fichier vide ou tronqué
    # ferait sinon passer presque toute la collection Inactif.
    if len(professeurs) < MIN_PROFESSEURS_ATTENDUS and not args.ignorer_seuils:
        LOG.error(f"GARDE-FOU : seulement {len(professeurs)} professeur(s) trouvé(s) dans "
                   f"'{args.fichier_txt}' (minimum attendu : {MIN_PROFESSEURS_ATTENDUS}) — fichier "
                   f"probablement vide, tronqué ou au mauvais format. Rien n'a été fait. "
                   f"Relancer avec --ignorer-seuils si c'est voulu.")
        return CODE_SORTIE_ERREUR_FATALE

    if args.limit is not None:
        professeurs = professeurs[:args.limit]

    total = len(professeurs)

    client = DSpaceClient(args.base_url)
    try:
        client.se_connecter(args.user, args.password)
    except Exception as e:
        LOG.error(f"Erreur de connexion à DSpace : {e}")
        return CODE_SORTIE_ERREUR_FATALE

    try:
        collection_uuid = args.collection or client.resoudre_collection(args.community)
    except Exception as e:
        LOG.error(f"Erreur lors de la résolution de la collection : {e}")
        return CODE_SORTIE_ERREUR_FATALE

    try:
        type_id, gauche, droite = client.resoudre_type_relation(NOM_RELATION_ORGUNIT)
        LOG.debug(f"Type de relation '{NOM_RELATION_ORGUNIT}' résolu : id={type_id}, "
                   f"gauche={gauche}, droite={droite}")
    except Exception as e:
        LOG.warning(f"Impossible de résoudre le type de relation '{NOM_RELATION_ORGUNIT}' : {e}")

    mode = "APPLICATION RÉELLE" if args.apply else "SIMULATION"
    limite = f", limité à {args.limit}" if args.limit is not None else ""
    LOG.info(f"{total} professeur(s){limite} | collection {collection_uuid} | mode {mode}\n")

    compteurs = {"cree": 0, "maj": 0, "inchange": 0, "erreur": 0, "liens_echoues": 0,
                 "orgunit_introuvable": 0, "unites_retirees": 0}
    heure_depart = time.time()
    arret_creations = False
    arret_erreurs = False
    erreurs_consecutives = 0

    for i, professeur in enumerate(professeurs, start=1):
        courriel = professeur[2]
        orcids = orcid_par_courriel.get(normaliser_courriel(courriel), [])
        orcid = orcids[0] if orcids else ""

        try:
            resultat, lien_echoue, orgunit_introuvable, nb_retirees = synchroniser_professeur(
                client, collection_uuid, professeur, orcid, args.apply
            )
            compteurs[resultat] += 1
            compteurs["unites_retirees"] += nb_retirees
            if lien_echoue:
                compteurs["liens_echoues"] += 1
            if orgunit_introuvable:
                compteurs["orgunit_introuvable"] += 1
            erreurs_consecutives = 0
        except Exception as e:
            LOG.error(f"[ERREUR] {courriel} : {e}")
            compteurs["erreur"] += 1
            erreurs_consecutives += 1

        if i % INTERVALLE_PROGRES == 0 or i == total:
            ecoule = time.time() - heure_depart
            rythme = i / ecoule if ecoule > 0 else 0
            restant = (total - i) / rythme if rythme > 0 else 0
            LOG.info(f"{i}/{total} ({100 * i / total:.0f}%) — "
                      f"{compteurs['cree']}c {compteurs['maj']}m {compteurs['inchange']}i "
                      f"{compteurs['erreur']}e — ~{restant / 60:.0f} min restantes")

        # Beaucoup de créations d'un coup = profs existants introuvables par la
        # recherche (index Discovery vide ou en reconstruction) -> on arrête avant
        # de créer des centaines de doublons.
        if args.apply and not args.ignorer_seuils and compteurs["cree"] >= SEUIL_MAX_CREATIONS:
            LOG.error(f"GARDE-FOU : {compteurs['cree']} création(s) atteintes (maximum "
                       f"{SEUIL_MAX_CREATIONS} par run) — synchronisation arrêtée après {i}/{total} "
                       f"prof(s). Vérifier que l'index de recherche DSpace est complet ; pour un "
                       f"premier chargement, relancer avec --ignorer-seuils.")
            arret_creations = True
            break

        if erreurs_consecutives >= MAX_ERREURS_CONSECUTIVES:
            LOG.error(f"ARRÊT : {erreurs_consecutives} erreurs d'affilée — DSpace semble indisponible. "
                       f"Synchronisation abandonnée après {i}/{total} prof(s) (le run suivant "
                       f"reprendra là où il faut, rien n'est perdu). Dernière erreur ci-dessus.")
            arret_erreurs = True
            break

    nb_departs, departs_bloques, erreurs_departs, departs_echoues = 0, False, 0, False
    if arret_creations or arret_erreurs:
        LOG.warning("Marquage des départs sauté (synchronisation interrompue avant la fin).")
    elif args.limit is not None:
        LOG.info("Marquage des départs sauté (--limit : la liste traitée est incomplète).")
    else:
        courriels_actifs = {normaliser_courriel(p[2]) for p in professeurs}
        try:
            nb_departs, departs_bloques, erreurs_departs = marquer_departs(
                client, collection_uuid, courriels_actifs, args.apply, args.ignorer_seuils
            )
        except Exception as e:
            LOG.error(f"Impossible de vérifier les départs : {e}")
            departs_echoues = True

    LOG.info(f"\nRésumé : {compteurs['cree']} créé(s), {compteurs['maj']} mis à jour, "
              f"{compteurs['inchange']} inchangé(s), {compteurs['erreur']} erreur(s)")
    if nb_departs and not departs_bloques:
        verbe = "marqué(s)" if args.apply else "à marquer"
        LOG.info(f"  {nb_departs} prof(s) absent(s) de la liste_personnel {verbe} Inactif "
                  f"(présents dans DSpace mais plus dans le fichier)")
    if erreurs_departs:
        LOG.info(f"  {erreurs_departs} erreur(s) lors du marquage Inactif")
    if compteurs["liens_echoues"]:
        LOG.info(f"  {compteurs['liens_echoues']} liaison(s) OrgUnit échouée(s) (probablement déjà liées)")
    if compteurs["unites_retirees"]:
        verbe = "retiré(s)" if args.apply else "à retirer"
        LOG.info(f"  {compteurs['unites_retirees']} lien(s) vers une ancienne OrgUnit {verbe} (changement d'unité)")
    if compteurs["orgunit_introuvable"]:
        LOG.info(f"  {compteurs['orgunit_introuvable']} CodeUnite sans OrgUnit correspondant dans DSpace.")
    if not args.apply:
        if compteurs["cree"] >= SEUIL_MAX_CREATIONS and not args.ignorer_seuils:
            LOG.warning(f"  ⚠ {compteurs['cree']} création(s) prévue(s) : en mode --apply, le garde-fou "
                         f"arrêterait le run à {SEUIL_MAX_CREATIONS} (voir --ignorer-seuils).")
        LOG.info("Mode simulation : rien n'a été écrit dans DSpace. Relance avec --apply pour appliquer.")

    if (compteurs["erreur"] or erreurs_departs or departs_echoues
            or departs_bloques or arret_creations or fichiers_perimes):
        if arret_erreurs:
            return CODE_SORTIE_ERREUR_FATALE
        return CODE_SORTIE_ERREURS_PARTIELLES
    return CODE_SORTIE_OK


def main():
    args = analyser_arguments()

    chemin_log = args.log_file or os.path.join(DOSSIER_LOGS, f"synchro_profs_papyrus_{datetime.now():%Y%m%d_%H%M%S}.log")
    configurer_journalisation(chemin_log, verbose=args.verbose)

    verrou = None
    if args.apply:
        verrou = acquerir_verrou(FICHIER_VERROU)
        if verrou is None:
            LOG.error(f"Une autre synchronisation est déjà en cours (verrou : {FICHIER_VERROU}) — "
                       f"abandon pour éviter des doublons.")
            sys.exit(CODE_SORTIE_DEJA_EN_COURS)

    try:
        code = executer(args)
    except Exception:
        # Toute erreur imprévue (fichier introuvable, Excel illisible...) est
        # écrite dans le .log avec sa trace, pas seulement à l'écran.
        LOG.exception("Erreur inattendue — synchronisation interrompue :")
        code = CODE_SORTIE_ERREUR_FATALE
    finally:
        if verrou:
            verrou.close()

    if code != CODE_SORTIE_OK:
        LOG.info(f"Code de sortie : {code} (détails dans {chemin_log})")
    sys.exit(code)


if __name__ == "__main__":
    main()
