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
import os
import sys
import time
from datetime import datetime

import requests

try:
    from dotenv import load_dotenv
    if not load_dotenv():
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
DOSSIER_OUTILS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "outils")
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

# --- Emplacement de dépôt (surchargeable par --community/--collection) ---
# UUID de la communauté où chercher automatiquement la collection Person.
COMMUNAUTE_PAR_DEFAUT = os.environ.get("DSPACE_COMMUNITY", "1acd99a0-6ffb-42f8-a261-30b96f3f2405")

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

# --- Journalisation et suivi de progression ---
INTERVALLE_PROGRES = 100  # affiche un point d'étape tous les N professeurs traités
DOSSIER_LOGS = "logs"     # dossier où écrire les fichiers .log par défaut (créé si absent)


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
          (nécessaire après une reconnexion, qui en génère un nouveau)."""
        derniere_erreur = None
        for tentative in range(1, NB_TENTATIVES_MAX + 1):
            appel_kwargs = dict(kwargs)
            if csrf:
                en_tetes = dict(appel_kwargs.get("headers", {}))
                en_tetes[EN_TETE_CSRF] = self._jeton_csrf()
                appel_kwargs["headers"] = en_tetes

            try:
                r = self.session.request(methode, url, **appel_kwargs)
            except requests.exceptions.RequestException as e:
                derniere_erreur = e
                if tentative < NB_TENTATIVES_MAX:
                    delai = DELAI_BACKOFF_SECONDES * (2 ** (tentative - 1))
                    LOG.warning(f"    (connexion échouée, nouvel essai dans {delai}s : {e})")
                    time.sleep(delai)
                    continue
                raise

            if r.status_code == 401 and _reessai_auth and self._identifiants:
                LOG.info("    (jeton expiré, reconnexion automatique...)")
                self._authentifier(*self._identifiants)
                return self._requete(methode, url, csrf=csrf, _reessai_auth=False, **kwargs)

            if r.status_code in CODES_HTTP_A_REESSAYER and tentative < NB_TENTATIVES_MAX:
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
        self.session.get(f"{self.base_url}/authn/status").raise_for_status()

        url = f"{self.base_url}/authn/login"
        for _ in range(5):  # au cas où plusieurs redirections s'enchaînent
            r = self.session.post(
                url,
                data={"user": utilisateur, "password": mot_de_passe},
                headers={EN_TETE_CSRF: self._jeton_csrf()},
                allow_redirects=False,
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

    def a_deja_une_relation_avec(self, item_uuid, autre_item_uuid):
        """Vérifie si item_uuid a déjà une relation (peu importe le type) avec
        autre_item_uuid, pour éviter de créer un doublon."""
        r = self._requete(
            "GET", f"{self.base_url}/core/items/{item_uuid}/relationships",
            params={"embed": "leftItem,rightItem", "size": 100},
        )
        r.raise_for_status()
        relations = r.json().get("_embedded", {}).get("relationships", [])
        for relation in relations:
            embed = relation.get("_embedded", {})
            for cote in ("leftItem", "rightItem"):
                if embed.get(cote, {}).get("uuid") == autre_item_uuid:
                    return True
        return False

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


def lier_orgunit(client, person_uuid, orgunit_uuid, appliquer):
    """Crée la relation Person<->OrgUnit si elle n'existe pas déjà. Retourne
    'lie', 'deja_lie' ou lève une exception en cas d'erreur."""
    if not appliquer:
        return "lie"  # en simulation, on ne peut pas vérifier via l'API ; on suppose l'action

    if client.a_deja_une_relation_avec(person_uuid, orgunit_uuid):
        return "deja_lie"

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
        if client.a_deja_une_relation_avec(person_uuid, orgunit_uuid):
            return "deja_lie"
        raise
    return "lie"


def marquer_departs(client, collection_uuid, courriels_actifs, appliquer):
    """Parcourt tous les items Person de la collection et marque
    CHAMP_STATUT='Inactif' sur ceux dont le courriel n'est PAS dans
    courriels_actifs (= absent de la liste_personnel actuelle). Ne touche
    jamais aux items déjà marqués Inactif, et ne retire/supprime jamais
    rien — juste une métadonnée, à réviser manuellement au besoin.
    Retourne le nombre d'items marqués (ou qui le seraient, en simulation)."""
    a_marquer = 0
    for item in client.lister_items_de_la_collection(collection_uuid):
        valeurs_courriel = item.get("metadata", {}).get(CHAMP_COURRIEL, [])
        courriel_item = normaliser_courriel(valeurs_courriel[0].get("value")) if valeurs_courriel else ""
        if courriel_item in courriels_actifs:
            continue  # présent dans la liste actuelle -> déjà traité comme Actif ailleurs

        operation = operation_maj_champ(item, CHAMP_STATUT, VALEUR_STATUT_INACTIF)
        if not operation:
            continue  # déjà Inactif

        a_marquer += 1
        if appliquer:
            client.mettre_a_jour_personne(item["uuid"], [operation])

    return a_marquer


# =============================================================================
# SYNCHRONISATION D'UN PROFESSEUR (création ou mise à jour)
# =============================================================================

def synchroniser_professeur(client, collection_uuid, professeur, orcid, appliquer):
    """Crée ou met à jour un professeur dans DSpace, et s'assure qu'il est lié
    à son OrgUnit (relation créée séparément des métadonnées). Retourne
    (resultat, lien_echoue, orgunit_introuvable) où resultat est l'une des
    chaînes 'cree', 'maj', 'inchange'. Un échec de LIAISON OrgUnit n'interrompt
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
                statut_lien = lier_orgunit(client, person_uuid, orgunit[0], appliquer)
                if statut_lien == "lie":
                    LOG.debug(f"  ↳ lié à l'OrgUnit '{orgunit[1]}'")
            except Exception:
                lien_echoue = True
        return "cree", lien_echoue, orgunit_introuvable

    operations = operations_maj_personne(item_existant, nom, prenom, statut, orcid, code_unite)
    if operations:
        LOG.debug(f"[MISE À JOUR] {nom}, {prenom} <{courriel}> — {len(operations)} champ(s) à modifier")
        if appliquer:
            client.mettre_a_jour_personne(item_existant["uuid"], operations)

    resultat = "maj" if operations else "inchange"

    lien_echoue = False
    if orgunit:
        try:
            statut_lien = lier_orgunit(client, item_existant["uuid"], orgunit[0], appliquer)
            if statut_lien == "lie":
                LOG.debug(f"  ↳ lié à l'OrgUnit '{orgunit[1]}' ({nom}, {prenom})")
                resultat = "maj"  # une relation a été ajoutée, même si aucune métadonnée n'a changé
        except Exception:
            lien_echoue = True

    return resultat, lien_echoue, orgunit_introuvable


# =============================================================================
# POINT D'ENTRÉE
# =============================================================================

def analyser_arguments():
    parser = argparse.ArgumentParser(
        description="Synchronise les professeurs extraits avec un dépôt DSpace 7 (crée/met à jour "
                     "les items Person par correspondance de courriel)."
    )
    parser.add_argument("fichier_txt", help="Fichier Synchro (.txt à largeur fixe)")
    parser.add_argument("--excel", required=True, metavar="FICHIER.xlsx",
                         help="Fichier Excel contenant les ORCID (colonnes 'Courriel' et 'ORCID')")
    parser.add_argument("--base-url", default=BASE_URL_PAR_DEFAUT,
                         help="URL de base de l'API REST DSpace")
    parser.add_argument("--user", default=UTILISATEUR_PAR_DEFAUT, help="Utilisateur admin DSpace")
    parser.add_argument("--password", default=MOT_DE_PASSE_PAR_DEFAUT, help="Mot de passe admin DSpace")
    parser.add_argument("--community", default=COMMUNAUTE_PAR_DEFAUT,
                         help="UUID de la communauté où chercher/créer la collection Person")
    parser.add_argument("--collection", help="UUID de la collection cible (évite la résolution "
                                              "automatique via --community)")
    parser.add_argument("--apply", action="store_true",
                         help="Applique réellement les changements (sinon, mode simulation)")
    parser.add_argument("--limit", type=int, metavar="N",
                         help="Ne traite que les N premiers professeurs trouvés (utile pour tester)")
    parser.add_argument("--log-file", metavar="FICHIER.log",
                         help="Chemin du fichier de journal (par défaut : horodaté, dans logs/)")
    parser.add_argument("--verbose", action="store_true",
                         help="Affiche le détail ligne par ligne dans la console (par défaut, "
                              "seul le fichier .log le garde ; la console reste compacte)")
    return parser.parse_args()


def main():
    args = analyser_arguments()

    chemin_log = args.log_file or os.path.join(DOSSIER_LOGS, f"synchro_profs_papyrus_{datetime.now():%Y%m%d_%H%M%S}.log")
    configurer_journalisation(chemin_log, verbose=args.verbose)

    lignes = lire_lignes(args.fichier_txt)
    professeurs_bruts = filtrer_professeurs(lignes)
    professeurs = [extraire_infos(ligne) + [statut] for ligne, statut in professeurs_bruts]
    orcid_par_courriel = charger_orcid_depuis_excel(args.excel)

    if args.limit is not None:
        professeurs = professeurs[:args.limit]

    total = len(professeurs)

    client = DSpaceClient(args.base_url)
    try:
        client.se_connecter(args.user, args.password)
    except Exception as e:
        LOG.error(f"Erreur de connexion à DSpace : {e}")
        sys.exit(1)

    try:
        collection_uuid = args.collection or client.resoudre_collection(args.community)
    except Exception as e:
        LOG.error(f"Erreur lors de la résolution de la collection : {e}")
        sys.exit(1)

    try:
        type_id, gauche, droite = client.resoudre_type_relation(NOM_RELATION_ORGUNIT)
        LOG.debug(f"Type de relation '{NOM_RELATION_ORGUNIT}' résolu : id={type_id}, "
                   f"gauche={gauche}, droite={droite}")
    except Exception as e:
        LOG.warning(f"Impossible de résoudre le type de relation '{NOM_RELATION_ORGUNIT}' : {e}")

    mode = "APPLICATION RÉELLE" if args.apply else "SIMULATION"
    limite = f", limité à {args.limit}" if args.limit is not None else ""
    LOG.info(f"{total} professeur(s){limite} | collection {collection_uuid} | mode {mode}\n")

    compteurs = {"cree": 0, "maj": 0, "inchange": 0, "erreur": 0, "liens_echoues": 0, "orgunit_introuvable": 0}
    heure_depart = time.time()

    for i, professeur in enumerate(professeurs, start=1):
        courriel = professeur[2]
        orcids = orcid_par_courriel.get(normaliser_courriel(courriel), [])
        orcid = orcids[0] if orcids else ""

        try:
            resultat, lien_echoue, orgunit_introuvable = synchroniser_professeur(
                client, collection_uuid, professeur, orcid, args.apply
            )
            compteurs[resultat] += 1
            if lien_echoue:
                compteurs["liens_echoues"] += 1
            if orgunit_introuvable:
                compteurs["orgunit_introuvable"] += 1
        except Exception as e:
            LOG.error(f"[ERREUR] {courriel} : {e}")
            compteurs["erreur"] += 1

        if i % INTERVALLE_PROGRES == 0 or i == total:
            ecoule = time.time() - heure_depart
            rythme = i / ecoule if ecoule > 0 else 0
            restant = (total - i) / rythme if rythme > 0 else 0
            LOG.info(f"{i}/{total} ({100 * i / total:.0f}%) — "
                      f"{compteurs['cree']}c {compteurs['maj']}m {compteurs['inchange']}i "
                      f"{compteurs['erreur']}e — ~{restant / 60:.0f} min restantes")

    courriels_actifs = {normaliser_courriel(p[2]) for p in professeurs}
    try:
        nb_departs = marquer_departs(client, collection_uuid, courriels_actifs, args.apply)
    except Exception as e:
        LOG.warning(f"Impossible de vérifier les départs : {e}")
        nb_departs = 0

    LOG.info(f"\nRésumé : {compteurs['cree']} créé(s), {compteurs['maj']} mis à jour, "
              f"{compteurs['inchange']} inchangé(s), {compteurs['erreur']} erreur(s)")
    if nb_departs:
        verbe = "marqué(s)" if args.apply else "à marquer"
        LOG.info(f"  {nb_departs} prof(s) absent(s) de la liste_personnel {verbe} Inactif "
                  f"(présents dans DSpace mais plus dans le fichier)")
    if compteurs["liens_echoues"]:
        LOG.info(f"  {compteurs['liens_echoues']} liaison(s) OrgUnit échouée(s) (probablement déjà liées)")
    if compteurs["orgunit_introuvable"]:
        LOG.info(f"  {compteurs['orgunit_introuvable']} CodeUnite sans OrgUnit correspondant dans DSpace.")
    if not args.apply:
        LOG.info("Mode simulation : rien n'a été écrit dans DSpace. Relance avec --apply pour appliquer.")


if __name__ == "__main__":
    main()