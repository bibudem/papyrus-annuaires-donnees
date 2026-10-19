#!/usr/bin/env python3
"""
Enrichit les fiches Person de Papyrus (DSpace 7) avec les données publiques
de leur profil ORCID :
  - mots-clés ORCID  -> dc.subject.discipline (une valeur par mot-clé)
  - biographie ORCID -> UdeM.DonneesORCID

Les fiches traitées sont celles de la collection Person qui ont un ORCID
(person.identifier.orcid) : le script ne lit ni liste_personnel ni le fichier
Excel, c'est synchro_profs_papyrus.py qui remplit l'ORCID.

Règles de mise à jour :
  - champ vide dans Papyrus -> rempli avec la valeur ORCID ;
  - champ déjà rempli et différent -> remplacé (ORCID est la seule source pour
    ces deux champs : les mots-clés saisis à la main sont remplacés) ;
  - rien dans ORCID (ou profil privé) -> la fiche n'est pas touchée, rien
    n'est jamais effacé.
Les mots-clés ORCID qui en regroupent plusieurs (« écologie, climat; eau »)
sont découpés sur la virgule, le point-virgule et le retour à la ligne.

Usage:
    python enrichir_orcid_papyrus.py                      # simulation
    python enrichir_orcid_papyrus.py --apply --limit 5    # essai réel sur 5 fiches
    python enrichir_orcid_papyrus.py --apply

Réutilise la connexion DSpace, le .env, le journal et le verrou de
synchro_profs_papyrus.py (les deux scripts ne peuvent donc pas écrire en même
temps). Dépendances : celles de synchro_profs_papyrus.py.
"""

import argparse
import os
import re
import sys
import time
import unicodedata
from datetime import datetime

import requests

from synchro_profs_papyrus import (
    BASE_URL_PAR_DEFAUT,
    CHAMP_COURRIEL,
    CHAMP_NOM,
    CHAMP_ORCID,
    CHAMP_PRENOM,
    CHAMP_STATUT,
    CODE_SORTIE_DEJA_EN_COURS,
    CODE_SORTIE_ERREUR_FATALE,
    CODE_SORTIE_ERREURS_PARTIELLES,
    CODE_SORTIE_OK,
    COLLECTION_PAR_DEFAUT,
    COMMUNAUTE_PAR_DEFAUT,
    CONNEXION_OBLIGATOIRE,
    DELAI_BACKOFF_SECONDES,
    DOSSIER_LOGS,
    FICHIER_VERROU,
    INTERVALLE_PROGRES,
    LOG,
    MAX_ERREURS_CONSECUTIVES,
    MOT_DE_PASSE_PAR_DEFAUT,
    NB_TENTATIVES_MAX,
    TIMEOUT_HTTP,
    UTILISATEUR_PAR_DEFAUT,
    VALEUR_STATUT_INACTIF,
    DSpaceClient,
    acquerir_verrou,
    configurer_journalisation,
    operation_maj_champ,
)

# =============================================================================
# CONFIGURATION
# =============================================================================

# --- Champs de la fiche Person remplis à partir d'ORCID ---
CHAMP_MOTS_CLES = "dc.subject.discipline"
CHAMP_BIOGRAPHIE = "UdeM.DonneesORCID"

# --- API publique ORCID ---
URL_API_ORCID = "https://pub.orcid.org/v3.0"
URL_JETON_ORCID = "https://orcid.org/oauth/token"
# Jeton facultatif (identifiants « Public API » ORCID, dans le .env) : donne
# des limites d'appel plus généreuses. Sans eux, l'accès anonyme suffit.
ORCID_CLIENT_ID = os.environ.get("ORCID_CLIENT_ID") or None
ORCID_CLIENT_SECRET = os.environ.get("ORCID_CLIENT_SECRET") or None
# Pause minimale entre deux appels ORCID, pour rester bien sous les limites de
# l'API (~5 appels/s, soit ~6 min pour 1 700 fiches).
DELAI_ENTRE_APPELS_ORCID = 0.2
# Réponses ORCID qui veulent dire « pas de profil utilisable » (identifiant
# inexistant, profil désactivé/verrouillé, ou fusionné dans un autre) : la
# fiche est sautée et signalée, ce n'est pas une erreur du script.
CODES_ORCID_INUTILISABLE = {404, 409, 410}
CODES_ORCID_A_REESSAYER = {429, 500, 502, 503, 504}

# --- Format des données ---
# Séparateurs utilisés pour découper un mot-clé ORCID qui en regroupe plusieurs.
SEPARATEURS_MOTS_CLES = r"[,;\n]"
MOTIF_ORCID = re.compile(r"\d{4}-\d{4}-\d{4}-\d{3}[\dX]")


# =============================================================================
# CLIENT ORCID
# =============================================================================

class OrcidClient:
    """Lit la section « person » (biographie, mots-clés) d'un profil ORCID
    via l'API publique, avec nouvelles tentatives sur les erreurs passagères."""

    def __init__(self):
        self.session = requests.Session()
        self.session.headers["Accept"] = "application/json"
        self._dernier_appel = 0.0

    def se_connecter(self):
        """Obtient un jeton « /read-public » si des identifiants ORCID sont
        configurés ; sinon l'accès reste anonyme."""
        if not (ORCID_CLIENT_ID and ORCID_CLIENT_SECRET):
            LOG.debug("API ORCID : accès anonyme (ORCID_CLIENT_ID / ORCID_CLIENT_SECRET absents du .env)")
            return
        r = requests.post(
            URL_JETON_ORCID,
            data={"client_id": ORCID_CLIENT_ID, "client_secret": ORCID_CLIENT_SECRET,
                  "grant_type": "client_credentials", "scope": "/read-public"},
            headers={"Accept": "application/json"}, timeout=TIMEOUT_HTTP,
        )
        if r.status_code != 200:
            raise RuntimeError(f"Échec d'obtention du jeton ORCID ({r.status_code}) : {r.text[:300]}")
        self.session.headers["Authorization"] = f"Bearer {r.json()['access_token']}"
        LOG.debug("API ORCID : accès avec jeton /read-public")

    def _patienter(self):
        """Respecte DELAI_ENTRE_APPELS_ORCID depuis l'appel précédent."""
        attente = DELAI_ENTRE_APPELS_ORCID - (time.time() - self._dernier_appel)
        if attente > 0:
            time.sleep(attente)
        self._dernier_appel = time.time()

    def lire_personne(self, orcid):
        """Retourne le JSON de la section « person » du profil, ou None si le
        profil n'est pas utilisable (CODES_ORCID_INUTILISABLE)."""
        url = f"{URL_API_ORCID}/{orcid}/person"
        for tentative in range(1, NB_TENTATIVES_MAX + 1):
            self._patienter()
            delai = DELAI_BACKOFF_SECONDES * (2 ** (tentative - 1))
            try:
                r = self.session.get(url, timeout=TIMEOUT_HTTP)
            except requests.exceptions.RequestException as e:
                if tentative == NB_TENTATIVES_MAX:
                    raise
                LOG.warning(f"    (ORCID injoignable, nouvel essai dans {delai}s : {e})")
                time.sleep(delai)
                continue

            if r.status_code == 200:
                return r.json()
            if r.status_code in CODES_ORCID_INUTILISABLE:
                return None
            if r.status_code in CODES_ORCID_A_REESSAYER and tentative < NB_TENTATIVES_MAX:
                # 429 : ORCID indique lui-même combien de temps attendre.
                retry_after = r.headers.get("Retry-After", "")
                if retry_after.isdigit():
                    delai = max(delai, int(retry_after))
                LOG.warning(f"    (ORCID a répondu {r.status_code}, nouvel essai dans {delai}s)")
                time.sleep(delai)
                continue
            raise RuntimeError(f"Erreur ORCID ({r.status_code}) pour {orcid} : {r.text[:300]}")


# =============================================================================
# EXTRACTION / NETTOYAGE DES DONNÉES ORCID
# =============================================================================

def nettoyer_texte(texte):
    """NFC + espaces superflus retirés (les retours à la ligne sont gardés)."""
    texte = unicodedata.normalize("NFC", texte or "")
    lignes = [" ".join(ligne.split()) for ligne in texte.splitlines()]
    return "\n".join(lignes).strip()


def extraire_mots_cles(personne):
    """Mots-clés publics du profil, découpés sur SEPARATEURS_MOTS_CLES, sans
    doublon (casse ignorée, première orthographe gardée), dans l'ordre ORCID."""
    entrees = ((personne.get("keywords") or {}).get("keyword")) or []
    mots_cles = []
    deja_vus = set()
    for entree in entrees:
        for morceau in re.split(SEPARATEURS_MOTS_CLES, entree.get("content") or ""):
            mot = nettoyer_texte(morceau)
            if mot and mot.casefold() not in deja_vus:
                deja_vus.add(mot.casefold())
                mots_cles.append(mot)
    return mots_cles


def extraire_biographie(personne):
    """Biographie publique du profil, ou '' s'il n'y en a pas."""
    return nettoyer_texte((personne.get("biography") or {}).get("content"))


# =============================================================================
# FICHES PERSON (lecture des métadonnées, sélection, opérations JSON Patch)
# =============================================================================

def valeurs_du_champ(item, champ):
    return [v.get("value", "") for v in item.get("metadata", {}).get(champ, [])]


def orcid_de_la_fiche(item):
    """ORCID de la fiche (accepte aussi la forme https://orcid.org/...), ou ''."""
    for valeur in valeurs_du_champ(item, CHAMP_ORCID):
        m = MOTIF_ORCID.search(valeur.upper())
        if m:
            return m.group(0)
    return ""


def decrire_fiche(item, orcid):
    """Libellé d'une fiche pour le journal : « Nom, Prénom <courriel> ORCID »."""
    nom = item.get("name") or ", ".join(valeurs_du_champ(item, CHAMP_NOM)
                                         + valeurs_du_champ(item, CHAMP_PRENOM))
    courriels = valeurs_du_champ(item, CHAMP_COURRIEL)
    return f"{nom or 'sans nom'} <{courriels[0] if courriels else 'sans courriel'}> {orcid}"


def selectionner_fiches(items, inclure_inactifs):
    """Garde les fiches qui ont un ORCID (et qui ne sont pas Inactif, sauf
    inclure_inactifs). Retourne (fiches, nb_sans_orcid, nb_inactifs), où
    fiches est une liste de (item, orcid)."""
    fiches = []
    nb_sans_orcid = nb_inactifs = 0
    for item in items:
        orcid = orcid_de_la_fiche(item)
        if not orcid:
            nb_sans_orcid += 1
        elif not inclure_inactifs and VALEUR_STATUT_INACTIF in valeurs_du_champ(item, CHAMP_STATUT):
            nb_inactifs += 1
        else:
            fiches.append((item, orcid))
    return fiches, nb_sans_orcid, nb_inactifs


def operations_enrichissement(item, mots_cles, biographie):
    """Opérations JSON Patch pour aligner la fiche sur ORCID. Un champ dont
    ORCID n'a pas de valeur n'est pas touché."""
    operations = []
    if mots_cles:
        actuels = valeurs_du_champ(item, CHAMP_MOTS_CLES)
        if actuels != mots_cles:
            # Retrait puis ajouts dans le même PATCH : toutes les anciennes
            # valeurs sont remplacées d'un coup, quel que soit leur nombre.
            # Un seul « add » avec une liste de valeurs n'en garde que la
            # première dans DSpace : chaque mot-clé suivant est donc ajouté à
            # la fin (/-) par sa propre opération.
            chemin = f"/metadata/{CHAMP_MOTS_CLES}"
            if actuels:
                operations.append({"op": "remove", "path": chemin})
            operations.append({"op": "add", "path": chemin, "value": [{"value": mots_cles[0]}]})
            operations += [{"op": "add", "path": f"{chemin}/-", "value": {"value": mot}}
                           for mot in mots_cles[1:]]
    if biographie:
        operation = operation_maj_champ(item, CHAMP_BIOGRAPHIE, biographie)
        if operation:
            operations.append(operation)
    return operations


# =============================================================================
# VÉRIFICATIONS DSPACE
# =============================================================================

def champ_existe(client, champ):
    """True si le champ est déclaré dans le registre des métadonnées DSpace,
    False s'il n'y est pas, None si la vérification elle-même a échoué."""
    r = client._requete("GET", f"{client.base_url}/core/metadatafields/search/byFieldName",
                        params={"exactName": champ})
    if r.status_code != 200:
        return None
    return bool(r.json().get("_embedded", {}).get("metadatafields"))


def verifier_champs(client):
    """Vérifie que les champs cibles existent dans DSpace. Un champ absent
    ferait échouer chaque PATCH (422) : mieux vaut s'arrêter tout de suite
    que d'accumuler des milliers d'erreurs. Retourne False s'il faut arrêter."""
    for champ in (CHAMP_MOTS_CLES, CHAMP_BIOGRAPHIE):
        existe = champ_existe(client, champ)
        if existe is False:
            LOG.error(f"Le champ '{champ}' n'existe pas dans le registre des métadonnées de DSpace — "
                      f"l'ajouter avant de relancer. Rien n'a été fait.")
            return False
        if existe is None:
            LOG.warning(f"⚠ Impossible de vérifier que le champ '{champ}' existe dans DSpace — on continue.")
    return True


# =============================================================================
# ENRICHISSEMENT D'UNE FICHE
# =============================================================================

def enrichir_fiche(client, orcid_client, item, orcid, appliquer):
    """Lit le profil ORCID et met la fiche à jour si besoin. Retourne
    'maj', 'inchange', 'vide' (rien de public dans ORCID) ou 'inutilisable'
    (ORCID introuvable, désactivé ou fusionné)."""
    personne = orcid_client.lire_personne(orcid)
    if personne is None:
        LOG.debug(f"[ORCID INUTILISABLE] {decrire_fiche(item, orcid)} (introuvable, désactivé ou fusionné)")
        return "inutilisable"

    mots_cles = extraire_mots_cles(personne)
    biographie = extraire_biographie(personne)
    if not (mots_cles or biographie):
        return "vide"

    operations = operations_enrichissement(item, mots_cles, biographie)
    if not operations:
        return "inchange"

    LOG.debug(f"[MISE À JOUR] {decrire_fiche(item, orcid)} — {len(mots_cles)} mot(s)-clé(s), "
              f"biographie {'oui' if biographie else 'non'}")
    if mots_cles:
        LOG.debug(f"  ↳ mots-clés : {' | '.join(mots_cles)}")
    if appliquer:
        client.mettre_a_jour_personne(item["uuid"], operations)
    return "maj"


# =============================================================================
# PROGRAMME PRINCIPAL
# =============================================================================

def analyser_arguments():
    parser = argparse.ArgumentParser(
        description=f"Remplit {CHAMP_MOTS_CLES} (mots-clés) et {CHAMP_BIOGRAPHIE} (biographie) des "
                    f"fiches Person à partir de leur profil ORCID public."
    )
    parser.add_argument("--base-url", help="URL de base de l'API REST DSpace (défaut : DSPACE_BASE_URL du .env)")
    parser.add_argument("--user", help="Utilisateur admin DSpace (défaut : DSPACE_USER du .env)")
    parser.add_argument("--password", help="Mot de passe admin DSpace (défaut : DSPACE_PASSWORD du .env)")
    parser.add_argument("--community", default=COMMUNAUTE_PAR_DEFAUT,
                        help="UUID de la communauté où chercher la collection Person")
    parser.add_argument("--collection", default=COLLECTION_PAR_DEFAUT,
                        help="UUID de la collection Person (défaut : DSPACE_COLLECTION du .env)")
    parser.add_argument("--apply", action="store_true",
                        help="Applique réellement les changements (sinon, mode simulation)")
    parser.add_argument("--limit", type=int, metavar="N",
                        help="Ne traite que les N premières fiches avec ORCID (utile pour tester)")
    parser.add_argument("--inclure-inactifs", action="store_true",
                        help=f"Traite aussi les fiches {CHAMP_STATUT} = {VALEUR_STATUT_INACTIF} "
                             f"(sautées par défaut)")
    parser.add_argument("--log-file", metavar="FICHIER.log",
                        help="Chemin du fichier de journal (par défaut : horodaté, dans logs/)")
    parser.add_argument("--verbose", action="store_true",
                        help="Affiche le détail fiche par fiche dans la console")
    args = parser.parse_args()

    args.connexion_par_defaut = [variable for option, variable in CONNEXION_OBLIGATOIRE.items()
                                 if getattr(args, option) is None and variable not in os.environ]
    args.base_url = args.base_url or BASE_URL_PAR_DEFAUT
    args.user = args.user or UTILISATEUR_PAR_DEFAUT
    args.password = args.password or MOT_DE_PASSE_PAR_DEFAUT
    return args


def afficher_resume(compteurs, orcid_inutilisables, appliquer):
    verbe = "mise(s) à jour" if appliquer else "à mettre à jour"
    LOG.info(f"\nRésumé : {compteurs['maj']} fiche(s) {verbe}, {compteurs['inchange']} déjà à jour, "
             f"{compteurs['vide']} sans mots-clés ni biographie publics, {compteurs['erreur']} erreur(s)")
    if orcid_inutilisables:
        LOG.info(f"  {len(orcid_inutilisables)} ORCID introuvable(s), désactivé(s) ou fusionné(s) "
                 f"— à vérifier dans les fiches :")
        for qui in orcid_inutilisables:
            LOG.info(f"    {qui}")
    if not appliquer:
        LOG.info("Mode simulation : rien n'a été écrit dans DSpace. Relance avec --apply pour appliquer.")


def executer(args):
    """Déroule tout l'enrichissement et retourne un code de sortie (CODE_SORTIE_*)."""
    if args.apply and args.connexion_par_defaut:
        LOG.error(f"Mode --apply refusé : {', '.join(args.connexion_par_defaut)} absent(s) du .env "
                  f"et de la ligne de commande.")
        return CODE_SORTIE_ERREUR_FATALE

    client = DSpaceClient(args.base_url)
    orcid_client = OrcidClient()
    try:
        client.se_connecter(args.user, args.password)
        collection_uuid = args.collection or client.resoudre_collection(args.community)
        orcid_client.se_connecter()
    except Exception as e:
        LOG.error(f"Erreur de connexion : {e}")
        return CODE_SORTIE_ERREUR_FATALE

    if not verifier_champs(client):
        return CODE_SORTIE_ERREUR_FATALE

    items = client.lister_items_de_la_collection(collection_uuid)
    fiches, nb_sans_orcid, nb_inactifs = selectionner_fiches(items, args.inclure_inactifs)
    if args.limit is not None:
        fiches = fiches[:args.limit]
    total = len(fiches)

    mode = "APPLICATION RÉELLE" if args.apply else "SIMULATION"
    limite = f", limité à {args.limit}" if args.limit is not None else ""
    LOG.info(f"{len(items)} fiche(s) dans la collection : {total} à traiter{limite} "
             f"({nb_sans_orcid} sans ORCID, {nb_inactifs} Inactif sautée(s)) | mode {mode}\n")

    compteurs = {"maj": 0, "inchange": 0, "vide": 0, "erreur": 0}
    orcid_inutilisables = []
    heure_depart = time.time()
    erreurs_consecutives = 0
    arret_erreurs = False

    for i, (item, orcid) in enumerate(fiches, start=1):
        try:
            resultat = enrichir_fiche(client, orcid_client, item, orcid, args.apply)
            if resultat == "inutilisable":
                orcid_inutilisables.append(decrire_fiche(item, orcid))
            else:
                compteurs[resultat] += 1
            erreurs_consecutives = 0
        except Exception as e:
            LOG.error(f"[ERREUR] {decrire_fiche(item, orcid)} : {e}")
            compteurs["erreur"] += 1
            erreurs_consecutives += 1

        if i % INTERVALLE_PROGRES == 0 or i == total:
            ecoule = time.time() - heure_depart
            rythme = i / ecoule if ecoule > 0 else 0
            restant = (total - i) / rythme if rythme > 0 else 0
            LOG.info(f"{i}/{total} ({100 * i / total:.0f}%) — {compteurs['maj']}m "
                     f"{compteurs['inchange']}i {compteurs['vide']}v {compteurs['erreur']}e "
                     f"— ~{restant / 60:.0f} min restantes")

        if erreurs_consecutives >= MAX_ERREURS_CONSECUTIVES:
            LOG.error(f"ARRÊT : {erreurs_consecutives} erreurs d'affilée — DSpace ou ORCID semble "
                      f"indisponible. Abandon après {i}/{total} fiche(s) (le run suivant reprendra, "
                      f"rien n'est perdu). Dernière erreur ci-dessus.")
            arret_erreurs = True
            break

    afficher_resume(compteurs, orcid_inutilisables, args.apply)

    if arret_erreurs:
        return CODE_SORTIE_ERREUR_FATALE
    if compteurs["erreur"]:
        return CODE_SORTIE_ERREURS_PARTIELLES
    return CODE_SORTIE_OK


def main():
    args = analyser_arguments()

    chemin_log = args.log_file or os.path.join(
        DOSSIER_LOGS, f"enrichir_orcid_papyrus_{datetime.now():%Y%m%d_%H%M%S}.log")
    configurer_journalisation(chemin_log, verbose=args.verbose)

    # Même verrou que synchro_profs_papyrus.py : les deux scripts modifient les
    # mêmes fiches et ne doivent pas écrire en même temps.
    verrou = None
    if args.apply:
        verrou = acquerir_verrou(FICHIER_VERROU)
        if verrou is None:
            LOG.error(f"Une synchronisation est déjà en cours (verrou : {FICHIER_VERROU}) — abandon.")
            sys.exit(CODE_SORTIE_DEJA_EN_COURS)

    try:
        code = executer(args)
    except Exception:
        LOG.exception("Erreur inattendue — enrichissement interrompu :")
        code = CODE_SORTIE_ERREUR_FATALE
    finally:
        if verrou:
            verrou.close()

    if code != CODE_SORTIE_OK:
        LOG.info(f"Code de sortie : {code} (détails dans {chemin_log})")
    sys.exit(code)


if __name__ == "__main__":
    main()
