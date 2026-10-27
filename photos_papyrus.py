#!/usr/bin/env python3
"""
Met à jour la photo des fiches Person de Papyrus (DSpace 9+) à partir de
l'archive des chercheurs (data/chercheurs.zip) :

    chercheurs/<codesim>/info.txt                   nom, prenom, courriel, codesim, idsadvr
    chercheurs/<codesim>/photoOfficielle.jpg        photo envoyée dans Papyrus
    chercheurs/<codesim>/photoOfficiellePetite.jpg  ignorée (trop petite)

Chaque dossier est rattaché à sa fiche par le courriel de info.txt. Sans
courriel, ou sans fiche avec ce courriel, le dossier est sauté et signalé.

Règles de mise à jour :
  - fiche sans photo -> la photo officielle est ajoutée ;
  - fiche avec la même photo (même empreinte MD5) -> rien n'est envoyé ;
  - fiche avec une autre image (ancienne photo, ou image ajoutée à la main)
    -> remplacée par la photo officielle.
La photo va dans le bundle ORIGINAL, avec une copie dans THUMBNAIL : elle
s'affiche tout de suite, sans attendre la tâche DSpace filter-media.

Usage:
    python photos_papyrus.py                      # simulation
    python photos_papyrus.py --apply --limit 5    # essai réel sur 5 fiches
    python photos_papyrus.py --apply

Réutilise la connexion DSpace, le .env, le journal et le verrou de
synchro_profs_papyrus.py (les scripts ne peuvent donc pas écrire en même temps).
"""

import argparse
import csv
import hashlib
import json
import os
import posixpath
import sys
import time
import unicodedata
import zipfile
from dataclasses import dataclass
from datetime import datetime

from synchro_profs_papyrus import (
    BASE_URL_PAR_DEFAUT,
    CHAMP_COURRIEL,
    CHAMP_NOM,
    CHAMP_PRENOM,
    CHAMP_STATUT,
    CODES_HTTP_SUCCES,
    CODE_SORTIE_DEJA_EN_COURS,
    CODE_SORTIE_ERREUR_FATALE,
    CODE_SORTIE_ERREURS_PARTIELLES,
    CODE_SORTIE_OK,
    COLLECTION_PAR_DEFAUT,
    COMMUNAUTE_PAR_DEFAUT,
    CONNEXION_OBLIGATOIRE,
    DOSSIER_LOGS,
    DOSSIER_SCRIPT,
    FICHIER_VERROU,
    INTERVALLE_PROGRES,
    LOG,
    MAX_ERREURS_CONSECUTIVES,
    MOT_DE_PASSE_PAR_DEFAUT,
    UTILISATEUR_PAR_DEFAUT,
    VALEUR_STATUT_INACTIF,
    DSpaceClient,
    acquerir_verrou,
    configurer_journalisation,
    resoudre_fichier_entree,
)
# Importé après synchro_profs_papyrus, qui ajoute outils/ au chemin des modules.
from extraire_professeurs import CSV_ENCODAGE, CSV_SEPARATEUR, normaliser_courriel

# =============================================================================
# CONFIGURATION
# =============================================================================

# --- Archive des chercheurs (FICHIER_PHOTOS du .env : chemin ou motif avec *) ---
MOTIF_FICHIER_PHOTOS = "chercheurs*.zip"
FICHIER_PHOTOS_PAR_DEFAUT = os.environ.get("FICHIER_PHOTOS") or os.path.join("data", MOTIF_FICHIER_PHOTOS)
NOM_FICHIER_INFO = "info.txt"
NOM_PHOTO = "photoOfficielle.jpg"

# --- Bundles DSpace ---
BUNDLE_ORIGINAL = "ORIGINAL"
BUNDLE_VIGNETTE = "THUMBNAIL"
# DSpace associe une vignette à son original par le nom « <original>.jpg ».
NOM_VIGNETTE = f"{NOM_PHOTO}.jpg"
TYPE_MIME_PHOTO = "image/jpeg"
# Seules les images du bundle ORIGINAL sont remplacées : un autre fichier
# (ex. un PDF ajouté à la main) n'est jamais touché.
EXTENSIONS_IMAGE = (".jpg", ".jpeg", ".png", ".gif", ".webp", ".bmp", ".tif", ".tiff")

# --- Rapport : motifs de saut, dans l'ordre d'affichage ---
# affichage : "liste" (total + un dossier par ligne), "total" (total seulement),
# "toujours" (liste, et ligne affichée même à 0).
MOTIFS_SAUT = {
    "sans_courriel_fiche": ("dossier(s) sans courriel dans info.txt, mais avec une fiche au même nom "
                            "dans Papyrus (photo non mise, courriel à ajouter dans info.txt)", "toujours"),
    "sans_courriel": ("dossier(s) sans courriel dans info.txt et sans fiche au même nom dans Papyrus",
                      "total"),
    "sans_photo": (f"dossier(s) sans {NOM_PHOTO}", "liste"),
    "courriel_incoherent": ("dossier(s) dont le courriel diffère de celui de la fiche au même nom "
                            "(photo non mise, courriel à corriger)", "liste"),
    "fiche_introuvable": ("dossier(s) sans fiche dans Papyrus (aucune fiche avec ce courriel ni ce nom)",
                          "total"),
    "plusieurs_fiches": ("dossier(s) dont le courriel est sur plusieurs fiches (doublons à corriger)",
                         "liste"),
    "inactif": ("fiche(s) Inactif sautée(s) (voir --inclure-inactifs)", "total"),
}


# =============================================================================
# ARCHIVE
# =============================================================================

@dataclass
class Dossier:
    """Un dossier de l'archive : chercheurs/<code>/."""
    code: str
    nom: str
    prenom: str
    courriel: str          # normalisé, '' si absent de info.txt
    idsadvr: str
    chemin_photo: str      # chemin dans l'archive, '' si la photo manque

    def decrire(self, item=None):
        """« code — Nom, Prénom <courriel> » (nom de la fiche si elle est connue)."""
        nom = nom_de_la_fiche(item) if item else f"{self.nom}, {self.prenom}"
        return f"{self.code} — {nom} <{self.courriel or 'sans courriel'}>"


def lire_info(texte):
    """Dict des lignes « cle=valeur » d'un info.txt."""
    info = {}
    for ligne in texte.splitlines():
        cle, egal, valeur = ligne.partition("=")
        if egal:
            info[cle.strip()] = valeur.strip()
    return info


def lire_archive(archive):
    """Liste des dossiers de l'archive (ZipFile ouvert) qui ont un info.txt.
    Les photos ne sont pas lues ici, seulement au moment de les envoyer."""
    noms = set(archive.namelist())
    dossiers = []
    for nom in sorted(n for n in noms if posixpath.basename(n) == NOM_FICHIER_INFO):
        repertoire = posixpath.dirname(nom)
        info = lire_info(archive.read(nom).decode("utf-8-sig", errors="replace"))
        chemin_photo = posixpath.join(repertoire, NOM_PHOTO)
        dossiers.append(Dossier(
            code=posixpath.basename(repertoire),
            nom=info.get("nom", ""),
            prenom=info.get("prenom", ""),
            courriel=normaliser_courriel(info.get("courriel") or ""),
            idsadvr=info.get("idsadvr", ""),
            chemin_photo=chemin_photo if chemin_photo in noms else "",
        ))
    return dossiers


def sans_accents(texte):
    """Forme de comparaison souple : minuscules, sans accents ni espaces superflus."""
    decompose = unicodedata.normalize("NFKD", texte or "")
    return " ".join("".join(c for c in decompose if not unicodedata.combining(c)).casefold().split())


def cle_nom(nom, prenom):
    return sans_accents(nom), sans_accents(prenom)


# =============================================================================
# FICHES PERSON ET RATTACHEMENT
# =============================================================================

def valeurs_du_champ(item, champ):
    return [v.get("value", "") for v in item.get("metadata", {}).get(champ, [])]


def indexer_fiches(items):
    """Retourne (par_courriel, par_nom) : courriel normalisé -> fiches, et
    (nom, prénom) -> fiches."""
    par_courriel, par_nom = {}, {}
    for item in items:
        for courriel in filter(None, valeurs_du_champ(item, CHAMP_COURRIEL)):
            par_courriel.setdefault(normaliser_courriel(courriel), []).append(item)
        for nom in valeurs_du_champ(item, CHAMP_NOM):
            for prenom in valeurs_du_champ(item, CHAMP_PRENOM):
                par_nom.setdefault(cle_nom(nom, prenom), []).append(item)
    return par_courriel, par_nom


def decrire_fiche(item):
    """« courriel (statut, uuid) » d'une fiche, pour le journal."""
    courriels = ", ".join(valeurs_du_champ(item, CHAMP_COURRIEL)) or "sans courriel"
    statut = ", ".join(valeurs_du_champ(item, CHAMP_STATUT)) or "statut inconnu"
    return f"{courriels} ({statut}, {item.get('uuid')})"


def nom_de_la_fiche(item):
    """« Nom, Prénom » tirés des champs de la fiche (le titre est souvent vide)."""
    nom = " ".join(valeurs_du_champ(item, CHAMP_NOM))
    prenom = " ".join(valeurs_du_champ(item, CHAMP_PRENOM))
    return f"{nom}, {prenom}" if nom or prenom else (item.get("name") or "sans nom")


def mots(texte):
    """Mots d'un nom, sans accents ni ponctuation, initiales (« L. ») exclues."""
    return {mot for mot in sans_accents(texte).replace("-", " ").replace(".", " ").split() if len(mot) > 1}


def nom_different(item, dossier):
    """True si le nom ou le prénom de info.txt n'est pas compatible avec la
    fiche. Compatible = les mots de l'un sont tous dans l'autre : « Borges »
    et « Ferreira Borges », ou « Adrian L. » et « Adrian », ne sont pas signalés."""
    for champ, attendu in ((CHAMP_NOM, dossier.nom), (CHAMP_PRENOM, dossier.prenom)):
        a, b = mots(attendu), mots(" ".join(valeurs_du_champ(item, champ)))
        if a and b and not (a <= b or b <= a):
            return True
    return False


def apparier(dossiers, par_courriel, par_nom, inclure_inactifs):
    """Rattache chaque dossier à sa fiche par le courriel. Retourne
    (a_traiter, sautes) : a_traiter liste les (dossier, fiche) ; sautes
    associe chaque motif de MOTIFS_SAUT à ses (dossier, ligne de rapport).
    Le nom ne sert jamais à rattacher, seulement à signaler une incohérence."""
    # Une fiche déjà rattachée par le courriel d'un autre dossier a sa photo :
    # un homonyme (ex. deux « Lavoie, Marc ») ne doit pas être signalé contre elle.
    deja_rattachees = {item["uuid"] for d in dossiers if d.courriel
                       for item in par_courriel.get(d.courriel, [])}
    a_traiter = []
    sautes = {motif: [] for motif in MOTIFS_SAUT}
    for dossier in dossiers:
        fiches = par_courriel.get(dossier.courriel, []) if dossier.courriel else []
        homonymes = [item for item in par_nom.get(cle_nom(dossier.nom, dossier.prenom), [])
                     if item["uuid"] not in deja_rattachees]
        if not dossier.courriel:
            if homonymes:
                fiches_trouvees = " | ".join(decrire_fiche(item) for item in homonymes)
                sautes["sans_courriel_fiche"].append(
                    (dossier, f"{dossier.decrire()} — fiche(s) au même nom : {fiches_trouvees}"))
            else:
                sautes["sans_courriel"].append((dossier, dossier.decrire()))
        elif not dossier.chemin_photo:
            sautes["sans_photo"].append((dossier, dossier.decrire()))
        elif not fiches:
            courriels = sorted({c for item in homonymes for c in valeurs_du_champ(item, CHAMP_COURRIEL)})
            if courriels:
                sautes["courriel_incoherent"].append(
                    (dossier, f"{dossier.decrire()} — courriel de la fiche : {', '.join(courriels)}"))
            else:
                sautes["fiche_introuvable"].append((dossier, dossier.decrire()))
        elif len(fiches) > 1:
            sautes["plusieurs_fiches"].append((dossier, dossier.decrire()))
        elif not inclure_inactifs and VALEUR_STATUT_INACTIF in valeurs_du_champ(fiches[0], CHAMP_STATUT):
            sautes["inactif"].append((dossier, dossier.decrire(fiches[0])))
        else:
            a_traiter.append((dossier, fiches[0]))
    return a_traiter, sautes


# =============================================================================
# PHOTOS DANS DSPACE (bundles et fichiers)
# =============================================================================

def lire_bundles(client, item_uuid):
    """nom du bundle -> (uuid du bundle, liste de ses fichiers)."""
    r = client._requete("GET", f"{client.base_url}/core/items/{item_uuid}/bundles",
                        params={"embed": "bitstreams", "size": 100})
    r.raise_for_status()
    return {bundle["name"]: (bundle["uuid"],
                             bundle.get("_embedded", {}).get("bitstreams", {}).get("_embedded", {})
                                   .get("bitstreams", []))
            for bundle in r.json().get("_embedded", {}).get("bundles", [])}


def fichiers_du_bundle(bundles, nom):
    return bundles.get(nom, (None, []))[1]


def images_originales(bundles):
    return [f for f in fichiers_du_bundle(bundles, BUNDLE_ORIGINAL)
            if (f.get("name") or "").lower().endswith(EXTENSIONS_IMAGE)]


def md5_du_fichier(fichier):
    somme = fichier.get("checkSum") or {}
    return (somme.get("value") or "").lower() if somme.get("checkSumAlgorithm") == "MD5" else ""


def photo_a_jour(bundles, md5):
    """True si la fiche porte exactement la photo officielle et sa vignette,
    sans autre image."""
    images = images_originales(bundles)
    vignettes = fichiers_du_bundle(bundles, BUNDLE_VIGNETTE)
    return (len(images) == 1 and len(vignettes) == 1
            and images[0].get("name") == NOM_PHOTO and md5_du_fichier(images[0]) == md5
            and vignettes[0].get("name") == NOM_VIGNETTE and md5_du_fichier(vignettes[0]) == md5)


def creer_bundle(client, item_uuid, nom):
    r = client._requete("POST", f"{client.base_url}/core/items/{item_uuid}/bundles",
                        json={"name": nom, "metadata": {}},
                        headers={"Content-Type": "application/json"}, csrf=True)
    if r.status_code not in CODES_HTTP_SUCCES:
        raise RuntimeError(f"Échec de création du bundle {nom} ({r.status_code}) : {r.text[:300]}")
    return r.json()["uuid"]


def envoyer_fichier(client, bundle_uuid, nom, contenu):
    r = client._requete("POST", f"{client.base_url}/core/bundles/{bundle_uuid}/bitstreams",
                        files={"file": (nom, contenu, TYPE_MIME_PHOTO)},
                        data={"properties": json.dumps({"name": nom})}, csrf=True)
    if r.status_code not in CODES_HTTP_SUCCES:
        raise RuntimeError(f"Échec d'envoi de {nom} ({r.status_code}) : {r.text[:300]}")


def supprimer_fichier(client, fichier):
    r = client._requete("DELETE", f"{client.base_url}/core/bitstreams/{fichier['uuid']}", csrf=True)
    if r.status_code not in (200, 204):
        raise RuntimeError(f"Échec de suppression de {fichier.get('name')} ({r.status_code}) : {r.text[:300]}")


def remplacer_photo(client, item_uuid, bundles, photo):
    """Envoie la photo (ORIGINAL) et sa vignette (THUMBNAIL), puis retire les
    anciennes. Dans cet ordre, une fiche n'est jamais laissée sans photo."""
    anciens = images_originales(bundles) + fichiers_du_bundle(bundles, BUNDLE_VIGNETTE)
    for nom_bundle, nom_fichier in ((BUNDLE_ORIGINAL, NOM_PHOTO), (BUNDLE_VIGNETTE, NOM_VIGNETTE)):
        bundle_uuid = bundles.get(nom_bundle, (None,))[0] or creer_bundle(client, item_uuid, nom_bundle)
        envoyer_fichier(client, bundle_uuid, nom_fichier, photo)
    for fichier in anciens:
        supprimer_fichier(client, fichier)


def traiter_dossier(client, archive, dossier, item, appliquer):
    """Met la photo du dossier sur la fiche si besoin. Retourne 'ajoutee',
    'remplacee' ou 'a_jour'."""
    photo = archive.read(dossier.chemin_photo)
    bundles = lire_bundles(client, item["uuid"])
    if photo_a_jour(bundles, hashlib.md5(photo).hexdigest()):
        return "a_jour"

    anciennes = [f["name"] for f in images_originales(bundles)]
    if anciennes:
        LOG.debug(f"[REMPLACEMENT] {dossier.decrire(item)} — ancienne(s) image(s) : {', '.join(anciennes)}")
    else:
        LOG.debug(f"[AJOUT] {dossier.decrire(item)}")
    if appliquer:
        remplacer_photo(client, item["uuid"], bundles, photo)
    return "remplacee" if anciennes else "ajoutee"


# =============================================================================
# RAPPORT
# =============================================================================

def afficher_resume(compteurs, sautes, noms_differents, appliquer):
    verbe = "" if appliquer else " (à faire)"
    LOG.info(f"\nRésumé{verbe} : {compteurs['ajoutee']} photo(s) ajoutée(s), {compteurs['remplacee']} "
             f"remplacée(s), {compteurs['a_jour']} déjà à jour, {compteurs['erreur']} erreur(s)")
    for motif, (libelle, affichage) in MOTIFS_SAUT.items():
        lignes = sautes[motif]
        if not lignes and affichage != "toujours":
            continue
        LOG.info(f"  {len(lignes)} {libelle}")
        if affichage != "total":
            for _, ligne in lignes:
                LOG.info(f"    {ligne}")
    if noms_differents:
        LOG.info(f"  {len(noms_differents)} fiche(s) au nom ou prénom différent de info.txt "
                 f"(photo mise quand même, le courriel correspond) — à vérifier :")
        for ligne in noms_differents:
            LOG.info(f"    {ligne}")
    if not appliquer:
        LOG.info("Mode simulation : rien n'a été écrit dans DSpace. Relance avec --apply pour appliquer.")


def afficher_bilan(compteurs, nb_dossiers, nb_rattaches, nb_fiches, appliquer, limite):
    """Combien de dossiers de photos ont mis à jour combien de fiches Papyrus."""
    nb_maj = compteurs["ajoutee"] + compteurs["remplacee"]
    verbe = "mise(s) à jour" if appliquer else "à mettre à jour"
    LOG.info(f"\nBILAN : {nb_dossiers} dossier(s) de photos dans l'archive, {nb_rattaches} rattaché(s) "
             f"à une fiche Papyrus par le courriel -> {nb_maj} fiche(s) Papyrus {verbe} "
             f"({compteurs['ajoutee']} ajout(s), {compteurs['remplacee']} remplacement(s)), "
             f"{compteurs['a_jour']} déjà à jour, {compteurs['erreur']} erreur(s).")
    # Avec --limit, toutes les fiches n'ont pas été vues : ce total serait faux.
    if limite is None:
        avec_photo = compteurs["a_jour"] + (nb_maj if appliquer else 0)
        LOG.info(f"        {avec_photo} fiche(s) sur {nb_fiches} dans la collection ont la photo officielle"
                 f"{'' if appliquer else ' (avant application)'}.")


def extraire_sans_fiche(dossiers):
    """Écrit les dossiers sans fiche dans Papyrus dans un CSV horodaté (logs/)."""
    chemin = os.path.join(DOSSIER_LOGS, f"photos_sans_fiche_{datetime.now():%Y%m%d_%H%M%S}.csv")
    with open(chemin, "w", newline="", encoding=CSV_ENCODAGE) as f:
        writer = csv.writer(f, delimiter=CSV_SEPARATEUR)
        writer.writerow(["codesim", "nom", "prenom", "courriel", "idsadvr"])
        writer.writerows([d.code, d.nom, d.prenom, d.courriel, d.idsadvr] for d in dossiers)
    LOG.info(f"Extrait des {len(dossiers)} dossier(s) sans fiche dans Papyrus : {chemin}")


# =============================================================================
# PROGRAMME PRINCIPAL
# =============================================================================

def analyser_arguments():
    parser = argparse.ArgumentParser(
        description="Met à jour la photo des fiches Person à partir de l'archive des chercheurs "
                    "(un dossier par personne : info.txt + photoOfficielle.jpg)."
    )
    parser.add_argument("fichier_zip", nargs="?",
                        help="Archive des chercheurs (.zip) ou motif avec * (la plus récente est "
                             "choisie). Défaut : FICHIER_PHOTOS du .env, sinon data/chercheurs*.zip")
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
                        help="Ne traite que les N premières fiches trouvées (utile pour tester)")
    parser.add_argument("--inclure-inactifs", action="store_true",
                        help=f"Traite aussi les fiches {CHAMP_STATUT} = {VALEUR_STATUT_INACTIF} "
                             f"(sautées par défaut)")
    parser.add_argument("--csv-logs", action="store_true",
                        help="Écrit aussi les dossiers sans fiche dans Papyrus dans "
                             "logs/photos_sans_fiche_<date>.csv (codesim, nom, prénom, courriel, idsadvr)")
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
    args.fichier_zip = args.fichier_zip or os.path.join(DOSSIER_SCRIPT, FICHIER_PHOTOS_PAR_DEFAUT)
    return args


def executer(args):
    """Déroule toute la mise à jour des photos et retourne un code de sortie (CODE_SORTIE_*)."""
    if args.apply and args.connexion_par_defaut:
        LOG.error(f"Mode --apply refusé : {', '.join(args.connexion_par_defaut)} absent(s) du .env "
                  f"et de la ligne de commande.")
        return CODE_SORTIE_ERREUR_FATALE

    chemin_zip = resoudre_fichier_entree(args.fichier_zip, MOTIF_FICHIER_PHOTOS)
    with zipfile.ZipFile(chemin_zip) as archive:
        dossiers = lire_archive(archive)
        LOG.info(f"Archive : {chemin_zip} ({len(dossiers)} dossier(s))")

        client = DSpaceClient(args.base_url)
        try:
            client.se_connecter(args.user, args.password)
            collection_uuid = args.collection or client.resoudre_collection(args.community)
        except Exception as e:
            LOG.error(f"Erreur de connexion à DSpace : {e}")
            return CODE_SORTIE_ERREUR_FATALE

        items = client.lister_items_de_la_collection(collection_uuid)
        a_traiter, sautes = apparier(dossiers, *indexer_fiches(items), args.inclure_inactifs)
        # Compté avant --limit : dossiers dont le courriel a trouvé sa fiche.
        nb_rattaches = len(a_traiter) + len(sautes["inactif"])
        if args.limit is not None:
            a_traiter = a_traiter[:args.limit]
        total = len(a_traiter)

        mode = "APPLICATION RÉELLE" if args.apply else "SIMULATION"
        limite = f", limité à {args.limit}" if args.limit is not None else ""
        LOG.info(f"{total} fiche(s) à traiter{limite} | collection {collection_uuid} | mode {mode}\n")

        compteurs = {"ajoutee": 0, "remplacee": 0, "a_jour": 0, "erreur": 0}
        noms_differents = []
        heure_depart = time.time()
        erreurs_consecutives = 0
        arret_erreurs = False

        for i, (dossier, item) in enumerate(a_traiter, start=1):
            if nom_different(item, dossier):
                noms_differents.append(f"{dossier.decrire(item)} — info.txt : {dossier.nom}, {dossier.prenom}")
            try:
                compteurs[traiter_dossier(client, archive, dossier, item, args.apply)] += 1
                erreurs_consecutives = 0
            except Exception as e:
                LOG.error(f"[ERREUR] {dossier.decrire(item)} : {e}")
                compteurs["erreur"] += 1
                erreurs_consecutives += 1

            if i % INTERVALLE_PROGRES == 0 or i == total:
                ecoule = time.time() - heure_depart
                rythme = i / ecoule if ecoule > 0 else 0
                restant = (total - i) / rythme if rythme > 0 else 0
                LOG.info(f"{i}/{total} ({100 * i / total:.0f}%) — {compteurs['ajoutee']}a "
                         f"{compteurs['remplacee']}r {compteurs['a_jour']}i {compteurs['erreur']}e "
                         f"— ~{restant / 60:.0f} min restantes")

            if erreurs_consecutives >= MAX_ERREURS_CONSECUTIVES:
                LOG.error(f"ARRÊT : {erreurs_consecutives} erreurs d'affilée — DSpace semble indisponible. "
                          f"Abandon après {i}/{total} fiche(s) (le run suivant reprendra, rien n'est "
                          f"perdu). Dernière erreur ci-dessus.")
                arret_erreurs = True
                break

    if args.csv_logs:
        extraire_sans_fiche([dossier for dossier, _ in sautes["fiche_introuvable"]])
    afficher_resume(compteurs, sautes, noms_differents, args.apply)
    afficher_bilan(compteurs, len(dossiers), nb_rattaches, len(items), args.apply, args.limit)

    if arret_erreurs:
        return CODE_SORTIE_ERREUR_FATALE
    if compteurs["erreur"]:
        return CODE_SORTIE_ERREURS_PARTIELLES
    return CODE_SORTIE_OK


def main():
    args = analyser_arguments()

    chemin_log = args.log_file or os.path.join(
        DOSSIER_LOGS, f"photos_papyrus_{datetime.now():%Y%m%d_%H%M%S}.log")
    configurer_journalisation(chemin_log, verbose=args.verbose)

    # Même verrou que les autres scripts : ils modifient les mêmes fiches.
    verrou = None
    if args.apply:
        verrou = acquerir_verrou(FICHIER_VERROU)
        if verrou is None:
            LOG.error(f"Une synchronisation est déjà en cours (verrou : {FICHIER_VERROU}) — abandon.")
            sys.exit(CODE_SORTIE_DEJA_EN_COURS)

    try:
        code = executer(args)
    except Exception:
        LOG.exception("Erreur inattendue — mise à jour des photos interrompue :")
        code = CODE_SORTIE_ERREUR_FATALE
    finally:
        if verrou:
            verrou.close()

    if code != CODE_SORTIE_OK:
        LOG.info(f"Code de sortie : {code} (détails dans {chemin_log})")
    sys.exit(code)


if __name__ == "__main__":
    main()
