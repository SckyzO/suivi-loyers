"""Test des protections de la saisie en cours côté interface (sans affichage).

Couvre les fonctions pures de `interface_flet` qui évitent de perdre ce qui est saisi dans
le logiciel : confirmation avant de supprimer un locataire, sauvegarde automatique de la
saisie en cours (session.json), sa reprise au démarrage, les saisies mises de côté, le
chargement d'une configuration et l'écriture des réglages. Importer `interface_flet`
demande Flet, mais aucune fenêtre n'est ouverte.
"""

import json
import os
import sys
import tempfile
from pathlib import Path

import interface_flet as ui


def main() -> int:
    # --- Suppression d'un locataire : confirmation ---
    present = {"nom": "Martin", "prenom": "Alice", "date_sortie": ""}
    parti = {"nom": "Bernard", "prenom": "Sophie", "date_sortie": "2024-12-31"}
    titre, paragraphes, depart = ui._messages_suppression(present)
    assert titre == "Supprimer Martin Alice ?", titre
    # Les saisies ne seront plus reportées : dit clairement, avec où elles restent.
    assert "copie de sécurité" in paragraphes[0], paragraphes
    # Locataire encore présent : on propose plutôt d'indiquer son départ.
    assert depart and "date de sortie" in paragraphes[-1], (depart, paragraphes)
    # Déjà parti : pas de proposition de départ.
    titre, paragraphes, depart = ui._messages_suppression(parti)
    assert titre == "Supprimer Bernard Sophie ?" and not depart, (titre, depart)
    assert len(paragraphes) == 1, paragraphes
    assert ui._messages_suppression({"nom": "", "prenom": ""})[0] == "Supprimer ce locataire ?"

    # --- Saisie en cours : y a-t-il quelque chose à protéger ? ---
    assert not ui._a_du_contenu({})
    assert not ui._a_du_contenu({"bailleur": {"nom": "  "}, "locataires": []})
    assert ui._a_du_contenu({"bailleur": {"nom": "Durand"}})
    assert ui._a_du_contenu({"bailleur": {}, "locataires": [{"nom": "Ana"}]})
    # L'empreinte ne dépend pas de l'ordre des clés (comparaison avec l'état enregistré).
    assert ui._empreinte({"a": 1, "b": 2}) == ui._empreinte({"b": 2, "a": 1})

    # --- Session précédente : la proposer seulement si elle n'était pas enregistrée ---
    cfg = {"bailleur": {"nom": "Durand", "sci": True, "sci_nom": "SCI Lilas"},
           "locataires": [{"nom": "Ana"}, {"nom": "Bob"}, {"nom": "Eve"}]}
    assert ui._session_a_reprendre({}) is None
    assert ui._session_a_reprendre({"config": cfg, "propre": True}) is None
    assert ui._session_a_reprendre({"config": {"bailleur": {}}, "propre": False}) is None
    assert ui._session_a_reprendre({"config": "abîmée", "propre": False}) is None
    assert ui._session_a_reprendre({"config": cfg, "propre": False}) == cfg
    assert ui._session_a_reprendre({"config": cfg}) is None   # sans drapeau : prudence
    session = {"config": cfg, "propre": False, "horodatage": "2026-10-29T18:34:05"}
    assert ui._resume_session(session) == \
        "SCI Lilas Durand — 3 locataires, modifiée le 29/10/2026 à 18:34.", \
        ui._resume_session(session)
    assert ui._resume_session({"config": {"bailleur": {}, "locataires": [{"nom": "A"}]}}) \
        == "Bailleur sans nom — 1 locataire."

    # --- « Charger config » : jamais un formulaire vidé sans prévenir ---
    assert len(ui._lire_config(cfg)[0]["locataires"]) == 3          # configuration normale
    # Une saisie mise de côté ({"config": …}) se recharge comme une configuration.
    relue, _ = ui._lire_config({"config": cfg, "propre": False, "horodatage": "x"})
    assert relue["bailleur"]["nom"] == "Durand" and len(relue["locataires"]) == 3, relue
    # Un fichier qui n'est pas une configuration est refusé : il chargeait un formulaire
    # vide, sans avertissement, et le marquait comme enregistré.
    for brut in ({}, {"verifier_maj": True, "classeurs": {}}, [], "texte",
                 {"config": "abîmée"}):
        try:
            ui._lire_config(brut)
        except ValueError:
            continue
        raise AssertionError(f"accepté à tort : {brut!r}")

    # --- Écriture, relecture et mise de côté du fichier de session ---
    os.environ["APPDATA"] = tempfile.mkdtemp()           # _fichier_reglages() le lit à l'appel
    dossier = Path(os.environ["APPDATA"]) / "SuiviLoyers"
    assert ui._charger_session() == {}                   # aucune session : rien à reprendre
    ui._enregistrer_session(cfg, propre=False)
    relu = ui._charger_session()
    assert relu["config"] == cfg and relu["propre"] is False and relu["horodatage"], relu
    assert sorted(p.name for p in dossier.iterdir()) == ["session.json"], \
        "écriture atomique : aucun fichier temporaire ne doit rester"
    # Session écartée : mise de côté dans son propre fichier. Une deuxième ne doit pas
    # effacer la première (elles partageaient session-ecartee.json, écrasé à chaque fois).
    premiere = ui._archiver_session()
    assert ui._charger_session() == {} and premiere and premiere.is_file(), premiere
    ui._enregistrer_session({"bailleur": {"nom": "Dupont"}, "locataires": []}, propre=False)
    seconde = ui._archiver_session()                     # même seconde : nom distinct
    assert seconde and seconde != premiere and seconde.is_file(), (premiere, seconde)
    assert json.loads(premiere.read_text(encoding="utf-8"))["config"] == cfg
    assert ui._archiver_session() is None                # plus de session : rien à faire
    (dossier / "session.json").write_text("{ illisible", encoding="utf-8")
    assert ui._charger_session() == {}                   # corrompue : ignorée, pas d'erreur

    # --- Saisies mises de côté : proposées dans les Réglages, la plus récente d'abord ---
    os.environ["APPDATA"] = tempfile.mkdtemp()
    assert ui._saisies_mises_de_cote() == []             # aucune : section masquée
    mdc = ui._dossier_mises_de_cote()
    mdc.mkdir(parents=True)
    for fichier, quand, nom in (("a.json", "2026-10-28T09:02:00", "Dupont"),
                                ("b.json", "2026-10-30T12:15:00", "Martin")):
        (mdc / fichier).write_text(json.dumps(
            {"config": {"bailleur": {"nom": nom}, "locataires": [{"nom": "Ana"}]},
             "propre": False, "horodatage": quand}), encoding="utf-8")
    (mdc / "abimee.json").write_text("{ illisible", encoding="utf-8")
    liste = ui._saisies_mises_de_cote()
    assert [s["config"]["bailleur"]["nom"] for s in liste] == ["Martin", "Dupont"], liste
    assert (mdc / "abimee.json").is_file()               # illisible : ignorée, pas supprimée
    assert ui._libelles_session(liste[0]) == \
        ("Martin — 1 locataire", "modifiée le 30/10/2026 à 12:15"), \
        ui._libelles_session(liste[0])
    assert ui._libelles_session({"config": {"bailleur": {}, "locataires": []}}) == \
        ("Bailleur sans nom — 0 locataire", "")

    # --- Réglages (dont le dernier classeur de chaque bailleur) : écriture atomique ---
    os.environ["APPDATA"] = tempfile.mkdtemp()
    ui._enregistrer_reglages({"aide_vue": True, "classeurs": {"durand": "/x/Suivi.xlsx"}})
    assert ui._charger_reglages()["classeurs"] == {"durand": "/x/Suivi.xlsx"}
    reglages = ui._fichier_reglages()
    assert sorted(p.name for p in reglages.parent.iterdir()) == ["settings.json"], \
        "écriture atomique : aucun fichier temporaire ne doit rester"
    # Un échec d'écriture remonte à l'appelant, qui le signale : il était avalé, et le
    # logiciel oubliait le dernier classeur sans prévenir.
    reglages.unlink()
    reglages.mkdir()                                     # cible impossible à remplacer
    try:
        ui._enregistrer_reglages({"aide_vue": True})
    except OSError:
        pass
    else:
        raise AssertionError("échec d'écriture des réglages avalé sans prévenir")

    print("UI PROTECTION OK")
    return 0


if __name__ == "__main__":
    sys.exit(main())
