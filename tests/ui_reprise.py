"""Test de la reprise des saisies côté interface (sans affichage).

Couvre les fonctions pures de `interface_flet` qui décident d'où reprendre les saisies
avant de générer (dernier classeur mémorisé par bailleur, question posée ou non) et qui
rédigent le compte rendu affiché après la génération. Importer `interface_flet` demande
Flet, mais aucune fenêtre n'est ouverte.
"""

import sys
import tempfile
from pathlib import Path

import generer_suivi_loyers as moteur
import interface_flet as ui


def main() -> int:
    # --- Faut-il demander d'où reprendre les saisies ? ---
    # Le fichier choisi est le dernier classeur généré pour ce bailleur : aucune question,
    # même s'il est encore vide (régénérations successives juste après la création).
    assert ui._cas_reprise(0, 0, True) == "direct"
    assert ui._cas_reprise(48, 0, True) == "direct"
    # Fichier choisi déjà rempli, aucun autre classeur connu : comportement historique.
    assert ui._cas_reprise(48, 0, False) == "direct"
    # Nouveau nom (ou fichier vide) alors que le dernier classeur a des saisies : c'est
    # l'erreur d'origine, on propose de les reprendre.
    assert ui._cas_reprise(None, 48, False) == "proposer"
    assert ui._cas_reprise(0, 48, False) == "proposer"
    # Les deux fichiers ont des saisies : l'utilisatrice choisit lesquelles reprendre.
    assert ui._cas_reprise(10, 48, False) == "choisir"
    # Neuf ou vide et aucun autre classeur connu : premier classeur, ou ancien à reprendre ?
    assert ui._cas_reprise(None, 0, False) == "premier"
    assert ui._cas_reprise(0, 0, False) == "premier"

    # --- Dernier classeur généré, mémorisé par bailleur (settings.json) ---
    tmp = Path(tempfile.mkdtemp())
    a = tmp / "Suivi_A.xlsx"
    a.write_bytes(b"x")
    reglages: dict = {}
    assert ui._classeur_memorise(reglages, "A") is None
    ui._memoriser_classeur(reglages, "A", a)
    assert ui._classeur_memorise(reglages, "A") == a.resolve()
    assert ui._classeur_memorise(reglages, "B") is None      # un classeur par bailleur
    a.unlink()
    assert ui._classeur_memorise(reglages, "A") is None      # fichier disparu : ignoré
    assert ui._classeur_memorise({"classeurs": "corrompu"}, "A") is None
    abime = {"classeurs": ["corrompu"]}
    ui._memoriser_classeur(abime, "A", tmp / "x.xlsx")       # réglage abîmé : réparé
    assert abime["classeurs"] == {"A": str((tmp / "x.xlsx").resolve())}, abime

    # --- Même fichier malgré un chemin écrit autrement ---
    assert ui._meme_fichier(tmp / "x.xlsx", tmp / "." / "x.xlsx")
    assert not ui._meme_fichier(tmp / "x.xlsx", tmp / "y.xlsx")

    # --- Accords ---
    assert ui._pluriel(1, "mois") == "1 mois" and ui._pluriel(2, "mois") == "2 mois"
    assert ui._pluriel(2, "année") == "2 années" and ui._pluriel(0, "indice") == "0 indice"
    assert ui._pluriel(3, "charge réelle", "charges réelles") == "3 charges réelles"

    # --- Compte rendu après génération ---
    R, E = moteur.RapportGeneration, moteur.Ecartee
    s = tmp / "Suivi.xlsx"
    bak = tmp / "Suivi.bak-20261029-182951.xlsx"
    copie = ("info", "Copie de sécurité de l'ancien fichier : "
                     "« Suivi.bak-20261029-182951.xlsx ».")
    # Classeur neuf.
    assert ui._lignes_rapport(R(), s) == [
        ("info", "Nouveau classeur : aucune saisie à reprendre.")]
    # Mise à jour sur place : saisies reprises + copie de sécurité.
    lignes = ui._lignes_rapport(R(source=s, sauvegarde=bak, mois_repris=48), s)
    assert lignes == [("ok", "Saisies reprises : 48 mois."), copie], lignes
    # Reprise depuis un autre classeur, avec charges réelles et indices IRL.
    old = tmp / "Ancien.xlsx"
    lignes = ui._lignes_rapport(R(source=old, mois_repris=1, charges_reprises=1,
                                  indices_repris=2), s)
    assert lignes == [("ok", "Saisies reprises de « Ancien.xlsx » : 1 mois, 1 charge "
                             "réelle et 2 indices IRL.")], lignes
    # Rien à reprendre dans le fichier lu.
    lignes = ui._lignes_rapport(R(source=s, sauvegarde=bak), s)
    assert lignes == [("info", "Aucune saisie à reprendre dans « Suivi.xlsx »."), copie]
    # Saisies écartées : chacune signalée avec son motif, puis où elles restent.
    lignes = ui._lignes_rapport(R(source=s, sauvegarde=bak, mois_repris=46, ecartees=[
        E("Dupont Jean", "locataire absent de la liste (renommé ou supprimé)", 2, "mois"),
        E("Charges réelles", "onglet Régularisation désactivé", 1, "année")]), s)
    assert lignes[2:] == [
        ("alerte", "Non repris : Dupont Jean, 2 mois — locataire absent de la liste "
                   "(renommé ou supprimé)."),
        ("alerte", "Non repris : Charges réelles, 1 année — onglet Régularisation "
                   "désactivé."),
        ("alerte", "Ces saisies restent dans « Suivi.bak-20261029-182951.xlsx » : rien "
                   "n'est perdu.")], lignes
    # Écartées d'un autre classeur : elles restent dans ce classeur-là (jamais modifié).
    lignes = ui._lignes_rapport(R(source=old, ecartees=[
        E("Ana", "mois hors de la période ou des dates du locataire", 1, "mois")]), s)
    assert lignes[-1] == ("alerte", "Ces saisies restent dans « Ancien.xlsx » : rien "
                                    "n'est perdu."), lignes

    print("UI REPRISE OK")
    return 0


if __name__ == "__main__":
    sys.exit(main())
