"""Test du filtrage de la table Locataires de l'interface (sans affichage).

Couvre les fonctions pures de `interface_flet` : recherche texte, masquage des
locataires partis (toggle « Afficher les locataires partis ») et libellé du
compteur. Importer `interface_flet` demande Flet, mais aucune fenêtre n'est ouverte.
"""

import sys

import interface_flet as ui

LOCS = [
    {"nom": "Martin", "prenom": "Alice", "identifiant": "A101", "adresse": "Paris",
     "date_entree": "2024-01-15", "date_sortie": ""},
    {"nom": "Bernard", "prenom": "Sophie", "identifiant": "M2", "adresse": "Boulogne",
     "date_entree": "2023-06-01", "date_sortie": "2024-12-31"},
    {"nom": "Nguyen", "prenom": "Thomas", "identifiant": "B3", "adresse": "Paris",
     "date_entree": "2025-03-01"},  # clé date_sortie absente = présent
]


def indices(resultat):
    return [i for i, _ in resultat]


def main() -> int:
    # Statut « parti » : même règle que le badge État (date_sortie renseignée).
    assert [ui._est_parti(loc) for loc in LOCS] == [False, True, False]

    # Défaut (partis affichés) : comportement historique, tout est listé.
    assert indices(ui._filtrer_locataires(LOCS, "", True)) == [0, 1, 2]

    # Partis masqués : seuls les présents, avec leur index d'ORIGINE (modifier/supprimer).
    assert indices(ui._filtrer_locataires(LOCS, "", False)) == [0, 2]

    # Recherche et masquage se combinent.
    assert indices(ui._filtrer_locataires(LOCS, "paris", False)) == [0, 2]
    assert indices(ui._filtrer_locataires(LOCS, "sophie", True)) == [1]
    assert indices(ui._filtrer_locataires(LOCS, "sophie", False)) == []
    assert indices(ui._filtrer_locataires(LOCS, "  SOPHIE ", True)) == [1]

    # Partis masqués par le toggle parmi les résultats de la recherche.
    assert ui._partis_masques(LOCS, "", False) == 1
    assert ui._partis_masques(LOCS, "", True) == 0
    assert ui._partis_masques(LOCS, "alice", False) == 0
    assert ui._partis_masques(LOCS, "sophie", False) == 1

    # Compteur : inchangé sans filtre, explicite quand des partis sont masqués.
    assert ui._libelle_compteur(3, 3, 0) == "3 locataires"
    assert ui._libelle_compteur(1, 1, 0) == "1 locataire"
    assert ui._libelle_compteur(3, 1, 0) == "1 / 3"
    assert ui._libelle_compteur(3, 2, 1) == "2 / 3 · 1 parti masqué"
    assert ui._libelle_compteur(5, 2, 3) == "2 / 5 · 3 partis masqués"

    print("UI FILTRES OK")
    return 0


if __name__ == "__main__":
    sys.exit(main())
