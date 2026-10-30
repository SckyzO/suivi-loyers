"""Test du filtrage de la table Locataires de l'interface (sans affichage).

Couvre les fonctions pures de `interface_flet` : recherche texte, masquage des
locataires partis (toggle « Afficher les partis »), libellé du compteur, colonnes
affichées et disposition de la barre selon la largeur du tableau. Importer
`interface_flet` demande Flet, mais aucune fenêtre n'est ouverte.
"""

import sys
from types import SimpleNamespace

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

    # Colonnes selon la largeur du tableau : en fenêtre étroite, les moins utiles sont
    # retirées (Adresse, CAF, Période, Bien) au lieu de tout tronquer. Définitions réelles
    # de la classe (les lambdas de tri ne sont pas appelées ici).
    faux = SimpleNamespace(m_caf=SimpleNamespace(value=True),
                           _loyer_loc=lambda loc: 0, _nom_complet=lambda loc: "")
    cols = ui.AppLoyers._colonnes(faux)

    def cles(largeur):
        return [c["cle"] for c in ui._colonnes_pour_largeur(cols, largeur)]
    toutes = ["nom", "bien", "adresse", "loyer", "caf", "periode", "etat", "actions"]
    assert cles(None) == toutes                  # avant le premier rendu : tout
    # Fenêtre par défaut (1180 px) : tableau ≈ 1180 - 400 (volet) - 48 (marges) - 2 = 730
    # px -> rien ne change.
    assert cles(730) == toutes, cles(730)
    assert cles(650) == ["nom", "bien", "loyer", "caf", "periode", "etat", "actions"]
    # Fenêtre minimale (1000 px) : tableau ≈ 550 px -> l'essentiel, lisible.
    assert cles(550) == ["nom", "bien", "loyer", "etat", "actions"], cles(550)
    # Jamais moins que nom, loyer, état et actions.
    assert cles(100) == ["nom", "loyer", "etat", "actions"]
    # Sans module CAF, la colonne n'existe pas : rien d'autre ne casse.
    sans_caf = ui.AppLoyers._colonnes(SimpleNamespace(
        m_caf=SimpleNamespace(value=False), _loyer_loc=lambda loc: 0,
        _nom_complet=lambda loc: ""))
    assert [c["cle"] for c in ui._colonnes_pour_largeur(sans_caf, 730)] == \
        [c for c in toutes if c != "caf"]

    # Valeur entière en bulle au survol pour les colonnes que la largeur coupe (« Martin
    # Ali… » à la taille par défaut) ; rien à montrer pour une valeur absente (« — »).
    faux_cell = SimpleNamespace(_nom_complet=ui._nom_complet, _dim_col=ui.AppLoyers._dim_col,
                                _mois_an=ui.AppLoyers._mois_an)
    faux_cell._periode_txt = lambda loc: ui.AppLoyers._periode_txt(faux_cell, loc)

    def bulle(cle, loc):
        cellule = ui.AppLoyers._cellule(faux_cell, next(c for c in cols if c["cle"] == cle),
                                        loc, 0)
        texte = cellule.content.controls[-1] if cle == "nom" else cellule.content
        return texte.tooltip
    sophie = LOCS[1]
    assert [bulle(c, sophie) for c in ("nom", "bien", "adresse", "periode")] == \
        ["Bernard Sophie", "M2", "Boulogne", "06/2023 → 12/2024"], \
        [bulle(c, sophie) for c in ("nom", "bien", "adresse", "periode")]
    inconnu = {"nom": "", "prenom": ""}
    assert [bulle(c, inconnu) for c in ("nom", "bien", "adresse", "periode")] == \
        [None] * 4

    # Barre de la liste (« Afficher les partis », recherche, « Ajouter ») : sur la ligne du
    # titre à la largeur par défaut, sous le titre en fenêtre minimale plutôt que coupée.
    assert ui._barre_sur_une_ligne(None)         # avant le premier rendu
    assert ui._barre_sur_une_ligne(730)
    assert not ui._barre_sur_une_ligne(550)

    print("UI FILTRES OK")
    return 0


if __name__ == "__main__":
    sys.exit(main())
