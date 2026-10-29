"""Test des protections de la saisie en cours côté interface (sans affichage).

Couvre les fonctions pures de `interface_flet` qui évitent de perdre ce qui est saisi dans
le logiciel : confirmation avant de supprimer un locataire. Importer `interface_flet`
demande Flet, mais aucune fenêtre n'est ouverte.
"""

import sys

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

    print("UI PROTECTION OK")
    return 0


if __name__ == "__main__":
    sys.exit(main())
