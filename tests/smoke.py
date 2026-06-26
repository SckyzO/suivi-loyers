"""Smoke test du moteur, autonome (dossier temporaire, ne pollue pas sorties/).

Couvre : architecture (feuille par locataire + Données masquée + documents), modularité,
période d'activité (rotation), et préservation des saisies lors d'une régénération.
"""

import sys
import tempfile
from pathlib import Path

from openpyxl import load_workbook

import generer_suivi_loyers as g

CFG_COMPLET = {
    "bailleur": {"nom": "Smoke Complet"},
    "periode": {"annee_debut": 2024, "annee_fin": 2025},
    "modules": {"loyer_nu_charges": True, "caf": True, "depot_garantie": True,
                "documents": True, "regularisation_charges": True, "irl": True},
    "locataires": [
        # Nom + prénom, parti en cours de période (caution + observation).
        {"nom": "Alice", "prenom": "A", "type_bien": "Appartement", "identifiant": "Appt 1",
         "adresse": "1 rue A", "loyer_nu": 500, "charges": 50, "part_caf": 200,
         "depot_garantie": 500, "date_entree": "2024-01-01", "date_sortie": "2024-06-30",
         "caution_rendue": False, "observation": "Travaux"},
        {"nom": "Bob", "type_bien": "Appartement", "identifiant": "Appt 2",
         "adresse": "1 rue A", "loyer_nu": 400, "charges": 40, "part_caf": 0,
         "depot_garantie": 400, "date_entree": "2025-07-01"},  # entré en cours de période
    ],
}

CFG_MINIMAL = {
    "bailleur": {"nom": "Smoke Minimal"},
    "periode": {"annee_debut": 2025, "annee_fin": 2025},
    "modules": {"loyer_nu_charges": False, "caf": False, "depot_garantie": False,
                "documents": False},
    "locataires": [{"nom": "Carl", "type_bien": "Maison", "identifiant": "Maison",
                    "adresse": "2 rue B", "loyer": 750, "date_entree": "2025-01-01"}],
}


def ligne_entete(ws):
    for r in range(1, 12):
        vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        if "Mois" in vals and "Année" in vals:
            return r, {v: c for c, v in enumerate(vals, 1) if v}
    raise AssertionError("en-tête introuvable")


def main() -> int:
    tmp = Path(tempfile.mkdtemp())

    # 1) Bailleur complet : architecture attendue.
    f1 = tmp / "complet.xlsx"
    g.generer_workbook(g.valider_config(CFG_COMPLET), f1)
    wb = load_workbook(f1)
    # Onglet locataire nommé « identifiant - Nom » (évite les doublons même appartement).
    attendus = ["Guide", "Locataires", "Appt 1 - Alice", "Appt 2 - Bob", "Données", "Bilan",
                "Régularisation charges", "Révision IRL",
                "Quittance", "Avis d'échéance", "Lettre de relance"]
    assert wb.sheetnames == attendus, wb.sheetnames
    assert wb["Données"].sheet_state == "hidden"
    assert wb["Quittance"]["B2"].value == "QUITTANCE DE LOYER"
    assert wb["Avis d'échéance"]["B2"].value == "AVIS D'ÉCHÉANCE"

    # Identité = « Nom Prénom ».
    assert g._identite({"nom": "Martin", "prenom": "Sophie"}) == "Martin Sophie"
    assert g._identite({"nom": "Old"}) == "Old"

    # Onglet Locataires : Type après Adresse + colonnes Caution / Observation renseignées.
    loc = wb["Locataires"]
    hl = {loc.cell(1, c).value: c for c in range(1, loc.max_column + 1)}
    assert list(hl).index("Type de bien") > list(hl).index("Adresse du logement")
    assert loc.cell(2, hl["Caution rendue"]).value == "Non"
    assert loc.cell(2, hl["Observation (motif de départ)"]).value == "Travaux"

    # Feuille locataire : titre = identité, colonnes CAF présentes.
    ws = wb["Appt 1 - Alice"]
    assert ws.cell(1, 2).value == "Alice A"
    assert ws.cell(1, 4).value == "Appt 1"
    _, ent = ligne_entete(ws)
    assert "CAF reçue" in ent and "Charges dues" in ent, list(ent)

    # 2) Période d'activité : Bob entré 07/2025 -> 6 mois. Totaux annuels présents.
    wsb = wb["Appt 2 - Bob"]
    rb, _ = ligne_entete(wsb)
    nb_b = sum(1 for r in range(rb + 1, wsb.max_row + 1)
               if wsb.cell(r, 3).value in g.MOIS)
    assert nb_b == 6, f"Bob devrait avoir 6 mois, obtenu {nb_b}"
    assert any(str(wsb.cell(r, 3).value or "").startswith("Total")
               for r in range(rb + 1, wsb.max_row + 1)), "ligne Total annuel manquante"

    # 3) Bailleur minimal : pas de CAF, pas de documents.
    f2 = tmp / "minimal.xlsx"
    g.generer_workbook(g.valider_config(CFG_MINIMAL), f2)
    wb2 = load_workbook(f2)
    assert "Quittance" not in wb2.sheetnames, wb2.sheetnames
    assert "Maison - Carl" in wb2.sheetnames, wb2.sheetnames
    _, ent2 = ligne_entete(wb2["Maison - Carl"])
    assert not any("CAF" in (k or "") for k in ent2), list(ent2)
    assert "Loyer dû" in ent2, list(ent2)

    # 4) Préservation : on saisit dans la feuille d'Alice, on régénère avec un locataire en plus.
    wb = load_workbook(f1)
    ws = wb["Appt 1 - Alice"]
    r_ent, ent = ligne_entete(ws)
    r0 = r_ent + 1
    ws.cell(r0, ent["Part locataire reçue"]).value = 333
    cle = (ws.cell(r0, ent["Locataire"]).value, ws.cell(r0, ent["Année"]).value,
           ws.cell(r0, ent["Mois"]).value)
    wb.save(f1)

    cfg2 = g.valider_config(CFG_COMPLET)
    cfg2["locataires"].append({"nom": "Dora", "type_bien": "Maison", "identifiant": "Villa",
                               "adresse": "9 rue C", "loyer_nu": 300, "charges": 20,
                               "part_caf": 100, "depot_garantie": 300,
                               "date_entree": "2025-01-01"})
    g.generer_workbook(cfg2, f1, preserver=True)

    wb = load_workbook(f1)
    assert "Villa - Dora" in wb.sheetnames, wb.sheetnames
    ws = wb["Appt 1 - Alice"]
    r_ent, ent = ligne_entete(ws)
    trouve = None
    for r in range(r_ent + 1, ws.max_row + 1):
        if (ws.cell(r, ent["Locataire"]).value, ws.cell(r, ent["Année"]).value,
                ws.cell(r, ent["Mois"]).value) == cle:
            trouve = ws.cell(r, ent["Part locataire reçue"]).value
            break
    assert trouve == 333, f"saisie non préservée: {trouve}"

    # 5) Régularisation : saisie d'une charge réelle, préservée après régénération.
    wb = load_workbook(f1)
    reg = wb["Régularisation charges"]
    ent = {reg.cell(1, c).value: c for c in range(1, reg.max_column + 1)}
    reg.cell(2, ent["Charges réelles (€)"]).value = 612
    cle_reg = (reg.cell(2, ent["Locataire"]).value, reg.cell(2, ent["Année"]).value)
    wb.save(f1)
    g.generer_workbook(g.valider_config(CFG_COMPLET), f1, preserver=True)
    reg = load_workbook(f1)["Régularisation charges"]
    ent = {reg.cell(1, c).value: c for c in range(1, reg.max_column + 1)}
    val = None
    for r in range(2, reg.max_row + 1):
        if (reg.cell(r, ent["Locataire"]).value, reg.cell(r, ent["Année"]).value) == cle_reg:
            val = reg.cell(r, ent["Charges réelles (€)"]).value
            break
    assert val == 612, f"charge réelle non préservée: {val}"

    # 5b) IRL : saisie d'un indice et d'un choix de révision, préservés après régénération.
    wb = load_workbook(f1)
    irl = wb["Révision IRL"]
    # 1er indice (ligne 5, colonne C) et 1re révision locataire.
    irl.cell(5, 3).value = 145.17                      # valeur IRL T1 année début
    h_rev = next(r for r in range(1, irl.max_row + 1)
                 if irl.cell(r, 7).value == "Nouveau loyer (€)")
    irl.cell(h_rev + 1, 3).value = "T2"                # trimestre réf. 1er locataire
    nom_irl = irl.cell(h_rev + 1, 1).value
    wb.save(f1)
    g.generer_workbook(g.valider_config(CFG_COMPLET), f1, preserver=True)
    irl = load_workbook(f1)["Révision IRL"]
    assert irl.cell(5, 3).value == 145.17, irl.cell(5, 3).value
    h_rev = next(r for r in range(1, irl.max_row + 1)
                 if irl.cell(r, 7).value == "Nouveau loyer (€)")
    assert irl.cell(h_rev + 1, 1).value == nom_irl
    assert irl.cell(h_rev + 1, 3).value == "T2", irl.cell(h_rev + 1, 3).value

    # 6) Rétro-compatibilité : une « ancienne » config (champ bien, module quittances)
    #    se migre et génère sans erreur.
    legacy = {
        "bailleur": {"nom": "Legacy"},
        "periode": {"annee_debut": 2025, "annee_fin": 2025},
        "modules": {"loyer_nu_charges": True, "caf": True, "depot_garantie": False,
                    "quittances": True},
        "locataires": [{"nom": "Old", "bien": "Appt Z", "loyer_nu": 400, "charges": 30,
                        "part_caf": 100, "date_entree": "2025-01-01"}],
    }
    migre, avertis = g.migrer_config(legacy)
    assert migre["modules"].get("documents") is True, migre["modules"]
    assert migre["locataires"][0]["identifiant"] == "Appt Z"
    assert avertis, "des avertissements de migration étaient attendus"
    f3 = tmp / "legacy.xlsx"
    g.generer_workbook(g.valider_config(legacy), f3)
    wb3 = load_workbook(f3)
    assert "Appt Z - Old" in wb3.sheetnames and "Quittance" in wb3.sheetnames, wb3.sheetnames

    print("SMOKE OK")
    return 0


if __name__ == "__main__":
    sys.exit(main())
