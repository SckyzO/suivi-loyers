"""Smoke test du moteur, autonome (écrit dans un dossier temporaire, ne pollue pas sorties/).

Vérifie : structure des onglets, colonnes pilotées par les modules (modularité), génération
sur la période d'activité (rotation), et préservation des saisies lors d'une régénération.
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
                "quittances": True},
    "locataires": [
        {"nom": "A", "bien": "L1", "loyer_nu": 500, "charges": 50, "part_caf": 200,
         "depot_garantie": 500, "date_entree": "2024-01-01"},
        {"nom": "B", "bien": "L2", "loyer_nu": 400, "charges": 40, "part_caf": 0,
         "depot_garantie": 400, "date_entree": "2025-07-01"},  # entré en cours d'année
    ],
}

CFG_MINIMAL = {
    "bailleur": {"nom": "Smoke Minimal"},
    "periode": {"annee_debut": 2025, "annee_fin": 2025},
    "modules": {"loyer_nu_charges": False, "caf": False, "depot_garantie": False},
    "locataires": [{"nom": "C", "bien": "M", "loyer": 750, "date_entree": "2025-01-01"}],
}


def entetes(ws):
    return [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]


def colonnes_map(ws):
    return {ws.cell(1, c).value: c for c in range(1, ws.max_column + 1)}


def main() -> int:
    tmp = Path(tempfile.mkdtemp())

    # 1) Bailleur complet : structure + colonnes CAF/charges + période d'activité.
    f1 = tmp / "complet.xlsx"
    g.generer_workbook(g.valider_config(CFG_COMPLET), f1)
    wb = load_workbook(f1)
    ws = wb["Suivi"]
    assert wb.sheetnames == ["Guide", "Locataires", "Suivi", "Bilan", "Quittance"], wb.sheetnames
    assert wb["Quittance"]["B2"].value == "QUITTANCE DE LOYER"
    cols = entetes(ws)
    assert "CAF reçue" in cols and "Charges dues" in cols, cols
    cmap = colonnes_map(ws)
    nb_b = sum(1 for r in range(2, ws.max_row + 1)
               if ws.cell(r, cmap["Locataire"]).value == "B")
    assert nb_b == 6, f"B devrait avoir 6 mois (07->12/2025), obtenu {nb_b}"

    # 2) Bailleur minimal : aucune colonne CAF, libellé « Loyer dû ».
    f2 = tmp / "minimal.xlsx"
    g.generer_workbook(g.valider_config(CFG_MINIMAL), f2)
    wb2 = load_workbook(f2)
    cols2 = entetes(wb2["Suivi"])
    assert not any("CAF" in (c or "") for c in cols2), cols2
    assert "Loyer dû" in cols2, cols2
    assert "Quittance" not in wb2.sheetnames, wb2.sheetnames

    # 3) Préservation des saisies : on saisit, on régénère avec un locataire en plus.
    wb = load_workbook(f1)
    ws = wb["Suivi"]
    ent = colonnes_map(ws)
    ws.cell(2, ent["Part locataire reçue"]).value = 333
    cle = (ws.cell(2, ent["Locataire"]).value, ws.cell(2, ent["Année"]).value,
           ws.cell(2, ent["Mois"]).value)
    wb.save(f1)

    cfg2 = g.valider_config(CFG_COMPLET)
    cfg2["locataires"].append({"nom": "D", "bien": "L3", "loyer_nu": 300, "charges": 20,
                               "part_caf": 100, "depot_garantie": 300, "date_entree": "2025-01-01"})
    g.generer_workbook(cfg2, f1, preserver=True)

    ws = load_workbook(f1)["Suivi"]
    ent = colonnes_map(ws)
    trouve = None
    for r in range(2, ws.max_row + 1):
        if (ws.cell(r, ent["Locataire"]).value, ws.cell(r, ent["Année"]).value,
                ws.cell(r, ent["Mois"]).value) == cle:
            trouve = ws.cell(r, ent["Part locataire reçue"]).value
            break
    assert trouve == 333, f"saisie non préservée: {trouve}"

    print("SMOKE OK")
    return 0


if __name__ == "__main__":
    sys.exit(main())
