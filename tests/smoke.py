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
                "documents": True, "regularisation_charges": True},
    "locataires": [
        {"nom": "Alice", "type_bien": "Appartement", "identifiant": "Appt 1",
         "adresse": "1 rue A", "loyer_nu": 500, "charges": 50, "part_caf": 200,
         "depot_garantie": 500, "date_entree": "2024-01-01"},
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
    attendus = ["Guide", "Locataires", "Appt 1", "Appt 2", "Données", "Bilan",
                "Régularisation charges", "Quittance", "Avis d'échéance", "Lettre de relance"]
    assert wb.sheetnames == attendus, wb.sheetnames
    assert wb["Données"].sheet_state == "hidden"
    assert wb["Quittance"]["B2"].value == "QUITTANCE DE LOYER"
    assert wb["Avis d'échéance"]["B2"].value == "AVIS D'ÉCHÉANCE"

    # Feuille locataire : colonnes CAF présentes + titre/identifiant.
    ws = wb["Appt 1"]
    assert ws.cell(1, 2).value == "Alice"
    assert ws.cell(1, 4).value == "Appt 1"
    _, ent = ligne_entete(ws)
    assert "CAF reçue" in ent and "Charges dues" in ent, list(ent)

    # 2) Période d'activité : Bob entré 07/2025 -> 6 lignes (07->12/2025).
    wsb = wb["Appt 2"]
    rb, _ = ligne_entete(wsb)
    nb_b = sum(1 for r in range(rb + 1, wsb.max_row + 1) if wsb.cell(r, 3).value)
    assert nb_b == 6, f"Bob devrait avoir 6 mois, obtenu {nb_b}"

    # 3) Bailleur minimal : pas de CAF, pas de documents.
    f2 = tmp / "minimal.xlsx"
    g.generer_workbook(g.valider_config(CFG_MINIMAL), f2)
    wb2 = load_workbook(f2)
    assert "Quittance" not in wb2.sheetnames, wb2.sheetnames
    assert "Maison" in wb2.sheetnames, wb2.sheetnames
    _, ent2 = ligne_entete(wb2["Maison"])
    assert not any("CAF" in (k or "") for k in ent2), list(ent2)
    assert "Loyer dû" in ent2, list(ent2)

    # 4) Préservation : on saisit dans la feuille d'Alice, on régénère avec un locataire en plus.
    wb = load_workbook(f1)
    ws = wb["Appt 1"]
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
    assert "Villa" in wb.sheetnames, wb.sheetnames
    ws = wb["Appt 1"]
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

    print("SMOKE OK")
    return 0


if __name__ == "__main__":
    sys.exit(main())
