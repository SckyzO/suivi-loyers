#!/usr/bin/env python3
"""Moteur de génération des classeurs Excel de suivi des loyers (part locataire + part CAF).

Un fichier .xlsx est produit par bailleur. Le contenu (colonnes et onglets) est piloté
par les `modules` activés : un bailleur sans CAF n'aura pas les colonnes CAF, un bailleur
sans distinction loyer nu / charges aura une simple colonne « Loyer dû », etc.

Deux usages :
  * Interface graphique (interface.py) — pour l'utilisateur final (Windows, .exe).
  * Ligne de commande / Docker — pour le mainteneur :
        python generer_suivi_loyers.py <config.yaml> [dossier_sortie]

Particularités importantes :
  * Les lignes du suivi ne sont créées que sur la période d'activité de chaque locataire
    (date d'entrée → date de sortie) : gère proprement les rotations fréquentes de locataires.
  * Lors d'une régénération sur un fichier existant, les montants déjà saisis (CAF reçue,
    part locataire reçue, dates) sont PRÉSERVÉS et réinjectés (clé = locataire + année + mois).
"""

from __future__ import annotations

import sys
import re
import datetime as dt
from pathlib import Path

import yaml
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.workbook.defined_name import DefinedName

# --------------------------------------------------------------------------- #
# Constantes de présentation
# --------------------------------------------------------------------------- #

MOIS = [
    "Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
    "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre",
]

FMT_EURO = '#,##0.00\\ "€"'
FMT_DATE = "DD/MM/YYYY"
FMT_PCT = "0.0%"

COUL_ENTETE = "1F4E79"
COUL_ENTETE_TXT = "FFFFFF"
COUL_SAISIE = "FFF7E6"      # jaune pâle : cellules à remplir
COUL_CALC = "EEF3F8"        # bleu très pâle : cellules calculées
COUL_SOLDE = "C6EFCE"       # vert
COUL_TROP = "FFEB9C"        # orange
COUL_PARTIEL = "FFC7CE"     # rouge
COUL_ATTENTE = "E7E6E6"     # gris

_THIN = Side(style="thin", color="BFBFBF")
BORDURE = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)

# Colonnes du Suivi dont la saisie utilisateur doit être préservée lors d'une régénération.
COLS_SAISIE = ("caf_recu", "caf_date", "loc_recu", "loc_date")

MODULES_DEFAUT = {
    "loyer_nu_charges": True,
    "caf": True,
    "depot_garantie": True,
    "quittances": False,             # phase 2
    "irl": False,                    # phase 2
    "regularisation_charges": False,  # phase 2
}


# --------------------------------------------------------------------------- #
# Config : validation (dict) + chargement YAML
# --------------------------------------------------------------------------- #

def valider_config(raw: dict) -> dict:
    """Valide et normalise une config (dict). Utilisée par le YAML et par l'interface."""
    if not isinstance(raw, dict):
        raise ValueError("La config doit être un mapping (clés bailleur, periode, …).")

    bailleur = raw.get("bailleur") or {}
    if not bailleur.get("nom"):
        raise ValueError("Le nom du bailleur est obligatoire.")

    periode = raw.get("periode") or {}
    annee_debut = int(periode.get("annee_debut", dt.date.today().year))
    annee_fin = int(periode.get("annee_fin", annee_debut))
    if annee_fin < annee_debut:
        raise ValueError("L'année de fin doit être supérieure ou égale à l'année de début.")

    modules = dict(MODULES_DEFAUT)
    modules.update(raw.get("modules") or {})

    locataires = raw.get("locataires") or []
    if not locataires:
        raise ValueError("Il faut au moins un locataire.")
    for i, loc in enumerate(locataires, 1):
        if not loc.get("nom"):
            raise ValueError(f"Le locataire #{i} n'a pas de nom.")

    return {
        "bailleur": bailleur,
        "annee_debut": annee_debut,
        "annee_fin": annee_fin,
        "modules": modules,
        "locataires": locataires,
    }


def charger_config(chemin: Path) -> dict:
    with Path(chemin).open(encoding="utf-8") as f:
        return valider_config(yaml.safe_load(f))


def _num(loc: dict, *cles) -> float | None:
    for c in cles:
        if loc.get(c) is not None:
            return float(loc[c])
    return None


def _date(val) -> dt.date | None:
    if val is None or val == "":
        return None
    if isinstance(val, dt.datetime):
        return val.date()
    if isinstance(val, dt.date):
        return val
    return dt.date.fromisoformat(str(val))


def _mois_actifs(loc: dict, annee_debut: int, annee_fin: int) -> list[tuple[int, int]]:
    """Liste des (année, mois) où le locataire est présent, bornée par entrée/sortie."""
    de = _date(loc.get("date_entree"))
    ds = _date(loc.get("date_sortie"))
    debut = (de.year, de.month) if de else (annee_debut, 1)
    fin = (ds.year, ds.month) if ds else (annee_fin, 12)
    actifs = []
    for annee in range(annee_debut, annee_fin + 1):
        for mois in range(1, 13):
            if debut <= (annee, mois) <= fin:
                actifs.append((annee, mois))
    return actifs


# --------------------------------------------------------------------------- #
# Helpers de style
# --------------------------------------------------------------------------- #

def style_entete(cell) -> None:
    cell.font = Font(bold=True, color=COUL_ENTETE_TXT)
    cell.fill = PatternFill("solid", fgColor=COUL_ENTETE)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = BORDURE


def style_cellule(cell, *, saisie=False, calc=False, fmt=None) -> None:
    if saisie:
        cell.fill = PatternFill("solid", fgColor=COUL_SAISIE)
    elif calc:
        cell.fill = PatternFill("solid", fgColor=COUL_CALC)
    if fmt:
        cell.number_format = fmt
    cell.border = BORDURE


# --------------------------------------------------------------------------- #
# Onglet Locataires (référentiel + plages nommées)
# --------------------------------------------------------------------------- #

def construire_locataires(wb: Workbook, cfg: dict) -> dict:
    mod = cfg["modules"]
    split, caf, depot = mod["loyer_nu_charges"], mod["caf"], mod["depot_garantie"]

    ws = wb.create_sheet("Locataires")

    cols: list[tuple[str, str]] = [("nom", "Locataire"), ("bien", "Bien / logement")]
    if split:
        cols += [("loyer_nu", "Loyer nu (€)"), ("charges", "Charges (€)")]
    cols.append(("loyer_total", "Loyer total (€)"))
    if caf:
        cols.append(("part_caf", "Part CAF / APL (€)"))
    cols.append(("reste", "Reste à charge (€)"))
    if depot:
        cols.append(("depot", "Dépôt garantie (€)"))
    cols += [("date_entree", "Date entrée"), ("date_sortie", "Date sortie")]

    idx = {champ: i + 1 for i, (champ, _) in enumerate(cols)}
    lettre = {champ: get_column_letter(i + 1) for i, (champ, _) in enumerate(cols)}

    for i, (_, titre) in enumerate(cols, 1):
        style_entete(ws.cell(row=1, column=i, value=titre))

    largeurs = {
        "nom": 24, "bien": 30, "loyer_nu": 13, "charges": 13, "loyer_total": 14,
        "part_caf": 16, "reste": 16, "depot": 16, "date_entree": 13, "date_sortie": 13,
    }
    for champ, l in lettre.items():
        ws.column_dimensions[l].width = largeurs.get(champ, 14)

    for r, loc in enumerate(cfg["locataires"], start=2):
        ws.cell(row=r, column=idx["nom"], value=loc.get("nom"))
        ws.cell(row=r, column=idx["bien"], value=loc.get("bien", ""))

        if split:
            ws.cell(row=r, column=idx["loyer_nu"], value=_num(loc, "loyer_nu") or 0)
            ws.cell(row=r, column=idx["charges"], value=_num(loc, "charges") or 0)
            ws.cell(row=r, column=idx["loyer_total"],
                    value=f"={lettre['loyer_nu']}{r}+{lettre['charges']}{r}")
        else:
            ws.cell(row=r, column=idx["loyer_total"],
                    value=_num(loc, "loyer", "loyer_total", "loyer_nu") or 0)

        if caf:
            ws.cell(row=r, column=idx["part_caf"], value=_num(loc, "part_caf") or 0)
            ws.cell(row=r, column=idx["reste"],
                    value=f"={lettre['loyer_total']}{r}-{lettre['part_caf']}{r}")
        else:
            ws.cell(row=r, column=idx["reste"], value=f"={lettre['loyer_total']}{r}")

        if depot:
            ws.cell(row=r, column=idx["depot"], value=_num(loc, "depot_garantie") or 0)

        if (de := _date(loc.get("date_entree"))):
            ws.cell(row=r, column=idx["date_entree"], value=de)
        if (ds := _date(loc.get("date_sortie"))):
            ws.cell(row=r, column=idx["date_sortie"], value=ds)

    derniere = len(cfg["locataires"]) + 1

    fmt_par_champ = {
        "loyer_nu": FMT_EURO, "charges": FMT_EURO, "loyer_total": FMT_EURO,
        "part_caf": FMT_EURO, "reste": FMT_EURO, "depot": FMT_EURO,
        "date_entree": FMT_DATE, "date_sortie": FMT_DATE,
    }
    cols_saisie = {"nom", "bien", "loyer_nu", "charges", "part_caf", "depot",
                   "date_entree", "date_sortie"}
    for r in range(2, derniere + 1):
        for champ, i in idx.items():
            cell = ws.cell(row=r, column=i)
            est_formule = isinstance(cell.value, str) and cell.value.startswith("=")
            style_cellule(cell, fmt=fmt_par_champ.get(champ),
                          calc=est_formule, saisie=champ in cols_saisie and not est_formule)

    ref = f"A1:{lettre['date_sortie']}{derniere}"
    table = Table(displayName="TblLocataires", ref=ref)
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2", showRowStripes=True, showColumnStripes=False)
    ws.add_table(table)
    ws.freeze_panes = "A2"

    der_col = lettre["date_sortie"]
    wb.defined_names.add(DefinedName(
        "LocatairesListe", attr_text=f"Locataires!$A$2:$A${derniere}"))
    wb.defined_names.add(DefinedName(
        "RefLocataires", attr_text=f"Locataires!$A$2:${der_col}${derniere}"))

    return {"idx": idx, "n": len(cfg["locataires"])}


# --------------------------------------------------------------------------- #
# Onglet Suivi (saisie mensuelle, colonnes dynamiques, période d'activité)
# --------------------------------------------------------------------------- #

def construire_suivi(wb: Workbook, cfg: dict, ref_loc: dict, saisies: dict | None = None) -> dict:
    saisies = saisies or {}
    mod = cfg["modules"]
    split, caf = mod["loyer_nu_charges"], mod["caf"]
    idx_loc = ref_loc["idx"]

    ws = wb.create_sheet("Suivi")

    cols: list[dict] = [
        {"key": "locataire", "titre": "Locataire", "w": 24, "kind": "input"},
        {"key": "annee", "titre": "Année", "w": 8, "kind": "input", "fmt": "0"},
        {"key": "mois", "titre": "Mois", "w": 11, "kind": "input"},
    ]
    if split:
        cols += [
            {"key": "loyer_nu_du", "titre": "Loyer nu dû", "w": 12, "kind": "lookup",
             "src": "loyer_nu", "fmt": FMT_EURO},
            {"key": "charges_du", "titre": "Charges dues", "w": 12, "kind": "lookup",
             "src": "charges", "fmt": FMT_EURO},
            {"key": "total_du", "titre": "Total dû", "w": 12, "kind": "calc", "fmt": FMT_EURO},
        ]
    else:
        cols.append(
            {"key": "total_du", "titre": "Loyer dû", "w": 12, "kind": "lookup",
             "src": "loyer_total", "fmt": FMT_EURO})
    if caf:
        cols += [
            {"key": "caf_attendu", "titre": "CAF attendue", "w": 12, "kind": "lookup",
             "src": "part_caf", "fmt": FMT_EURO},
            {"key": "caf_recu", "titre": "CAF reçue", "w": 12, "kind": "input", "fmt": FMT_EURO},
            {"key": "caf_date", "titre": "Date CAF", "w": 12, "kind": "input", "fmt": FMT_DATE},
        ]
    cols += [
        {"key": "rac_attendu", "titre": "Reste à charge attendu", "w": 16, "kind": "calc",
         "fmt": FMT_EURO},
        {"key": "loc_recu", "titre": "Part locataire reçue", "w": 16, "kind": "input",
         "fmt": FMT_EURO},
        {"key": "loc_date", "titre": "Date locataire", "w": 13, "kind": "input", "fmt": FMT_DATE},
        {"key": "total_recu", "titre": "Total reçu", "w": 12, "kind": "calc", "fmt": FMT_EURO},
        {"key": "ecart", "titre": "Écart", "w": 11, "kind": "calc", "fmt": FMT_EURO},
        {"key": "statut", "titre": "Statut", "w": 16, "kind": "calc"},
    ]

    L = {c["key"]: get_column_letter(i) for i, c in enumerate(cols, 1)}
    col_de = {c["key"]: i for i, c in enumerate(cols, 1)}

    for i, c in enumerate(cols, 1):
        style_entete(ws.cell(row=1, column=i, value=c["titre"]))
        ws.column_dimensions[get_column_letter(i)].width = c["w"]

    def vlookup(src_field: str, row: int) -> str:
        return (f"=IFERROR(VLOOKUP(${L['locataire']}{row},RefLocataires,"
                f"{idx_loc[src_field]},FALSE),0)")

    r = 2
    for loc in cfg["locataires"]:
        nom = loc.get("nom")
        for (annee, m) in _mois_actifs(loc, cfg["annee_debut"], cfg["annee_fin"]):
            nom_mois = MOIS[m - 1]
            preserve = saisies.get((str(nom), int(annee), nom_mois), {})
            for c in cols:
                key = c["key"]
                cell = ws.cell(row=r, column=col_de[key])
                if key == "locataire":
                    cell.value = nom
                elif key == "annee":
                    cell.value = annee
                elif key == "mois":
                    cell.value = nom_mois
                elif c["kind"] == "lookup":
                    cell.value = vlookup(c["src"], r)
                elif key == "total_du" and split:
                    cell.value = f"={L['loyer_nu_du']}{r}+{L['charges_du']}{r}"
                elif key == "rac_attendu":
                    cell.value = (f"={L['total_du']}{r}-{L['caf_attendu']}{r}" if caf
                                  else f"={L['total_du']}{r}")
                elif key == "total_recu":
                    cell.value = (f"={L['caf_recu']}{r}+{L['loc_recu']}{r}" if caf
                                  else f"={L['loc_recu']}{r}")
                elif key == "ecart":
                    cell.value = f"={L['total_recu']}{r}-{L['total_du']}{r}"
                elif key == "statut":
                    tr, ec = f"{L['total_recu']}{r}", f"{L['ecart']}{r}"
                    cell.value = (f'=IF({tr}=0,"À encaisser",'
                                  f'IF(ABS({ec})<=0.005,"Soldé",'
                                  f'IF({ec}>0,"Trop-perçu","Partiel")))')
                elif key in COLS_SAISIE and key in preserve:
                    cell.value = preserve[key]   # réinjection des saisies existantes

                style_cellule(
                    cell, fmt=c.get("fmt"),
                    saisie=c["kind"] == "input" and key not in ("locataire", "annee", "mois"),
                    calc=c["kind"] in ("calc", "lookup"))
            r += 1

    derniere = max(r - 1, 2)

    der_col = get_column_letter(len(cols))
    table = Table(displayName="TblSuivi", ref=f"A1:{der_col}{derniere}")
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleLight9", showRowStripes=True, showColumnStripes=False)
    ws.add_table(table)
    ws.freeze_panes = "D2"

    dv = DataValidation(type="list", formula1="=LocatairesListe", allow_blank=True)
    dv.add(f"{L['locataire']}2:{L['locataire']}{derniere}")
    ws.add_data_validation(dv)

    plage_statut = f"{L['statut']}2:{L['statut']}{derniere}"
    for texte, couleur in (("Soldé", COUL_SOLDE), ("Trop-perçu", COUL_TROP),
                           ("Partiel", COUL_PARTIEL), ("À encaisser", COUL_ATTENTE)):
        ws.conditional_formatting.add(
            plage_statut,
            FormulaRule(formula=[f'${L["statut"]}2="{texte}"'],
                        fill=PatternFill("solid", fgColor=couleur)))

    def nommer(nom_plage: str, key: str) -> None:
        col = L[key]
        wb.defined_names.add(DefinedName(
            nom_plage, attr_text=f"Suivi!${col}$2:${col}${derniere}"))

    nommer("Suivi_Locataire", "locataire")
    nommer("Suivi_TotalDu", "total_du")
    nommer("Suivi_TotalRecu", "total_recu")
    nommer("Suivi_LocRecu", "loc_recu")
    if caf:
        nommer("Suivi_CAFRecue", "caf_recu")

    return {"cols": L, "derniere": derniere}


# --------------------------------------------------------------------------- #
# Onglet Bilan (synthèse SUMIFS)
# --------------------------------------------------------------------------- #

def construire_bilan(wb: Workbook, cfg: dict) -> None:
    caf = cfg["modules"]["caf"]
    locs = cfg["locataires"]
    ws = wb.create_sheet("Bilan")

    cols = [("Locataire", 24), ("Total dû", 14), ("Total reçu", 14)]
    if caf:
        cols += [("dont CAF", 13), ("dont locataire", 14)]
    cols += [("Solde", 14), ("Taux recouvrement", 16)]

    for i, (titre, w) in enumerate(cols, 1):
        style_entete(ws.cell(row=1, column=i, value=titre))
        ws.column_dimensions[get_column_letter(i)].width = w

    keys = ["nom", "du", "recu"] + (["caf", "loc"] if caf else []) + ["solde", "taux"]
    B = {k: get_column_letter(i) for i, k in enumerate(keys, 1)}
    pos = {k: i for i, k in enumerate(keys, 1)}

    for r, loc in enumerate(locs, start=2):
        nomc = f"${B['nom']}{r}"
        ws.cell(row=r, column=1, value=loc.get("nom"))
        ws.cell(row=r, column=pos["du"],
                value=f"=SUMIFS(Suivi_TotalDu,Suivi_Locataire,{nomc})")
        ws.cell(row=r, column=pos["recu"],
                value=f"=SUMIFS(Suivi_TotalRecu,Suivi_Locataire,{nomc})")
        if caf:
            ws.cell(row=r, column=pos["caf"],
                    value=f"=SUMIFS(Suivi_CAFRecue,Suivi_Locataire,{nomc})")
            ws.cell(row=r, column=pos["loc"],
                    value=f"=SUMIFS(Suivi_LocRecu,Suivi_Locataire,{nomc})")
        ws.cell(row=r, column=pos["solde"], value=f"={B['recu']}{r}-{B['du']}{r}")
        ws.cell(row=r, column=pos["taux"],
                value=f'=IFERROR({B["recu"]}{r}/{B["du"]}{r},"")')

    der = len(locs) + 1
    total_r = der + 1
    ws.cell(row=total_r, column=1, value="TOTAL").font = Font(bold=True)
    for k in (["du", "recu"] + (["caf", "loc"] if caf else []) + ["solde"]):
        col = B[k]
        c = ws.cell(row=total_r, column=pos[k], value=f"=SUM({col}2:{col}{der})")
        c.font = Font(bold=True)
    c = ws.cell(row=total_r, column=pos["taux"],
                value=f'=IFERROR({B["recu"]}{total_r}/{B["du"]}{total_r},"")')
    c.font = Font(bold=True)

    for r in range(2, total_r + 1):
        for k in keys:
            cell = ws.cell(row=r, column=pos[k])
            if k == "taux":
                cell.number_format = FMT_PCT
            elif k != "nom":
                cell.number_format = FMT_EURO
            cell.border = BORDURE

    plage_solde = f"{B['solde']}2:{B['solde']}{der}"
    ws.conditional_formatting.add(
        plage_solde, FormulaRule(formula=[f"${B['solde']}2<-0.005"],
                                 fill=PatternFill("solid", fgColor=COUL_PARTIEL)))
    ws.conditional_formatting.add(
        plage_solde, FormulaRule(formula=[f"${B['solde']}2>0.005"],
                                 fill=PatternFill("solid", fgColor=COUL_TROP)))
    ws.freeze_panes = "A2"


# --------------------------------------------------------------------------- #
# Onglet Guide
# --------------------------------------------------------------------------- #

def construire_guide(wb: Workbook, cfg: dict) -> None:
    ws = wb.create_sheet("Guide", 0)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 36
    ws.column_dimensions["C"].width = 64

    b = cfg["bailleur"]
    ws["B2"] = "Suivi des loyers"
    ws["B2"].font = Font(bold=True, size=18, color=COUL_ENTETE)
    ws["B3"] = f"Bailleur : {b.get('nom', '')}"
    ws["B3"].font = Font(bold=True, size=12)
    coord = " · ".join(str(b[k]) for k in ("adresse", "tel", "email") if b.get(k))
    if coord:
        ws["B4"] = coord

    r = 6
    sections = [
        ("Mode d'emploi", [
            ("1.", "Onglet « Locataires » : vérifiez les loyers de référence (loyer, "
                   "charges, part CAF, dépôt) et les dates d'entrée / sortie."),
            ("2.", "Onglet « Suivi » : chaque mois, saisissez les montants REÇUS dans les "
                   "cellules jaunes (CAF reçue, part locataire reçue, dates)."),
            ("3.", "Les montants attendus, totaux, écarts et statuts se calculent tout seuls "
                   "(cellules bleutées — ne pas y toucher)."),
            ("4.", "Onglet « Bilan » : synthèse automatique par locataire (dû, reçu, solde, "
                   "taux de recouvrement)."),
        ]),
    ]
    for titre, items in sections:
        ws.cell(row=r, column=2, value=titre).font = Font(bold=True, size=13, color=COUL_ENTETE)
        r += 1
        for a, t in items:
            ws.cell(row=r, column=2, value=a).font = Font(bold=True)
            ws.cell(row=r, column=3, value=t)
            r += 1
        r += 1

    ws.cell(row=r, column=2, value="Légende des statuts").font = Font(
        bold=True, size=13, color=COUL_ENTETE)
    r += 1
    for nom, couleur, desc in (
        ("Soldé", COUL_SOLDE, "Le total reçu couvre le total dû."),
        ("Trop-perçu", COUL_TROP, "Reçu supérieur au dû (avance, régularisation à prévoir)."),
        ("Partiel", COUL_PARTIEL, "Reçu inférieur au dû (impayé partiel)."),
        ("À encaisser", COUL_ATTENTE, "Aucun paiement saisi pour ce mois."),
    ):
        cell = ws.cell(row=r, column=2, value=nom)
        cell.fill = PatternFill("solid", fgColor=couleur)
        cell.alignment = Alignment(horizontal="center")
        cell.border = BORDURE
        ws.cell(row=r, column=3, value=desc)
        r += 1

    r += 1
    ws.cell(row=r, column=2, value="Modules actifs").font = Font(
        bold=True, size=13, color=COUL_ENTETE)
    r += 1
    for k, libelle in (
        ("loyer_nu_charges", "Loyer nu / charges séparés"),
        ("caf", "Suivi de la part CAF (tiers payant)"),
        ("depot_garantie", "Dépôt de garantie"),
        ("quittances", "Quittances (à venir)"),
        ("irl", "Révision IRL (à venir)"),
        ("regularisation_charges", "Régularisation des charges (à venir)"),
    ):
        ws.cell(row=r, column=2,
                value=("✓" if cfg["modules"].get(k) else "—")).alignment = Alignment(
            horizontal="center")
        ws.cell(row=r, column=3, value=libelle)
        r += 1


# --------------------------------------------------------------------------- #
# Préservation des saisies existantes
# --------------------------------------------------------------------------- #

# Titres d'en-tête du Suivi -> clé interne des colonnes de saisie.
_TITRE_VERS_KEY = {
    "CAF reçue": "caf_recu", "Date CAF": "caf_date",
    "Part locataire reçue": "loc_recu", "Date locataire": "loc_date",
}


def recolter_saisies(chemin_xlsx: Path) -> dict:
    """Lit un classeur existant et renvoie les saisies utilisateur à préserver.

    Clé : (nom_locataire, année, nom_mois) -> {caf_recu, caf_date, loc_recu, loc_date}.
    """
    chemin_xlsx = Path(chemin_xlsx)
    if not chemin_xlsx.is_file():
        return {}
    wb = load_workbook(chemin_xlsx, data_only=False)
    if "Suivi" not in wb.sheetnames:
        return {}
    ws = wb["Suivi"]

    entetes = {ws.cell(row=1, column=c).value: c for c in range(1, ws.max_column + 1)}
    c_loc, c_an, c_mo = (entetes.get("Locataire"), entetes.get("Année"), entetes.get("Mois"))
    if not all((c_loc, c_an, c_mo)):
        return {}

    saisies: dict = {}
    for r in range(2, ws.max_row + 1):
        nom = ws.cell(row=r, column=c_loc).value
        annee = ws.cell(row=r, column=c_an).value
        mois = ws.cell(row=r, column=c_mo).value
        if not nom or annee in (None, "") or not mois:
            continue
        valeurs = {}
        for titre, key in _TITRE_VERS_KEY.items():
            col = entetes.get(titre)
            if col:
                v = ws.cell(row=r, column=col).value
                if v not in (None, ""):
                    valeurs[key] = v
        if valeurs:
            saisies[(str(nom), int(annee), str(mois))] = valeurs
    return saisies


# --------------------------------------------------------------------------- #
# Orchestration
# --------------------------------------------------------------------------- #

def _slug(nom: str) -> str:
    s = re.sub(r"[^\w\-]+", "_", nom.strip(), flags=re.UNICODE)
    return s.strip("_") or "bailleur"


def generer_workbook(cfg: dict, sortie: Path, *, preserver: bool = True) -> Path:
    """Construit le classeur depuis une config validée et l'enregistre dans `sortie`.

    Si `preserver` et que `sortie` existe déjà, les saisies utilisateur sont reprises.
    """
    cfg = valider_config(cfg) if "annee_debut" not in cfg else cfg
    sortie = Path(sortie)

    saisies = recolter_saisies(sortie) if preserver else {}

    wb = Workbook()
    wb.remove(wb.active)
    ref_loc = construire_locataires(wb, cfg)
    construire_suivi(wb, cfg, ref_loc, saisies=saisies)
    construire_bilan(wb, cfg)
    construire_guide(wb, cfg)

    sortie.parent.mkdir(parents=True, exist_ok=True)
    wb.save(sortie)
    return sortie


def generer(chemin_config: Path, dossier_sortie: Path) -> Path:
    cfg = charger_config(chemin_config)
    sortie = Path(dossier_sortie) / f"Suivi_{_slug(cfg['bailleur']['nom'])}.xlsx"
    return generer_workbook(cfg, sortie)


def main(argv: list[str]) -> int:
    if len(argv) < 2:
        print(__doc__)
        return 2
    chemin_config = Path(argv[1])
    dossier_sortie = Path(argv[2]) if len(argv) > 2 else Path("sorties")
    if not chemin_config.is_file():
        print(f"Config introuvable : {chemin_config}", file=sys.stderr)
        return 1
    print(f"✔ Classeur généré : {generer(chemin_config, dossier_sortie)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
