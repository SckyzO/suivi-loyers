#!/usr/bin/env python3
"""Moteur de génération des classeurs Excel de suivi des loyers (part locataire + part CAF).

Un fichier .xlsx par bailleur. Le contenu (colonnes, feuilles, documents) est piloté par les
`modules` activés.

Organisation du classeur :
  * Guide                 : mode d'emploi.
  * Locataires            : référentiel (nom, type/identifiant/adresse du bien, loyers, dépôt).
  * Une feuille PAR locataire : saisie mensuelle, nommée par le n° d'appartement (ou le nom de
    la maison). Lisible même avec beaucoup de locataires.
  * Données (masquée)     : consolide toutes les feuilles locataire par formules ; alimente le
    Bilan et les documents (SUMIFS sur plages nommées). Aucune double saisie.
  * Bilan                 : synthèse par locataire.
  * Quittance / Avis d'échéance / Lettre de relance : documents à imprimer (si module actif).

Particularités :
  * Les lignes ne couvrent que la période d'activité de chaque locataire (entrée -> sortie) :
    gère les rotations fréquentes.
  * Régénération sur un fichier existant : les montants déjà saisis (CAF reçue, part locataire,
    dates) sont préservés (clé = locataire + année + mois).

Usage CLI / Docker :  python generer_suivi_loyers.py <config.yaml> [dossier_sortie]
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
# Constantes
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
COUL_SOLDE = "C6EFCE"
COUL_TROP = "FFEB9C"
COUL_PARTIEL = "FFC7CE"
COUL_ATTENTE = "E7E6E6"

_THIN = Side(style="thin", color="BFBFBF")
BORDURE = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)

TYPES_BIEN = ["Appartement", "Maison"]

# Colonnes de saisie utilisateur à préserver lors d'une régénération.
COLS_SAISIE = ("caf_recu", "caf_date", "loc_recu", "loc_date")

# Feuilles « système » (tout le reste = une feuille locataire).
FEUILLES_SYSTEME = {"Guide", "Locataires", "Données", "Bilan", "Régularisation charges",
                    "Quittance", "Avis d'échéance", "Lettre de relance"}

MODULES_DEFAUT = {
    "loyer_nu_charges": True,
    "caf": True,
    "depot_garantie": True,
    "documents": True,               # quittance + avis d'échéance + lettre de relance
    "irl": False,                    # phase 2
    "regularisation_charges": False,  # phase 2
}


# --------------------------------------------------------------------------- #
# Config
# --------------------------------------------------------------------------- #

def valider_config(raw: dict) -> dict:
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
    # Rétro-compat : ancienne clé `quittances`.
    if "quittances" in (raw.get("modules") or {}):
        modules["documents"] = bool(raw["modules"]["quittances"])

    locataires = raw.get("locataires") or []
    if not locataires:
        raise ValueError("Il faut au moins un locataire.")
    for i, loc in enumerate(locataires, 1):
        if not loc.get("nom"):
            raise ValueError(f"Le locataire #{i} n'a pas de nom.")
        loc.setdefault("type_bien", TYPES_BIEN[0])
        # Compat : ancien champ unique `bien` -> identifiant si rien d'autre.
        if not loc.get("identifiant"):
            loc["identifiant"] = loc.get("bien") or loc["nom"]

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
    de = _date(loc.get("date_entree"))
    ds = _date(loc.get("date_sortie"))
    debut = (de.year, de.month) if de else (annee_debut, 1)
    fin = (ds.year, ds.month) if ds else (annee_fin, 12)
    return [(a, m) for a in range(annee_debut, annee_fin + 1)
            for m in range(1, 13) if debut <= (a, m) <= fin]


def _slug(nom: str) -> str:
    s = re.sub(r"[^\w\-]+", "_", nom.strip(), flags=re.UNICODE)
    return s.strip("_") or "bailleur"


_CAR_INTERDITS = set('[]:*?/\\')


def _nom_feuille(ident: str, pris: set) -> str:
    """Nom de feuille Excel valide et unique (<= 31 car., sans caractères interdits)."""
    base = "".join(c for c in str(ident) if c not in _CAR_INTERDITS).strip()[:31] or "Locataire"
    nom, i = base, 2
    while nom in pris:
        suff = f" ({i})"
        nom = base[:31 - len(suff)] + suff
        i += 1
    pris.add(nom)
    return nom


def _ref(feuille: str, cellule: str) -> str:
    """Référence inter-feuilles robuste (gère espaces et apostrophes)."""
    return "'" + feuille.replace("'", "''") + "'!" + cellule


# --------------------------------------------------------------------------- #
# Styles
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
# Onglet Locataires (référentiel)
# --------------------------------------------------------------------------- #

def construire_locataires(wb: Workbook, cfg: dict) -> dict:
    mod = cfg["modules"]
    split, caf, depot = mod["loyer_nu_charges"], mod["caf"], mod["depot_garantie"]

    ws = wb.create_sheet("Locataires")

    cols: list[tuple[str, str]] = [
        ("nom", "Nom / Prénom"),
        ("type_bien", "Type de bien"),
        ("identifiant", "N° appart. / Nom maison"),
        ("adresse", "Adresse du logement"),
    ]
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

    largeurs = {"nom": 22, "type_bien": 14, "identifiant": 22, "adresse": 30,
                "loyer_nu": 13, "charges": 13, "loyer_total": 14, "part_caf": 16,
                "reste": 16, "depot": 16, "date_entree": 13, "date_sortie": 13}
    for champ, l in lettre.items():
        ws.column_dimensions[l].width = largeurs.get(champ, 14)

    for r, loc in enumerate(cfg["locataires"], start=2):
        ws.cell(r, idx["nom"], loc.get("nom"))
        ws.cell(r, idx["type_bien"], loc.get("type_bien", TYPES_BIEN[0]))
        ws.cell(r, idx["identifiant"], loc.get("identifiant", ""))
        ws.cell(r, idx["adresse"], loc.get("adresse", ""))
        if split:
            ws.cell(r, idx["loyer_nu"], _num(loc, "loyer_nu") or 0)
            ws.cell(r, idx["charges"], _num(loc, "charges") or 0)
            ws.cell(r, idx["loyer_total"], f"={lettre['loyer_nu']}{r}+{lettre['charges']}{r}")
        else:
            ws.cell(r, idx["loyer_total"], _num(loc, "loyer", "loyer_total", "loyer_nu") or 0)
        if caf:
            ws.cell(r, idx["part_caf"], _num(loc, "part_caf") or 0)
            ws.cell(r, idx["reste"], f"={lettre['loyer_total']}{r}-{lettre['part_caf']}{r}")
        else:
            ws.cell(r, idx["reste"], f"={lettre['loyer_total']}{r}")
        if depot:
            ws.cell(r, idx["depot"], _num(loc, "depot_garantie") or 0)
        if (de := _date(loc.get("date_entree"))):
            ws.cell(r, idx["date_entree"], de)
        if (ds := _date(loc.get("date_sortie"))):
            ws.cell(r, idx["date_sortie"], ds)

    derniere = len(cfg["locataires"]) + 1

    fmt_par_champ = {"loyer_nu": FMT_EURO, "charges": FMT_EURO, "loyer_total": FMT_EURO,
                     "part_caf": FMT_EURO, "reste": FMT_EURO, "depot": FMT_EURO,
                     "date_entree": FMT_DATE, "date_sortie": FMT_DATE}
    for r in range(2, derniere + 1):
        for champ, i in idx.items():
            cell = ws.cell(r, i)
            est_formule = isinstance(cell.value, str) and cell.value.startswith("=")
            style_cellule(cell, fmt=fmt_par_champ.get(champ), calc=est_formule,
                          saisie=not est_formule and champ != "reste")

    # Validation : type de bien.
    dv = DataValidation(type="list", formula1='"%s"' % ",".join(TYPES_BIEN), allow_blank=True)
    dv.add(f"{lettre['type_bien']}2:{lettre['type_bien']}{derniere}")
    ws.add_data_validation(dv)

    table = Table(displayName="TblLocataires",
                  ref=f"A1:{lettre['date_sortie']}{derniere}")
    table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(table)
    ws.freeze_panes = "A2"

    der_col = lettre["date_sortie"]
    wb.defined_names.add(DefinedName("LocatairesListe",
                                     attr_text=f"Locataires!$A$2:$A${derniere}"))
    wb.defined_names.add(DefinedName("RefLocataires",
                                     attr_text=f"Locataires!$A$2:${der_col}${derniere}"))

    return {"idx": idx, "lettre": lettre, "n": len(cfg["locataires"])}


# --------------------------------------------------------------------------- #
# Une feuille de saisie par locataire
# --------------------------------------------------------------------------- #

def _colonnes_locataire(split: bool, caf: bool) -> list[dict]:
    cols = [
        {"key": "locataire", "titre": "Locataire", "w": 20, "kind": "meta", "cache": True},
        {"key": "annee", "titre": "Année", "w": 8, "kind": "meta", "fmt": "0"},
        {"key": "mois", "titre": "Mois", "w": 11, "kind": "meta"},
    ]
    if split:
        cols += [
            {"key": "loyer_nu_du", "titre": "Loyer nu dû", "w": 12, "kind": "ref",
             "src": "loyer_nu", "fmt": FMT_EURO},
            {"key": "charges_du", "titre": "Charges dues", "w": 12, "kind": "ref",
             "src": "charges", "fmt": FMT_EURO},
            {"key": "total_du", "titre": "Total dû", "w": 12, "kind": "calc", "fmt": FMT_EURO},
        ]
    else:
        cols.append({"key": "total_du", "titre": "Loyer dû", "w": 12, "kind": "ref",
                     "src": "loyer_total", "fmt": FMT_EURO})
    if caf:
        cols += [
            {"key": "caf_attendu", "titre": "CAF attendue", "w": 12, "kind": "ref",
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
    return cols


# Ligne d'en-tête du tableau dans une feuille locataire (après le titre).
PL_LIGNE_ENTETE = 4


def construire_feuilles_locataires(wb: Workbook, cfg: dict, ref_loc: dict,
                                   saisies: dict) -> list[dict]:
    mod = cfg["modules"]
    split, caf = mod["loyer_nu_charges"], mod["caf"]
    lettre_loc = ref_loc["lettre"]
    cols = _colonnes_locataire(split, caf)
    L = {c["key"]: get_column_letter(i) for i, c in enumerate(cols, 1)}
    col_de = {c["key"]: i for i, c in enumerate(cols, 1)}

    pris: set = set()
    infos: list[dict] = []

    for loc_index, loc in enumerate(cfg["locataires"]):
        nom = loc.get("nom")
        ident = loc.get("identifiant") or nom
        feuille = _nom_feuille(ident, pris)
        ws = wb.create_sheet(feuille)
        ws.sheet_view.showGridLines = False
        rloc = loc_index + 2  # ligne du locataire dans l'onglet Locataires

        # Titre + adresse.
        ws.cell(1, 2, f"{nom}").font = Font(bold=True, size=14, color=COUL_ENTETE)
        ws.cell(1, 4, str(ident)).font = Font(bold=True, size=12)
        ws.cell(2, 2, "Adresse :").font = Font(bold=True)
        ws.cell(2, 4, "=" + _ref("Locataires", f"${lettre_loc['adresse']}${rloc}"))

        # En-têtes du tableau.
        for c in cols:
            i = col_de[c["key"]]
            style_entete(ws.cell(PL_LIGNE_ENTETE, i, c["titre"]))
            ws.column_dimensions[get_column_letter(i)].width = c["w"]
            if c.get("cache"):
                ws.column_dimensions[get_column_letter(i)].hidden = True

        def refloc(field: str) -> str:
            return "=" + _ref("Locataires", f"${lettre_loc[field]}${rloc}")

        rows_map: dict = {}
        r = PL_LIGNE_ENTETE + 1
        for (annee, m) in _mois_actifs(loc, cfg["annee_debut"], cfg["annee_fin"]):
            nom_mois = MOIS[m - 1]
            preserve = saisies.get((str(nom), int(annee), nom_mois), {})
            for c in cols:
                key = c["key"]
                cell = ws.cell(r, col_de[key])
                if key == "locataire":
                    cell.value = nom
                elif key == "annee":
                    cell.value = annee
                elif key == "mois":
                    cell.value = nom_mois
                elif c["kind"] == "ref":
                    cell.value = refloc(c["src"])
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
                    cell.value = preserve[key]
                style_cellule(cell, fmt=c.get("fmt"),
                              saisie=c["kind"] == "input",
                              calc=c["kind"] in ("calc", "ref"))
            rows_map[(annee, m)] = r
            r += 1

        derniere = r - 1
        if derniere >= PL_LIGNE_ENTETE + 1:
            plage_statut = f"{L['statut']}{PL_LIGNE_ENTETE + 1}:{L['statut']}{derniere}"
            for texte, couleur in (("Soldé", COUL_SOLDE), ("Trop-perçu", COUL_TROP),
                                   ("Partiel", COUL_PARTIEL), ("À encaisser", COUL_ATTENTE)):
                ws.conditional_formatting.add(
                    plage_statut,
                    FormulaRule(formula=[f'${L["statut"]}{PL_LIGNE_ENTETE + 1}="{texte}"'],
                                fill=PatternFill("solid", fgColor=couleur)))
        ws.freeze_panes = ws.cell(PL_LIGNE_ENTETE + 1, 4).coordinate

        infos.append({"loc": loc, "nom": nom, "feuille": feuille, "cols": L,
                      "rows": rows_map})

    return infos


# --------------------------------------------------------------------------- #
# Feuille Données (consolidée, masquée) : alimente Bilan + documents
# --------------------------------------------------------------------------- #

def construire_donnees(wb: Workbook, cfg: dict, feuilles: list[dict]) -> None:
    mod = cfg["modules"]
    split, caf = mod["loyer_nu_charges"], mod["caf"]

    cols = ["locataire", "annee", "mois"]
    if split:
        cols += ["loyer_nu_du", "charges_du"]
    cols.append("total_du")
    if caf:
        cols += ["caf_attendu", "caf_recu"]
    cols += ["loc_recu", "loc_date", "total_recu"]
    pos = {k: i for i, k in enumerate(cols, 1)}

    ws = wb.create_sheet("Données")
    ws.sheet_state = "hidden"
    for i, k in enumerate(cols, 1):
        ws.cell(1, i, k)

    r = 2
    for info in feuilles:
        f, Lpl, rows = info["feuille"], info["cols"], info["rows"]
        for (annee, m), rpl in rows.items():
            ws.cell(r, pos["locataire"], info["nom"])
            ws.cell(r, pos["annee"], annee)
            ws.cell(r, pos["mois"], MOIS[m - 1])
            for k in cols[3:]:
                ws.cell(r, pos[k], "=" + _ref(f, f"{Lpl[k]}{rpl}"))
            r += 1
    derniere = max(r - 1, 2)

    def nommer(nom_plage: str, key: str) -> None:
        col = get_column_letter(pos[key])
        wb.defined_names.add(DefinedName(
            nom_plage, attr_text=f"Données!${col}$2:${col}${derniere}"))

    nommer("Suivi_Locataire", "locataire")
    nommer("Suivi_Annee", "annee")
    nommer("Suivi_Mois", "mois")
    nommer("Suivi_TotalDu", "total_du")
    nommer("Suivi_TotalRecu", "total_recu")
    nommer("Suivi_LocRecu", "loc_recu")
    nommer("Suivi_LocDate", "loc_date")
    if split:
        nommer("Suivi_LoyerNuDu", "loyer_nu_du")
        nommer("Suivi_ChargesDu", "charges_du")
    if caf:
        nommer("Suivi_CAFRecue", "caf_recu")


# --------------------------------------------------------------------------- #
# Onglet Bilan
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
        style_entete(ws.cell(1, i, titre))
        ws.column_dimensions[get_column_letter(i)].width = w

    keys = ["nom", "du", "recu"] + (["caf", "loc"] if caf else []) + ["solde", "taux"]
    B = {k: get_column_letter(i) for i, k in enumerate(keys, 1)}
    pos = {k: i for i, k in enumerate(keys, 1)}

    for r, loc in enumerate(locs, start=2):
        nomc = f"${B['nom']}{r}"
        ws.cell(r, 1, loc.get("nom"))
        ws.cell(r, pos["du"], f"=SUMIFS(Suivi_TotalDu,Suivi_Locataire,{nomc})")
        ws.cell(r, pos["recu"], f"=SUMIFS(Suivi_TotalRecu,Suivi_Locataire,{nomc})")
        if caf:
            ws.cell(r, pos["caf"], f"=SUMIFS(Suivi_CAFRecue,Suivi_Locataire,{nomc})")
            ws.cell(r, pos["loc"], f"=SUMIFS(Suivi_LocRecu,Suivi_Locataire,{nomc})")
        ws.cell(r, pos["solde"], f"={B['recu']}{r}-{B['du']}{r}")
        ws.cell(r, pos["taux"], f'=IFERROR({B["recu"]}{r}/{B["du"]}{r},"")')

    der = len(locs) + 1
    total_r = der + 1
    ws.cell(total_r, 1, "TOTAL").font = Font(bold=True)
    for k in (["du", "recu"] + (["caf", "loc"] if caf else []) + ["solde"]):
        col = B[k]
        c = ws.cell(total_r, pos[k], f"=SUM({col}2:{col}{der})")
        c.font = Font(bold=True)
    c = ws.cell(total_r, pos["taux"], f'=IFERROR({B["recu"]}{total_r}/{B["du"]}{total_r},"")')
    c.font = Font(bold=True)

    for r in range(2, total_r + 1):
        for k in keys:
            cell = ws.cell(r, pos[k])
            if k == "taux":
                cell.number_format = FMT_PCT
            elif k != "nom":
                cell.number_format = FMT_EURO
            cell.border = BORDURE

    plage_solde = f"{B['solde']}2:{B['solde']}{der}"
    ws.conditional_formatting.add(plage_solde, FormulaRule(
        formula=[f"${B['solde']}2<-0.005"], fill=PatternFill("solid", fgColor=COUL_PARTIEL)))
    ws.conditional_formatting.add(plage_solde, FormulaRule(
        formula=[f"${B['solde']}2>0.005"], fill=PatternFill("solid", fgColor=COUL_TROP)))
    ws.freeze_panes = "A2"


# --------------------------------------------------------------------------- #
# Documents à imprimer (quittance, avis d'échéance, lettre de relance)
# --------------------------------------------------------------------------- #

def construire_document(wb: Workbook, cfg: dict, ref_loc: dict, kind: str) -> None:
    mod = cfg["modules"]
    split, caf = mod["loyer_nu_charges"], mod["caf"]
    idx = ref_loc["idx"]
    annees = list(range(cfg["annee_debut"], cfg["annee_fin"] + 1))

    specs = {
        "quittance": {
            "feuille": "Quittance", "titre": "QUITTANCE DE LOYER",
            "base": "recu",
            "intro": "déclare avoir reçu de",
            "note": "Quittance valable uniquement si le loyer est intégralement réglé.",
        },
        "avis": {
            "feuille": "Avis d'échéance", "titre": "AVIS D'ÉCHÉANCE",
            "base": "du",
            "intro": "informe",
            "note": "Ce document ne peut tenir lieu de quittance.",
        },
        "relance": {
            "feuille": "Lettre de relance", "titre": "LETTRE DE RELANCE",
            "base": "du",
            "intro": "n'a pas reçu le paiement de",
            "note": "En cas de nouveau retard, des frais de relance pourront être facturés.",
        },
    }
    spec = specs[kind]

    ws = wb.create_sheet(spec["feuille"])
    ws.sheet_view.showGridLines = False
    for col, w in (("A", 3), ("B", 26), ("C", 26), ("D", 16), ("E", 16)):
        ws.column_dimensions[col].width = w

    ws.merge_cells("B2:E2")
    titre = ws["B2"]
    titre.value = spec["titre"]
    titre.font = Font(bold=True, size=18, color=COUL_ENTETE)
    titre.alignment = Alignment(horizontal="center")

    # Sélecteurs.
    ws["B4"], ws["C4"] = "Locataire :", cfg["locataires"][0]["nom"]
    ws["B5"], ws["C5"] = "Mois :", MOIS[0]
    ws["D5"], ws["E5"] = "Année :", annees[0]
    for lab in ("B4", "B5", "D5"):
        ws[lab].font = Font(bold=True)
    for sel in ("C4", "C5", "E5"):
        ws[sel].fill = PatternFill("solid", fgColor=COUL_SAISIE)
        ws[sel].border = BORDURE

    def valider(cellule: str, formule: str) -> None:
        dv = DataValidation(type="list", formula1=formule, allow_blank=False)
        dv.add(cellule)
        ws.add_data_validation(dv)

    valider("C4", "=LocatairesListe")
    valider("C5", '"%s"' % ",".join(MOIS))
    valider("E5", '"%s"' % ",".join(str(a) for a in annees))

    def vlook(field: str) -> str:
        return f'IFERROR(VLOOKUP($C$4,RefLocataires,{idx[field]},FALSE),"")'

    # Bloc bailleur (gauche) et locataire (droite).
    b = cfg["bailleur"]
    ws["B7"] = "Le bailleur :"
    ws["B7"].font = Font(bold=True)
    ws["B8"] = b.get("nom", "")
    for i, k in enumerate(("adresse", "tel", "email"), start=9):
        if b.get(k):
            ws.cell(i, 2, str(b[k]))

    ws["D7"] = "Le locataire :"
    ws["D7"].font = Font(bold=True)
    ws["D8"] = "=$C$4"
    ws["D9"] = f"={vlook('identifiant')}"
    ws["D10"] = f"={vlook('adresse')}"

    crit = "Suivi_Locataire,$C$4,Suivi_Annee,$E$5,Suivi_Mois,$C$5"

    def sif(plage: str) -> str:
        return f"SUMIFS({plage},{crit})"

    # Tableau des montants.
    r0 = 13
    lignes = []
    if spec["base"] == "recu":
        if split:
            lignes += [("Loyer nu", sif("Suivi_LoyerNuDu")),
                       ("Charges", sif("Suivi_ChargesDu"))]
        lignes.append(("Total loyer + charges dû", sif("Suivi_TotalDu")))
        if caf:
            lignes += [("Dont part CAF reçue", sif("Suivi_CAFRecue")),
                       ("Dont part locataire reçue", sif("Suivi_LocRecu"))]
        lignes.append(("Montant total reçu", sif("Suivi_TotalRecu")))
    else:  # avis / relance : montants dus
        if split:
            lignes += [("Loyer pour la période", sif("Suivi_LoyerNuDu")),
                       ("Charges (provisions)", sif("Suivi_ChargesDu"))]
        lignes.append(("Total à régler", sif("Suivi_TotalDu")))
        if spec["base"] == "du" and kind == "relance":
            lignes.append(("Déjà reçu", sif("Suivi_TotalRecu")))
            lignes.append(("Reste dû", f'{sif("Suivi_TotalDu")}-{sif("Suivi_TotalRecu")}'))

    montant_row = r0
    for i, (label, formule) in enumerate(lignes):
        r = r0 + i
        gras = label.startswith(("Montant total", "Total", "Reste dû"))
        ws.cell(r, 2, label).font = Font(bold=gras)
        cell = ws.cell(r, 3, f"={formule}")
        cell.number_format = FMT_EURO
        cell.border = BORDURE
        if gras:
            montant_row = r
    montant_cell = f"$C${montant_row}"

    fin = r0 + len(lignes)
    if kind == "quittance":
        r_date = fin
        ws.cell(r_date, 2, "Date de paiement").font = Font(bold=True)
        cd = ws.cell(r_date, 3,
                     f'=IF({sif("Suivi_TotalRecu")}=0,"",{sif("Suivi_LocDate")})')
        cd.number_format = FMT_DATE
        cd.border = BORDURE
        fin = r_date + 1

    # Corps + mention.
    r_corps = fin + 1
    ws.merge_cells(start_row=r_corps, start_column=2, end_row=r_corps + 2, end_column=5)
    if kind == "quittance":
        corps = (f'="Je soussigné(e) "&$B$8&", bailleur, déclare avoir reçu de "&$C$4'
                 f'&" la somme de "&TEXT({montant_cell},"0.00")&" € au titre du loyer et des '
                 f'charges pour la période de "&$C$5&" "&$E$5'
                 f'&", et lui en donne quittance, sous réserve de tous mes droits."')
    elif kind == "avis":
        corps = (f'="Madame, Monsieur, veuillez trouver le montant de votre loyer pour la '
                 f'période de "&$C$5&" "&$E$5&", soit "&TEXT({montant_cell},"0.00")'
                 f'&" €, à régler sous 8 jours."')
    else:
        corps = (f'="Madame, Monsieur, sauf erreur de notre part, le loyer de la période de "'
                 f'&$C$5&" "&$E$5&" reste impayé. Nous vous remercions de régulariser la somme '
                 f'de "&TEXT({montant_cell},"0.00")&" € sous 8 jours."')
    cc = ws.cell(r_corps, 2, corps)
    cc.alignment = Alignment(wrap_text=True, vertical="top")

    r_note = r_corps + 4
    ws.cell(r_note, 2, spec["note"]).font = Font(italic=True)
    r_sign = r_note + 2
    ws.cell(r_sign, 2, "Fait à ……………………………, le ……………………………")
    ws.cell(r_sign + 2, 2, "Signature du bailleur :").font = Font(bold=True)


def construire_documents(wb: Workbook, cfg: dict, ref_loc: dict) -> None:
    for kind in ("quittance", "avis", "relance"):
        construire_document(wb, cfg, ref_loc, kind)


# --------------------------------------------------------------------------- #
# Onglet Régularisation des charges (annuelle, par locataire)
# --------------------------------------------------------------------------- #

def construire_regularisation(wb: Workbook, cfg: dict, saisies_reg: dict) -> None:
    # Sans distinction loyer nu / charges, il n'y a pas de provisions à régulariser.
    if not (cfg["modules"].get("regularisation_charges") and cfg["modules"]["loyer_nu_charges"]):
        return

    ws = wb.create_sheet("Régularisation charges")
    titres = [("Locataire", 24), ("Année", 10), ("Provisions appelées (€)", 20),
              ("Charges réelles (€)", 18), ("Solde (€)", 14), ("Sens", 28)]
    for i, (t, w) in enumerate(titres, 1):
        style_entete(ws.cell(1, i, t))
        ws.column_dimensions[get_column_letter(i)].width = w

    r = 2
    for loc in cfg["locataires"]:
        nom = loc.get("nom")
        annees = sorted({a for (a, _) in _mois_actifs(loc, cfg["annee_debut"], cfg["annee_fin"])})
        for annee in annees:
            ws.cell(r, 1, nom).border = BORDURE
            ca = ws.cell(r, 2, annee)
            ca.number_format = "0"
            ca.border = BORDURE
            cp = ws.cell(r, 3, f"=SUMIFS(Suivi_ChargesDu,Suivi_Locataire,$A{r},Suivi_Annee,$B{r})")
            cp.number_format = FMT_EURO
            cp.border = BORDURE
            cr = ws.cell(r, 4)
            v = saisies_reg.get((str(nom), int(annee)))
            if v not in (None, ""):
                cr.value = v
            style_cellule(cr, saisie=True, fmt=FMT_EURO)
            cs = ws.cell(r, 5, f"=C{r}-D{r}")
            cs.number_format = FMT_EURO
            cs.border = BORDURE
            csens = ws.cell(r, 6, f'=IF(D{r}=0,"Charges réelles à saisir",'
                                  f'IF(ABS(C{r}-D{r})<=0.005,"Équilibré",'
                                  f'IF(C{r}>D{r},"À rembourser au locataire","Complément à demander")))')
            csens.border = BORDURE
            r += 1
    der = r - 1
    if der >= 2:
        ws.conditional_formatting.add(f"E2:E{der}", FormulaRule(
            formula=["$E2<-0.005"], fill=PatternFill("solid", fgColor=COUL_PARTIEL)))
        ws.conditional_formatting.add(f"E2:E{der}", FormulaRule(
            formula=["$E2>0.005"], fill=PatternFill("solid", fgColor=COUL_TROP)))
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
    items = [
        ("1.", "Onglet « Locataires » : vérifiez les biens (type, n° / nom, adresse), les "
               "loyers de référence et les dates d'entrée / sortie."),
        ("2.", "Une feuille par locataire : chaque mois, saisissez les montants REÇUS dans les "
               "cellules jaunes (CAF reçue, part locataire reçue, dates)."),
        ("3.", "Totaux, écarts et statuts se calculent automatiquement (cellules bleutées)."),
        ("4.", "Onglet « Bilan » : synthèse par locataire."),
        ("5.", "Documents à imprimer : choisissez le locataire et la période dans les listes "
               "déroulantes ; le document se remplit seul."),
    ]
    if cfg["modules"].get("regularisation_charges") and cfg["modules"]["loyer_nu_charges"]:
        items.append(
            ("6.", "Onglet « Régularisation charges » : saisissez les charges réelles annuelles ; "
                   "le solde par locataire (à rembourser ou à compléter) se calcule seul."))
    ws.cell(r, 2, "Mode d'emploi").font = Font(bold=True, size=13, color=COUL_ENTETE)
    r += 1
    for a, t in items:
        ws.cell(r, 2, a).font = Font(bold=True)
        ws.cell(r, 3, t)
        r += 1
    r += 1

    ws.cell(r, 2, "Légende des statuts").font = Font(bold=True, size=13, color=COUL_ENTETE)
    r += 1
    for nom, couleur, desc in (
        ("Soldé", COUL_SOLDE, "Le total reçu couvre le total dû."),
        ("Trop-perçu", COUL_TROP, "Reçu supérieur au dû."),
        ("Partiel", COUL_PARTIEL, "Reçu inférieur au dû (impayé partiel)."),
        ("À encaisser", COUL_ATTENTE, "Aucun paiement saisi pour ce mois."),
    ):
        cell = ws.cell(r, 2, nom)
        cell.fill = PatternFill("solid", fgColor=couleur)
        cell.alignment = Alignment(horizontal="center")
        cell.border = BORDURE
        ws.cell(r, 3, desc)
        r += 1


# --------------------------------------------------------------------------- #
# Préservation des saisies (lecture des feuilles locataire)
# --------------------------------------------------------------------------- #

_TITRE_VERS_KEY = {
    "CAF reçue": "caf_recu", "Date CAF": "caf_date",
    "Part locataire reçue": "loc_recu", "Date locataire": "loc_date",
}


def recolter_saisies(chemin_xlsx: Path) -> dict:
    """Saisies utilisateur d'un classeur existant : (nom, année, mois) -> {colonne: valeur}."""
    chemin_xlsx = Path(chemin_xlsx)
    if not chemin_xlsx.is_file():
        return {}
    wb = load_workbook(chemin_xlsx, data_only=False)
    saisies: dict = {}

    for nom_feuille in wb.sheetnames:
        if nom_feuille in FEUILLES_SYSTEME:
            continue
        ws = wb[nom_feuille]
        # Repérer la ligne d'en-tête (celle qui contient « Mois »).
        ligne_ent = None
        for r in range(1, min(ws.max_row, 12) + 1):
            valeurs = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
            if "Mois" in valeurs and "Année" in valeurs:
                ligne_ent = r
                entetes = {v: c for c, v in enumerate(valeurs, 1) if v}
                break
        if ligne_ent is None:
            continue
        c_loc, c_an, c_mo = (entetes.get("Locataire"), entetes.get("Année"), entetes.get("Mois"))
        if not all((c_loc, c_an, c_mo)):
            continue
        for r in range(ligne_ent + 1, ws.max_row + 1):
            nom = ws.cell(r, c_loc).value
            annee = ws.cell(r, c_an).value
            mois = ws.cell(r, c_mo).value
            if not nom or annee in (None, "") or not mois:
                continue
            valeurs = {}
            for titre, key in _TITRE_VERS_KEY.items():
                col = entetes.get(titre)
                if col:
                    v = ws.cell(r, col).value
                    if v not in (None, ""):
                        valeurs[key] = v
            if valeurs:
                saisies[(str(nom), int(annee), str(mois))] = valeurs
    return saisies


def recolter_regularisation(chemin_xlsx: Path) -> dict:
    """Charges réelles déjà saisies : (nom, année) -> montant."""
    chemin_xlsx = Path(chemin_xlsx)
    if not chemin_xlsx.is_file():
        return {}
    wb = load_workbook(chemin_xlsx, data_only=False)
    if "Régularisation charges" not in wb.sheetnames:
        return {}
    ws = wb["Régularisation charges"]
    ent = {ws.cell(1, c).value: c for c in range(1, ws.max_column + 1)}
    cL, cA, cR = ent.get("Locataire"), ent.get("Année"), ent.get("Charges réelles (€)")
    if not all((cL, cA, cR)):
        return {}
    res: dict = {}
    for r in range(2, ws.max_row + 1):
        nom, an, v = ws.cell(r, cL).value, ws.cell(r, cA).value, ws.cell(r, cR).value
        if nom and an not in (None, "") and v not in (None, ""):
            res[(str(nom), int(an))] = v
    return res


# --------------------------------------------------------------------------- #
# Orchestration
# --------------------------------------------------------------------------- #

def generer_workbook(cfg: dict, sortie: Path, *, preserver: bool = True) -> Path:
    cfg = valider_config(cfg) if "annee_debut" not in cfg else cfg
    sortie = Path(sortie)
    saisies = recolter_saisies(sortie) if preserver else {}
    saisies_reg = recolter_regularisation(sortie) if preserver else {}

    wb = Workbook()
    wb.remove(wb.active)

    ref_loc = construire_locataires(wb, cfg)
    feuilles = construire_feuilles_locataires(wb, cfg, ref_loc, saisies)
    construire_donnees(wb, cfg, feuilles)
    construire_bilan(wb, cfg)
    construire_regularisation(wb, cfg, saisies_reg)
    if cfg["modules"].get("documents"):
        construire_documents(wb, cfg, ref_loc)
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
