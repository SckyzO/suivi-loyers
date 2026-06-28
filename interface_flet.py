#!/usr/bin/env python3
"""Interface graphique Flet (moteur de rendu Flutter) pour le générateur de
suivi des loyers — POC parallèle à `interface.py` (Tkinter).

Direction graphique : minimalisme professionnel. Police par défaut de Flet
(Roboto, embarquée dans le moteur, offline), palette neutre Material 3 + un accent
dérivé du thème du classeur, mise en page « app » (sidebar + zone Locataires).

Le moteur `generer_suivi_loyers` n'est pas touché : cette UI assemble une config
et appelle `valider_config` / `generer_workbook`.

Lancement (dev)        : python interface_flet.py
Packaging .exe unique  : flet pack interface_flet.py -n SuiviLoyers
"""
from __future__ import annotations

import datetime as dt
import json
import math
import threading
from pathlib import Path

import flet as ft

import generer_suivi_loyers as moteur

APP_TITRE = "Suivi des loyers"
APP_SOUS_TITRE = "Générateur de classeur Excel — un fichier par bailleur"
APP_VERSION = "0.1.0"
GITHUB_URL = ""   # placeholder (dépôt à créer) -> bouton désactivé tant que vide

# Échelle typographique unique (homogénéité). DISPLAY=titre appli, TITRE=section
# principale, SECTION=en-têtes majuscules, LABEL=libellés/sous-titres, CORPS=
# champs/cellules, PETIT=badges/puces.
TS_DISPLAY, TS_TITRE, TS_SECTION, TS_LABEL, TS_CORPS, TS_PETIT = 22, 18, 12, 12, 15, 11

# --- Registres lus depuis le moteur (source de vérité) ----------------------
POLICES = ["Tahoma", "Calibri", "Arial", "Verdana", "Segoe UI", "Georgia",
           "Times New Roman"]
MODES = [("comprises", "Loyer charges comprises"),
         ("separees", "Loyer + charges séparés"),
         ("sans", "Loyer seul (sans charges)")]
MODE_LABEL = dict(MODES)
MODE_KEY = {lbl: cle for cle, lbl in MODES}
def _libelle_theme(spec: dict) -> str:
    """« Nom — couleur » dans le menu déroulant (couleur si la teinte est définie)."""
    teinte = spec.get("teinte")
    return f'{spec["label"]} — {teinte}' if teinte else spec["label"]


THEME_LABEL = {tid: _libelle_theme(spec) for tid, spec in moteur.THEMES.items()}
THEME_KEY = {lbl: tid for tid, lbl in THEME_LABEL.items()}

APPARENCES = [("systeme", "Système"), ("clair", "Clair"), ("sombre", "Sombre")]
APPARENCE_ICONE = {"systeme": ft.Icons.BRIGHTNESS_AUTO,
                   "clair": ft.Icons.LIGHT_MODE,
                   "sombre": ft.Icons.DARK_MODE}
APPARENCE_MODE = {"systeme": ft.ThemeMode.SYSTEM,
                  "clair": ft.ThemeMode.LIGHT,
                  "sombre": ft.ThemeMode.DARK}

# --- Jetons de design (couleurs sémantiques M3 => clair/sombre auto) ---------
C_PANEL = ft.Colors.with_opacity(0.035, ft.Colors.ON_SURFACE)
C_FIELD = ft.Colors.SURFACE_CONTAINER_HIGHEST   # opaque (sinon menus transparents)
C_LINE = ft.Colors.OUTLINE_VARIANT
C_MUTED = ft.Colors.ON_SURFACE_VARIANT
C_HEAD = ft.Colors.with_opacity(0.05, ft.Colors.ON_SURFACE)
C_OK = ft.Colors.GREEN
C_PARTI = ft.Colors.AMBER_700
DLG_CHAMP = 300   # largeur fixe des champs du dialogue (modal statique)
RAYON = 12


def _parse_nombre(txt: str) -> float:
    """Coercition douce d'un montant saisi (virgule décimale, €, espaces)."""
    txt = (txt or "").strip().replace(",", ".").replace("€", "").replace(" ", "")
    if not txt:
        return 0.0
    val = float(txt)
    if not math.isfinite(val):
        raise ValueError("Montant invalide.")
    return val


def _seed_theme_classeur(theme_id: str) -> str:
    """Couleur d'accent de l'UI dérivée du thème du classeur sélectionné."""
    spec = moteur.THEMES.get(theme_id) or moteur.THEMES[moteur.THEME_DEFAUT]
    return "#" + spec["primaire"]


def _initiales(nom: str, prenom: str) -> str:
    a = (prenom or nom or "?").strip()[:1]
    b = (nom if prenom else (prenom or "")).strip()[:1]
    return (a + b).upper() or "?"


STYLE_BTN = ft.ButtonStyle(shape=ft.RoundedRectangleBorder(radius=RAYON))


def champ_texte(valeur="", label="", largeur=None, expand=False,
                icone=None) -> ft.TextField:
    """Champ texte au style maison (rempli, arrondi, compact)."""
    return ft.TextField(
        value="" if valeur is None else str(valeur), label=label, width=largeur,
        expand=expand, filled=True, bgcolor=C_FIELD, border_radius=RAYON,
        border_color=ft.Colors.TRANSPARENT, prefix_icon=icone,
        focused_border_color=ft.Colors.PRIMARY, dense=True,
        text_style=ft.TextStyle(size=15, weight=ft.FontWeight.W_500),
        content_padding=ft.padding.symmetric(12, 14),
        label_style=ft.TextStyle(size=TS_LABEL, color=C_MUTED))


def champ_liste(options, valeur=None, label="", expand=False,
                editable=False, largeur=None, icone=None) -> ft.Dropdown:
    return ft.Dropdown(
        label=label, value=valeur, expand=expand, width=largeur, filled=True,
        bgcolor=C_FIELD, prefix_icon=icone,
        border_radius=RAYON, border_color=ft.Colors.TRANSPARENT, editable=editable,
        focused_border_color=ft.Colors.PRIMARY, text_size=15, dense=True,
        content_padding=ft.padding.symmetric(12, 14),
        label_style=ft.TextStyle(size=TS_LABEL, color=C_MUTED),
        options=[ft.dropdown.Option(o) for o in options])


def titre_section(texte: str, icone=None) -> ft.Control:
    ligne = []
    if icone is not None:
        ligne.append(ft.Icon(icone, size=15, color=C_MUTED))
    ligne.append(ft.Text(texte.upper(), size=TS_SECTION, weight=ft.FontWeight.W_700,
                         color=C_MUTED, spans=None))
    return ft.Container(ft.Row(ligne, spacing=7, tight=True),
                        margin=ft.margin.only(top=6, bottom=2))


# ===========================================================================
#  Champ date : TextField (ISO) + icône calendrier (suffix) -> DatePicker natif
# ===========================================================================
class ChampDate(ft.Container):
    """Saisie de date optionnelle, avec sélecteur natif Material en suffixe.

    Valeur vide autorisée. `get()` renvoie "" ou une date ISO.
    """

    def __init__(self, page: ft.Page, valeur: str = "", actif: bool = True,
                 largeur=None):
        super().__init__()
        self.page = page
        self._champ = ft.TextField(
            value=valeur, hint_text="AAAA-MM-JJ", width=largeur, filled=True,
            bgcolor=C_FIELD,
            border_radius=RAYON, border_color=ft.Colors.TRANSPARENT,
            focused_border_color=ft.Colors.PRIMARY, dense=True,
            text_style=ft.TextStyle(size=15, weight=ft.FontWeight.W_500),
            content_padding=ft.padding.symmetric(12, 14), disabled=not actif,
            # Icône cliquable compacte (un IconButton gonflerait la hauteur du champ).
            suffix=ft.Container(
                ft.Icon(ft.Icons.CALENDAR_MONTH, size=18, color=C_MUTED),
                on_click=self._ouvrir, tooltip="Choisir une date",
                padding=ft.padding.only(left=4, right=2)))
        self.content = self._champ

    def _ouvrir(self, _):
        try:
            courant = dt.date.fromisoformat((self._champ.value or "").strip())
        except ValueError:
            courant = None
        self.page.open(ft.DatePicker(
            first_date=dt.date(2000, 1, 1), last_date=dt.date(2100, 12, 31),
            value=courant, on_change=self._choisir))

    def _choisir(self, e):
        d = e.control.value
        if d is not None:
            self._champ.value = d.strftime("%Y-%m-%d")
            self._champ.update()

    def get(self) -> str:
        txt = (self._champ.value or "").strip()
        if not txt:
            return ""
        dt.date.fromisoformat(txt)  # lève ValueError si format invalide
        return txt

    def set_actif(self, actif: bool):
        self._champ.disabled = not actif
        if not actif:
            self._champ.value = ""
        self.update()


# ===========================================================================
#  Dialogue d'édition d'un locataire — 2 colonnes, sans défilement
# ===========================================================================
class DialogueLocataire:
    def __init__(self, page: ft.Page, modules: dict, adresses: list[str],
                 loc: dict | None, on_valider):
        self.page = page
        self.on_valider = on_valider
        loc = loc or {}

        mode = modules.get("mode_charges") or (
            "separees" if modules.get("loyer_nu_charges", True) else "sans")
        self.split = mode in ("comprises", "separees")
        self.caf = bool(modules.get("caf", True))
        self.depot = bool(modules.get("depot_garantie", True))

        self.nom = champ_texte(loc.get("nom", ""), "Nom *", largeur=DLG_CHAMP,
                               icone=ft.Icons.PERSON_OUTLINE)
        self.prenom = champ_texte(loc.get("prenom", ""), "Prénom", largeur=DLG_CHAMP)
        self.type_bien = champ_liste(
            moteur.TYPES_BIEN, loc.get("type_bien", moteur.TYPES_BIEN[0]),
            "Type de bien", largeur=DLG_CHAMP, icone=ft.Icons.HOME_WORK_OUTLINED)
        self.identifiant = champ_texte(loc.get("identifiant", ""),
                                       "N° appart. / Nom du bien", largeur=DLG_CHAMP)
        self.adresse = champ_texte(loc.get("adresse", ""), "Adresse du logement",
                                   largeur=DLG_CHAMP,
                                   icone=ft.Icons.PLACE_OUTLINED)
        # Suggestions = adresses déjà saisies (puces cliquables). Champ libre =>
        # le texte tapé est toujours conservé (le dropdown éditable Flet 0.28 ne
        # persiste pas une valeur hors options : c'était la régression).
        self._sugg_adresses = [a for a in dict.fromkeys(adresses) if a]

        if self.split:
            self.loyer_nu = champ_texte(loc.get("loyer_nu", ""), "Loyer nu (€)",
                                        largeur=DLG_CHAMP)
            self.charges = champ_texte(loc.get("charges", ""), "Charges (€)",
                                       largeur=DLG_CHAMP)
        else:
            self.loyer = champ_texte(loc.get("loyer", ""), "Loyer (€)",
                                     largeur=DLG_CHAMP)
        if self.caf:
            self.part_caf = champ_texte(loc.get("part_caf", ""),
                                        "Part CAF / APL (€)", largeur=DLG_CHAMP)
        if self.depot:
            self.depot_garantie = champ_texte(loc.get("depot_garantie", ""),
                                              "Dépôt de garantie (€)",
                                              largeur=DLG_CHAMP)

        parti = bool(loc.get("date_sortie"))
        self.date_entree = ChampDate(page, loc.get("date_entree", ""),
                                     largeur=DLG_CHAMP)
        self.parti = ft.Switch(label="Le locataire est parti", value=parti,
                               on_change=self._maj_depart)
        self.date_sortie = ChampDate(page, loc.get("date_sortie", ""), actif=parti,
                                     largeur=DLG_CHAMP)
        self.caution_rendue = ft.Switch(
            label="Caution rendue", value=bool(loc.get("caution_rendue", False)),
            disabled=not (parti and self.depot))
        obs = loc.get("observation", "")
        opts_obs = list(moteur.OBSERVATIONS)
        if obs and obs not in opts_obs:
            opts_obs.append(obs)
        self.observation = champ_liste(opts_obs, obs or None, "Motif de départ",
                                       largeur=DLG_CHAMP)
        self.observation.disabled = not parti
        # Champs propres au bail, repris sur les documents de ce locataire.
        self.date_bail = ChampDate(page, loc.get("date_bail", ""), largeur=DLG_CHAMP)
        self.mode_paiement = champ_texte(loc.get("mode_paiement", ""),
                                         "Mode de paiement", largeur=DLG_CHAMP)
        self.jour_echeance = champ_texte(loc.get("jour_echeance", ""),
                                         "Jour d'échéance (ex. 5)", largeur=DLG_CHAMP)

        self.dlg = ft.AlertDialog(
            modal=True,
            shape=ft.RoundedRectangleBorder(radius=18),
            bgcolor=ft.Colors.SURFACE, surface_tint_color=ft.Colors.TRANSPARENT,
            title=ft.Text("Locataire" if loc else "Nouveau locataire",
                          weight=ft.FontWeight.W_700),
            content=self._corps(),
            actions=[
                ft.TextButton("Annuler", on_click=self._fermer, style=STYLE_BTN),
                ft.FilledButton("Valider", icon=ft.Icons.CHECK,
                                on_click=self._valider, style=STYLE_BTN),
            ],
            actions_alignment=ft.MainAxisAlignment.END)

    def _corps(self) -> ft.Control:
        if self.split:
            loyer = [self.loyer_nu, self.charges]
        else:
            loyer = [self.loyer]
        if self.caf:
            loyer.append(self.part_caf)
        if self.depot:
            loyer.append(self.depot_garantie)

        items_g = [
            titre_section("Identité", ft.Icons.BADGE_OUTLINED),
            self.nom, self.prenom,
            titre_section("Logement", ft.Icons.HOME_OUTLINED),
            self.type_bien, self.identifiant, self.adresse,
        ]
        if self._sugg_adresses:
            items_g.append(ft.Row(
                [ft.Chip(ft.Text(a, size=TS_PETIT),
                         on_click=lambda e, v=a: self._remplir_adresse(v))
                 for a in self._sugg_adresses[:6]],
                wrap=True, spacing=6, run_spacing=4))
        items_g += [
            titre_section("Documents (bail)", ft.Icons.RECEIPT_LONG_OUTLINED),
            ft.Text("Date du bail", size=TS_LABEL, color=C_MUTED),
            self.date_bail, self.mode_paiement, self.jour_echeance,
        ]
        gauche = ft.Column(items_g, spacing=12, tight=True,
                           scroll=ft.ScrollMode.AUTO)
        droite = ft.Column([
            titre_section("Loyer & charges", ft.Icons.EURO_ROUNDED), *loyer,
            titre_section("Bail", ft.Icons.DESCRIPTION_OUTLINED),
            ft.Text("Date d'entrée", size=TS_LABEL, color=C_MUTED),
            self.date_entree, self.parti,
            ft.Text("Date de sortie", size=TS_LABEL, color=C_MUTED),
            self.date_sortie, self.caution_rendue, self.observation,
        ], spacing=12, tight=True, scroll=ft.ScrollMode.AUTO)

        # Hauteur bornée + colonnes défilantes : le dialogue ne déborde jamais,
        # quel que soit le nombre de champs actifs selon les modules.
        return ft.Container(
            width=720, height=520,
            content=ft.Row([
                ft.Container(gauche, width=334),
                ft.VerticalDivider(width=1, color=C_LINE),
                ft.Container(droite, width=334),
            ], spacing=18, vertical_alignment=ft.CrossAxisAlignment.STRETCH))

    def _remplir_adresse(self, valeur: str):
        self.adresse.value = valeur
        self.adresse.update()

    def _maj_depart(self, _):
        parti = self.parti.value
        self.date_sortie.set_actif(parti)
        self.caution_rendue.disabled = not (parti and self.depot)
        self.observation.disabled = not parti
        self.caution_rendue.update()
        self.observation.update()

    def ouvrir(self):
        self.page.open(self.dlg)

    def _fermer(self, _):
        self.page.close(self.dlg)

    def _erreur(self, msg: str):
        self.page.open(ft.SnackBar(ft.Text(msg), bgcolor=ft.Colors.ERROR))

    def _valider(self, _):
        nom = self.nom.value.strip()
        if not nom:
            self._erreur("Le nom du locataire est obligatoire.")
            return
        data = {
            "nom": nom, "prenom": self.prenom.value.strip(),
            "type_bien": self.type_bien.value or moteur.TYPES_BIEN[0],
            "identifiant": self.identifiant.value.strip(),
            "adresse": (self.adresse.value or "").strip(),
        }
        try:
            if self.split:
                data["loyer_nu"] = _parse_nombre(self.loyer_nu.value)
                data["charges"] = _parse_nombre(self.charges.value)
            else:
                data["loyer"] = _parse_nombre(self.loyer.value)
            if self.caf:
                data["part_caf"] = _parse_nombre(self.part_caf.value)
            if self.depot:
                data["depot_garantie"] = _parse_nombre(self.depot_garantie.value)
            data["date_entree"] = self.date_entree.get()
            parti = self.parti.value
            data["date_sortie"] = self.date_sortie.get() if parti else ""
            data["date_bail"] = self.date_bail.get()
        except ValueError as exc:
            self._erreur(str(exc) or "Saisie invalide.")
            return
        data["caution_rendue"] = bool(self.caution_rendue.value) if parti else False
        data["observation"] = (self.observation.value or "").strip() if parti else ""
        data["mode_paiement"] = (self.mode_paiement.value or "").strip()
        data["jour_echeance"] = (self.jour_echeance.value or "").strip()
        self.page.close(self.dlg)
        self.on_valider(data)


# ===========================================================================
#  Application
# ===========================================================================
class AppLoyers:
    def __init__(self, page: ft.Page):
        self.page = page
        self.locataires: list[dict] = []
        self.apparence = "systeme"
        self._recherche = ""          # filtre texte de la table
        self._tri_cle = "nom"         # colonne de tri active
        self._tri_asc = True          # sens du tri
        page.title = APP_TITRE
        page.window.width = 1180
        page.window.height = 800
        page.window.min_width = 1000
        page.window.min_height = 680
        page.padding = 0
        self._appliquer_apparence()
        page.add(self._construire())
        self._rafraichir_table()

    # --- thème / apparence -------------------------------------------------
    def _theme_id_courant(self) -> str:
        champ = getattr(self, "theme", None)
        if champ is not None and champ.value:
            return THEME_KEY.get(champ.value, moteur.THEME_DEFAUT)
        return moteur.THEME_DEFAUT

    def _theme(self, seed: str) -> ft.Theme:
        # Police par défaut de Flet (Roboto, embarquée dans le moteur, offline).
        # text_theme : épaissit TOUT le texte par défaut (switches, table, libellés…),
        # pas seulement les champs. Les titres gardent leur poids explicite (W_700+).
        w5 = ft.FontWeight.W_500
        return ft.Theme(
            color_scheme_seed=seed,
            text_theme=ft.TextTheme(
                body_large=ft.TextStyle(size=15, weight=w5),
                body_medium=ft.TextStyle(size=14, weight=w5),
                body_small=ft.TextStyle(size=13, weight=w5),
                label_large=ft.TextStyle(size=14, weight=w5),
                label_medium=ft.TextStyle(size=13, weight=w5)))

    def _appliquer_apparence(self):
        self.page.theme_mode = APPARENCE_MODE[self.apparence]
        seed = _seed_theme_classeur(self._theme_id_courant())
        self.page.theme = self._theme(seed)
        self.page.dark_theme = self._theme(seed)
        self.page.update()

    def _changer_apparence(self, cle: str):
        self.apparence = cle
        self._maj_cartes_apparence()
        self._appliquer_apparence()

    def _carte_apparence(self, cle: str, label: str):
        """Carte cliquable d'apparence (icône + libellé). La sélection est rendue
        par _maj_cartes_apparence (bordure + fond teinté primaire)."""
        carte = ft.Container(
            on_click=lambda e, k=cle: self._changer_apparence(k),
            border_radius=RAYON, expand=True, ink=True,
            padding=ft.padding.symmetric(vertical=12, horizontal=8),
            alignment=ft.alignment.center,
            content=ft.Column(
                [ft.Icon(APPARENCE_ICONE[cle], size=22),
                 ft.Text(label, size=TS_CORPS, weight=ft.FontWeight.W_500)],
                spacing=6, tight=True,
                horizontal_alignment=ft.CrossAxisAlignment.CENTER))
        self._cartes_apparence[cle] = carte
        return carte

    def _maj_cartes_apparence(self):
        for cle, carte in self._cartes_apparence.items():
            actif = cle == self.apparence
            carte.bgcolor = (ft.Colors.PRIMARY_CONTAINER if actif
                             else ft.Colors.SURFACE_CONTAINER_HIGHEST)
            carte.border = ft.border.all(
                2, ft.Colors.PRIMARY if actif else ft.Colors.TRANSPARENT)
            coul = (ft.Colors.ON_PRIMARY_CONTAINER if actif
                    else ft.Colors.ON_SURFACE_VARIANT)
            carte.content.controls[0].color = coul
            carte.content.controls[1].color = coul

    def _changer_theme_classeur(self, _):
        self._appliquer_apparence()

    def _maj_sci(self, _):
        self.b_sci_nom.disabled = not self.b_sci.value
        self.b_sci_nom.update()

    # --- assemblage de la config ------------------------------------------
    def _modules_courants(self) -> dict:
        return {
            "mode_charges": MODE_KEY.get(self.mode_charges.value, "comprises"),
            "caf": self.m_caf.value, "depot_garantie": self.m_depot.value,
            "documents": self.m_docs.value, "tableau_bord": self.m_tdb.value,
            "regularisation_charges": self.m_regul.value, "irl": self.m_irl.value,
        }

    def _config(self) -> dict:
        return {
            "version": moteur.CONFIG_VERSION,
            "bailleur": {
                "nom": self.b_nom.value.strip(),
                "prenom": self.b_prenom.value.strip(),
                "sci": self.b_sci.value, "sci_nom": self.b_sci_nom.value.strip(),
                "adresse": self.b_adresse.value.strip(),
                "tel": self.b_tel.value.strip(),
                "email": self.b_email.value.strip(),
                "iban": self.b_iban.value.strip(),
            },
            "periode": {
                "annee_debut": int(self.an_debut.value or 0),
                "annee_fin": int(self.an_fin.value or 0),
            },
            "modules": self._modules_courants(),
            "theme": self._theme_id_courant(),
            "police": self.police.value or moteur.POLICE_DEFAUT,
            "style_excel": self.style_excel.value,
            "locataires": self.locataires,
        }

    # --- locataires --------------------------------------------------------
    def _adresses_connues(self) -> list[str]:
        return [loc.get("adresse", "") for loc in self.locataires if loc.get("adresse")]

    def _ajouter_loc(self, _):
        DialogueLocataire(self.page, self._modules_courants(),
                          self._adresses_connues(), None,
                          self._on_loc_valide(None)).ouvrir()

    def _modifier_loc(self, idx: int):
        DialogueLocataire(self.page, self._modules_courants(),
                          self._adresses_connues(), self.locataires[idx],
                          self._on_loc_valide(idx)).ouvrir()

    def _on_loc_valide(self, idx):
        def cb(data):
            if idx is None:
                self.locataires.append(data)
            else:
                self.locataires[idx] = data
            self._rafraichir_table()
        return cb

    def _supprimer_loc(self, idx: int):
        del self.locataires[idx]
        self._rafraichir_table()

    def split_actuel(self) -> bool:
        return MODE_KEY.get(self.mode_charges.value, "comprises") in (
            "comprises", "separees")

    def _badge(self, texte: str, couleur) -> ft.Control:
        return ft.Container(
            ft.Text(texte, size=TS_PETIT, weight=ft.FontWeight.W_600, color=couleur,
                    no_wrap=True, text_align=ft.TextAlign.CENTER),
            width=84, alignment=ft.alignment.center,
            bgcolor=ft.Colors.with_opacity(0.15, couleur),
            padding=ft.padding.symmetric(4, 6), border_radius=20)

    # --- table custom (colonnes flexibles : remplit la largeur, vs DataTable
    #     qui se dimensionne à la largeur intrinsèque de ses colonnes) ---------
    @staticmethod
    def _mois_an(iso: str) -> str:
        """Date ISO -> « MM/AAAA » (granularité utile pour une période de bail)."""
        iso = (iso or "").strip()
        if not iso:
            return ""
        try:
            return dt.date.fromisoformat(iso).strftime("%m/%Y")
        except ValueError:
            return iso

    def _nom_complet(self, loc: dict) -> str:
        return " ".join(x for x in (loc.get("nom", ""), loc.get("prenom", "")) if x)

    def _loyer_loc(self, loc: dict) -> float:
        if self.split_actuel():
            return (loc.get("loyer_nu", 0) or 0) + (loc.get("charges", 0) or 0)
        return loc.get("loyer", 0) or 0

    @staticmethod
    def _eur(valeur) -> str:
        return f"{(valeur or 0):,.0f} €".replace(",", " ")

    def _periode_txt(self, loc: dict) -> str:
        e = self._mois_an(loc.get("date_entree", ""))
        s = self._mois_an(loc.get("date_sortie", ""))
        if e and s:
            return f"{e} → {s}"
        if e:
            return f"depuis {e}"
        if s:
            return f"→ {s}"
        return "—"

    def _texte_recherche(self, loc: dict) -> str:
        return " ".join(
            str(loc.get(k, "") or "")
            for k in ("nom", "prenom", "identifiant", "type_bien", "adresse")).lower()

    def _colonnes(self) -> list[dict]:
        """Colonnes actives selon les modules. `flex` = poids `expand` ; `tri` =
        clé de tri (None => colonne non triable, ex. Actions)."""
        cols = [
            {"cle": "nom", "label": "Locataire", "flex": 24,
             "tri": lambda loc: self._nom_complet(loc).lower()},
            {"cle": "bien", "label": "Bien", "flex": 16,
             "tri": lambda loc: (loc.get("identifiant") or loc.get("type_bien") or "").lower()},
            {"cle": "adresse", "label": "Adresse", "flex": 22,
             "tri": lambda loc: (loc.get("adresse") or "").lower()},
            {"cle": "loyer", "label": "Loyer", "flex": 14, "droite": True,
             "tri": self._loyer_loc},
        ]
        if self.m_caf.value:
            cols.append({"cle": "caf", "label": "CAF", "flex": 12, "droite": True,
                         "tri": lambda loc: loc.get("part_caf", 0) or 0})
        cols += [
            {"cle": "periode", "label": "Période", "flex": 20,
             "tri": lambda loc: loc.get("date_entree") or ""},
            {"cle": "etat", "label": "État", "flex": 14,
             "tri": lambda loc: 1 if loc.get("date_sortie") else 0},
            {"cle": "actions", "label": "", "flex": 12, "droite": True, "tri": None},
        ]
        return cols

    def _locataires_affiches(self) -> list[tuple[int, dict]]:
        """(index d'origine, locataire) filtrés par recherche et triés. L'index
        d'origine reste celui de self.locataires (modifier/supprimer corrects)."""
        items = list(enumerate(self.locataires))
        q = self._recherche.strip().lower()
        if q:
            items = [(i, loc) for i, loc in items if q in self._texte_recherche(loc)]
        col = next((c for c in self._cols_courantes if c["cle"] == self._tri_cle), None)
        if col and col.get("tri"):
            items.sort(key=lambda it: col["tri"](it[1]), reverse=not self._tri_asc)
        return items

    def _trier_par(self, cle: str):
        if self._tri_cle == cle:
            self._tri_asc = not self._tri_asc
        else:
            self._tri_cle, self._tri_asc = cle, True
        self._rafraichir_table()

    def _entete_table(self) -> ft.Control:
        cells = []
        for col in self._cols_courantes:
            actif = col.get("tri") and self._tri_cle == col["cle"]
            coul = ft.Colors.PRIMARY if actif else C_MUTED
            enfants = [ft.Text(col["label"].upper(), size=TS_SECTION,
                               weight=ft.FontWeight.W_700, color=coul, no_wrap=True)]
            if actif:
                fleche = (ft.Icons.ARROW_UPWARD if self._tri_asc
                          else ft.Icons.ARROW_DOWNWARD)
                enfants.append(ft.Icon(fleche, size=14, color=ft.Colors.PRIMARY))
            rang = ft.Row(
                enfants, spacing=4, tight=True,
                alignment=(ft.MainAxisAlignment.END if col.get("droite")
                           else ft.MainAxisAlignment.START),
                vertical_alignment=ft.CrossAxisAlignment.CENTER)
            cell = ft.Container(rang, expand=col["flex"],
                                padding=ft.padding.symmetric(horizontal=10))
            if col.get("tri"):
                cell.on_click = lambda e, c=col["cle"]: self._trier_par(c)
                cell.ink, cell.border_radius, cell.tooltip = True, 8, "Trier"
            cells.append(cell)
        return ft.Container(ft.Row(cells, spacing=0), bgcolor=C_HEAD,
                            padding=ft.padding.symmetric(vertical=12))

    def _cellule(self, col: dict, loc: dict, orig_idx: int) -> ft.Control:
        cle, droite = col["cle"], col.get("droite", False)
        if cle == "nom":
            contenu = ft.Row([
                ft.CircleAvatar(
                    content=ft.Text(
                        _initiales(loc.get("nom", ""), loc.get("prenom", "")),
                        size=12, weight=ft.FontWeight.W_700),
                    radius=15, bgcolor=ft.Colors.PRIMARY_CONTAINER,
                    color=ft.Colors.ON_PRIMARY_CONTAINER),
                ft.Text(self._nom_complet(loc) or "—", weight=ft.FontWeight.W_600,
                        no_wrap=True, overflow=ft.TextOverflow.ELLIPSIS, expand=True),
            ], spacing=10, vertical_alignment=ft.CrossAxisAlignment.CENTER)
        elif cle == "etat":
            contenu = (self._badge("Parti", C_PARTI) if loc.get("date_sortie")
                       else self._badge("Présent", C_OK))
        elif cle == "actions":
            contenu = ft.Row([
                ft.IconButton(ft.Icons.EDIT_OUTLINED, icon_size=18, tooltip="Modifier",
                              on_click=lambda e, k=orig_idx: self._modifier_loc(k)),
                ft.IconButton(ft.Icons.DELETE_OUTLINE, icon_size=18,
                              icon_color=ft.Colors.ERROR, tooltip="Supprimer",
                              on_click=lambda e, k=orig_idx: self._supprimer_loc(k)),
            ], spacing=0, tight=True, alignment=ft.MainAxisAlignment.END)
        else:
            if cle == "bien":
                txt, mute = loc.get("identifiant") or loc.get("type_bien") or "—", True
            elif cle == "adresse":
                txt, mute = loc.get("adresse") or "—", True
            elif cle == "loyer":
                txt, mute = self._eur(self._loyer_loc(loc)), False
            elif cle == "caf":
                v = loc.get("part_caf", 0) or 0
                txt, mute = (self._eur(v) if v else "—"), True
            else:  # periode
                txt, mute = self._periode_txt(loc), True
            contenu = ft.Text(
                txt, no_wrap=True, overflow=ft.TextOverflow.ELLIPSIS,
                color=C_MUTED if mute else None,
                weight=None if mute else ft.FontWeight.W_600,
                text_align=ft.TextAlign.RIGHT if droite else ft.TextAlign.LEFT,
                tooltip=txt if cle == "adresse" and txt != "—" else None)
        aligne = None if isinstance(contenu, ft.Text) else (
            ft.alignment.center_right if droite else ft.alignment.center_left)
        return ft.Container(contenu, expand=col["flex"], alignment=aligne,
                            padding=ft.padding.symmetric(horizontal=10))

    def _ligne(self, orig_idx: int, loc: dict) -> ft.Control:
        cells = [self._cellule(col, loc, orig_idx) for col in self._cols_courantes]
        return ft.Container(
            ft.Row(cells, spacing=0, vertical_alignment=ft.CrossAxisAlignment.CENTER),
            padding=ft.padding.symmetric(vertical=8),
            border=ft.border.only(bottom=ft.BorderSide(1, C_LINE)),
            on_hover=self._hover_ligne)

    @staticmethod
    def _hover_ligne(e):
        e.control.bgcolor = (ft.Colors.with_opacity(0.04, ft.Colors.ON_SURFACE)
                             if e.data == "true" else None)
        e.control.update()

    def _on_recherche(self, e):
        self._recherche = e.control.value or ""
        self._rafraichir_table()

    def _rafraichir_table(self):
        self._cols_courantes = self._colonnes()
        n = len(self.locataires)
        if n:
            affiches = self._locataires_affiches()
            if affiches:
                corps = ft.Column([self._ligne(i, loc) for i, loc in affiches],
                                  scroll=ft.ScrollMode.AUTO, expand=True, spacing=0)
                self._zone_loc.content = ft.Column(
                    [self._entete_table(), corps], spacing=0, expand=True)
            else:
                self._zone_loc.content = self._aucun_resultat()
            n_aff = len(affiches)
            self._compteur.value = (f"{n_aff} / {n}" if self._recherche.strip()
                                    and n_aff != n
                                    else f"{n} locataire{'s' if n > 1 else ''}")
        else:
            self._zone_loc.content = self._etat_vide()
            self._compteur.value = "0 locataire"
        self.page.update()

    def _etat_vide(self) -> ft.Control:
        return ft.Container(
            ft.Column([
                ft.Icon(ft.Icons.GROUPS_OUTLINED, size=52, color=C_MUTED),
                ft.Text("Aucun locataire", size=16, weight=ft.FontWeight.W_600),
                ft.Text("Ajoutez votre premier locataire pour commencer.",
                        color=C_MUTED),
                ft.Container(height=6),
                ft.FilledButton("Ajouter un locataire", icon=ft.Icons.ADD,
                                on_click=self._ajouter_loc, style=STYLE_BTN),
            ], horizontal_alignment=ft.CrossAxisAlignment.CENTER, spacing=6,
                tight=True),
            alignment=ft.alignment.center, expand=True)

    def _aucun_resultat(self) -> ft.Control:
        return ft.Container(
            ft.Column([
                ft.Icon(ft.Icons.SEARCH_OFF, size=48, color=C_MUTED),
                ft.Text("Aucun résultat", size=16, weight=ft.FontWeight.W_600),
                ft.Text("Aucun locataire ne correspond à la recherche.",
                        color=C_MUTED),
            ], horizontal_alignment=ft.CrossAxisAlignment.CENTER, spacing=6,
                tight=True),
            alignment=ft.alignment.center, expand=True)

    # --- charger / enregistrer config -------------------------------------
    def _charger_config(self, _):
        self._fp_open.pick_files(
            dialog_title="Charger une configuration",
            allowed_extensions=["json", "yaml", "yml"], allow_multiple=False)

    def _on_open_result(self, e: ft.FilePickerResultEvent):
        if not e.files:
            return
        p = Path(e.files[0].path)
        try:
            if p.suffix.lower() in (".yaml", ".yml"):
                import yaml
                brut = yaml.safe_load(p.read_text(encoding="utf-8"))
            else:
                brut = json.loads(p.read_text(encoding="utf-8"))
            cfg, avertis = moteur.migrer_config(brut)
        except Exception as exc:  # noqa: BLE001 — remonter la vraie cause à l'UI
            self._toast(f"Chargement impossible : {exc}", ok=False)
            return
        self._peupler_depuis(cfg)
        msg = "Configuration chargée."
        if avertis:
            msg += " Adaptations : " + " ; ".join(avertis)
        self._toast(msg)

    def _peupler_depuis(self, cfg: dict):
        b = cfg.get("bailleur", {}) or {}
        self.b_nom.value = b.get("nom", "")
        self.b_prenom.value = b.get("prenom", "")
        self.b_sci.value = bool(b.get("sci", False))
        self.b_sci_nom.value = b.get("sci_nom", "")
        self.b_sci_nom.disabled = not self.b_sci.value
        self.b_adresse.value = b.get("adresse", "")
        self.b_tel.value = b.get("tel", "")
        self.b_email.value = b.get("email", "")
        self.b_iban.value = b.get("iban", "")
        per = cfg.get("periode", {}) or {}
        self.an_debut.value = str(per.get("annee_debut", cfg.get("annee_debut", "")))
        self.an_fin.value = str(per.get("annee_fin", cfg.get("annee_fin", "")))
        mod = cfg.get("modules", {}) or {}
        self.mode_charges.value = MODE_LABEL.get(
            mod.get("mode_charges", "comprises"), MODE_LABEL["comprises"])
        self.m_caf.value = bool(mod.get("caf", True))
        self.m_depot.value = bool(mod.get("depot_garantie", True))
        self.m_docs.value = bool(mod.get("documents", True))
        self.m_tdb.value = bool(mod.get("tableau_bord", True))
        self.m_regul.value = bool(mod.get("regularisation_charges", False))
        self.m_irl.value = bool(mod.get("irl", False))
        self.theme.value = THEME_LABEL.get(
            cfg.get("theme", moteur.THEME_DEFAUT), THEME_LABEL[moteur.THEME_DEFAUT])
        self.police.value = cfg.get("police", moteur.POLICE_DEFAUT)
        self.style_excel.value = bool(cfg.get("style_excel", True))
        self.locataires = list(cfg.get("locataires", []) or [])
        self._appliquer_apparence()
        self._rafraichir_table()

    def _enregistrer_config(self, _):
        self._fp_save_cfg.save_file(
            dialog_title="Enregistrer la configuration",
            file_name="config_loyers.json", allowed_extensions=["json"])

    def _on_save_cfg(self, e: ft.FilePickerResultEvent):
        if not e.path:
            return
        p = Path(e.path)
        if p.suffix.lower() != ".json":
            p = p.with_suffix(".json")
        try:
            p.write_text(json.dumps(self._config(), ensure_ascii=False, indent=2),
                         encoding="utf-8")
        except Exception as exc:  # noqa: BLE001
            self._toast(f"Enregistrement impossible : {exc}", ok=False)
            return
        self._toast(f"Configuration enregistrée : {p.name}")

    # --- génération du classeur -------------------------------------------
    def _generer(self, _):
        try:
            cfg = moteur.valider_config(self._config())
        except Exception as exc:  # noqa: BLE001 — erreurs de validation -> UI
            self._toast(str(exc), ok=False)
            return
        self._cfg_valide = cfg
        defaut = "Suivi_" + (moteur.base_slug(cfg["bailleur"]) or "bailleur") + ".xlsx"
        self._fp_xlsx.save_file(
            dialog_title="Générer le classeur", file_name=defaut,
            allowed_extensions=["xlsx"])

    def _on_save_xlsx(self, e: ft.FilePickerResultEvent):
        if not e.path:
            return
        sortie = Path(e.path)
        if sortie.suffix.lower() != ".xlsx":
            sortie = sortie.with_suffix(".xlsx")
        self._btn_generer.disabled = True
        self._btn_generer.text = "Génération…"
        self.page.update()
        threading.Thread(target=self._generer_async, args=(sortie,),
                         daemon=True).start()

    def _generer_async(self, sortie: Path):
        try:
            moteur.generer_workbook(self._cfg_valide, sortie)
            if self._enregistrer_aussi.value:
                sortie.with_suffix(".json").write_text(
                    json.dumps(self._config(), ensure_ascii=False, indent=2),
                    encoding="utf-8")
            msg, ok = f"Classeur généré : {sortie.name}", True
        except Exception as exc:  # noqa: BLE001
            msg, ok = f"Échec de la génération : {exc}", False
        self._btn_generer.disabled = False
        self._btn_generer.text = "Générer le classeur"
        self._toast(msg, ok=ok)

    def _toast(self, msg: str, ok: bool = True):
        self.page.open(ft.SnackBar(ft.Text(msg),
                                   bgcolor=None if ok else ft.Colors.ERROR))
        self.page.update()

    def _aide(self, texte: str) -> ft.Control:
        """« ? » d'aide : infobulle au survol ET bandeau au clic (tactile/clic)."""
        return ft.Container(
            ft.Icon(ft.Icons.HELP_OUTLINE, size=16, color=C_MUTED),
            tooltip=texte, on_click=lambda e: self._toast(texte),
            padding=2, border_radius=20)

    def _avec_aide(self, control: ft.Control, texte: str) -> ft.Control:
        return ft.Row([control, self._aide(texte)], spacing=6,
                      vertical_alignment=ft.CrossAxisAlignment.CENTER)

    # --- construction de l'UI ---------------------------------------------
    def _carte(self, titre: str, contenu: list, icone=None) -> ft.Control:
        return ft.Container(
            ft.Column([titre_section(titre, icone), *contenu], spacing=10,
                      tight=True),
            padding=ft.padding.only(bottom=6))

    def _build_reglages(self):
        """Crée les contrôles déportés (apparence, thème, police) + le dialogue
        Réglages ouvert depuis l'engrenage de l'en-tête."""
        self._cartes_apparence = {}
        rangee_apparence = ft.Row(
            [self._carte_apparence(k, lbl) for k, lbl in APPARENCES],
            spacing=10)
        self._maj_cartes_apparence()
        self.theme = champ_liste(list(THEME_LABEL.values()),
                                 THEME_LABEL[moteur.THEME_DEFAUT], "Thème",
                                 expand=True, icone=ft.Icons.PALETTE_OUTLINED)
        self.theme.on_change = self._changer_theme_classeur
        self.police = champ_liste(POLICES, moteur.POLICE_DEFAUT, "Police",
                                  expand=True,
                                  icone=ft.Icons.FONT_DOWNLOAD_OUTLINED)
        self.style_excel = ft.Switch(label="Générer pour Microsoft Excel", value=True)
        lien_gh = ft.OutlinedButton(
            "Voir sur GitHub", icon=ft.Icons.CODE, style=STYLE_BTN,
            on_click=self._ouvrir_github, disabled=not GITHUB_URL)
        contenu = ft.Container(width=420, content=ft.Column([
            titre_section("Apparence de l'application",
                          ft.Icons.BRIGHTNESS_6_OUTLINED),
            rangee_apparence,
            titre_section("Apparence du classeur", ft.Icons.PALETTE_OUTLINED),
            self.theme, self.police,
            self._avec_aide(self.style_excel,
                "Applique aux graphiques du tableau de bord le style natif de Microsoft "
                "Excel (barres pleines, coins arrondis) — recommandé si vous ouvrez le "
                "fichier dans Excel. Décochez pour un rendu neutre (ex. LibreOffice)."),
            titre_section("À propos", ft.Icons.INFO_OUTLINE),
            ft.Text(f"{APP_TITRE} — version {APP_VERSION}", size=TS_CORPS,
                    weight=ft.FontWeight.W_500),
            lien_gh,
        ], spacing=12, tight=True))
        self._dlg_reglages = ft.AlertDialog(
            modal=True, shape=ft.RoundedRectangleBorder(radius=18),
            bgcolor=ft.Colors.SURFACE, surface_tint_color=ft.Colors.TRANSPARENT,
            title=ft.Text("Réglages", weight=ft.FontWeight.W_700),
            content=contenu,
            actions=[ft.TextButton(
                "Fermer", style=STYLE_BTN,
                on_click=lambda e: self.page.close(self._dlg_reglages))],
            actions_alignment=ft.MainAxisAlignment.END)

    def _ouvrir_reglages(self, _):
        self.page.open(self._dlg_reglages)

    def _ouvrir_github(self, _):
        if GITHUB_URL:
            self.page.launch_url(GITHUB_URL)

    def _entete(self) -> ft.Control:
        marque = ft.Container(
            ft.Icon(ft.Icons.RECEIPT_LONG_ROUNDED, color=ft.Colors.ON_PRIMARY,
                    size=22),
            width=44, height=44, border_radius=12, bgcolor=ft.Colors.PRIMARY,
            alignment=ft.alignment.center)
        return ft.Container(
            ft.Row([
                ft.Row([marque, ft.Column([
                    ft.Text(APP_TITRE, size=TS_DISPLAY, weight=ft.FontWeight.W_800),
                    ft.Text(APP_SOUS_TITRE, size=TS_LABEL, color=C_MUTED),
                ], spacing=1, tight=True)], spacing=14),
                ft.IconButton(ft.Icons.SETTINGS_OUTLINED, tooltip="Réglages",
                              on_click=self._ouvrir_reglages),
            ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
            padding=ft.padding.symmetric(16, 24))

    def _sidebar(self) -> ft.Control:
        self.b_nom = champ_texte(label="Nom *", expand=True)
        self.b_prenom = champ_texte(label="Prénom", expand=True)
        self.b_sci = ft.Switch(label="SCI", value=False, on_change=self._maj_sci)
        self.b_sci_nom = champ_texte(label="Nom de la SCI", expand=True)
        self.b_sci_nom.disabled = True   # grisé tant que SCI décoché (sans décaler)
        self.b_adresse = champ_texte(label="Adresse", expand=True,
                                     icone=ft.Icons.PLACE_OUTLINED)
        self.b_tel = champ_texte(label="Téléphone", expand=True,
                                 icone=ft.Icons.PHONE_OUTLINED)
        self.b_email = champ_texte(label="E-mail", expand=True,
                                   icone=ft.Icons.MAIL_OUTLINED)
        # IBAN du bailleur (compte de réception) repris sur les documents.
        self.b_iban = champ_texte(label="IBAN", expand=True,
                                  icone=ft.Icons.ACCOUNT_BALANCE_OUTLINED)

        self.an_debut = champ_texte("2024", "Année de début", expand=True)
        self.an_fin = champ_texte("2026", "Année de fin", expand=True)

        self.mode_charges = champ_liste(
            [lbl for _, lbl in MODES], MODE_LABEL["comprises"], "Mode de charges",
            expand=True, icone=ft.Icons.TUNE)
        self.mode_charges.on_change = lambda e: self._rafraichir_table()
        self.m_caf = ft.Switch(label="Part CAF / APL", value=True,
                               on_change=lambda e: self._rafraichir_table())
        self.m_depot = ft.Switch(label="Dépôt de garantie", value=True)
        self.m_docs = ft.Switch(label="Documents (quittance, avis…)", value=True)
        self.m_tdb = ft.Switch(label="Tableau de bord", value=True)
        self.m_regul = ft.Switch(label="Régularisation des charges", value=False)
        self.m_irl = ft.Switch(label="Révision IRL", value=False)

        contenu = ft.Column([
            self._carte("Bailleur", [
                ft.Row([self.b_nom, self.b_prenom], spacing=10),
                ft.Row([self.b_sci, self.b_sci_nom], spacing=10,
                       vertical_alignment=ft.CrossAxisAlignment.CENTER),
                self.b_adresse,
                ft.Row([self.b_tel, self.b_email], spacing=10),
                self.b_iban,
            ], icone=ft.Icons.PERSON_OUTLINE),
            ft.Divider(height=1, color=C_LINE),
            self._carte("Période", [
                ft.Row([self.an_debut, self.an_fin], spacing=10),
            ], icone=ft.Icons.CALENDAR_MONTH),
            ft.Divider(height=1, color=C_LINE),
            self._carte("Contenu du classeur", [
                self._avec_aide(self.mode_charges,
                          "Comprises : loyer TTC, charges incluses. Séparés : loyer "
                          "nu + provision charges détaillés. Sans : loyer seul."),
                self._avec_aide(self.m_caf, "Part du loyer versée directement par la "
                          "CAF/MSA (tiers payant). Décochez si le locataire paie tout."),
                self._avec_aide(self.m_depot, "Suivi du dépôt de garantie versé à l'entrée "
                          "et de sa restitution."),
                self._avec_aide(self.m_docs, "Quittances, avis d'échéance et lettres de "
                          "relance imprimables."),
                self._avec_aide(self.m_tdb, "Onglet de graphiques de synthèse."),
                self._avec_aide(self.m_regul, "Onglet annuel : provisions encaissées vs "
                          "charges réelles → solde à régulariser."),
                self._avec_aide(self.m_irl, "Révision annuelle du loyer selon l'indice IRL "
                          "(INSEE). Onglet de saisie des indices."),
            ], icone=ft.Icons.WIDGETS_OUTLINED),
        ], spacing=14, scroll=ft.ScrollMode.AUTO, expand=True)

        return ft.Container(
            contenu, width=400, bgcolor=C_PANEL,
            padding=ft.padding.symmetric(18, 22),
            border=ft.border.only(right=ft.BorderSide(1, C_LINE)))

    def _zone_principale(self) -> ft.Control:
        self._compteur = ft.Text("0 locataire", size=TS_LABEL, color=C_MUTED)
        self._zone_loc = ft.Container(expand=True)

        cadre = ft.Container(
            self._zone_loc, expand=True, bgcolor=ft.Colors.SURFACE,
            border=ft.border.all(1, C_LINE), border_radius=RAYON,
            padding=0, clip_behavior=ft.ClipBehavior.HARD_EDGE)

        self._champ_recherche = ft.TextField(
            hint_text="Rechercher…", width=260, filled=True, bgcolor=C_FIELD,
            border_radius=RAYON, border_color=ft.Colors.TRANSPARENT,
            focused_border_color=ft.Colors.PRIMARY, dense=True,
            prefix_icon=ft.Icons.SEARCH, on_change=self._on_recherche,
            text_style=ft.TextStyle(size=14),
            content_padding=ft.padding.symmetric(8, 14))

        barre = ft.Row([
            ft.Row([ft.Text("Locataires", size=TS_TITRE, weight=ft.FontWeight.W_700),
                    self._compteur], spacing=12,
                   vertical_alignment=ft.CrossAxisAlignment.CENTER),
            ft.Row([self._champ_recherche,
                    ft.FilledButton("Ajouter", icon=ft.Icons.ADD,
                                    on_click=self._ajouter_loc, style=STYLE_BTN)],
                   spacing=12, vertical_alignment=ft.CrossAxisAlignment.CENTER),
        ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
            vertical_alignment=ft.CrossAxisAlignment.CENTER)

        return ft.Container(
            ft.Column([barre, cadre], spacing=14, expand=True),
            expand=True, padding=ft.padding.symmetric(18, 24))

    def _pied(self) -> ft.Control:
        self._enregistrer_aussi = ft.Switch(label="Enregistrer la config à côté",
                                            value=False)
        self._btn_generer = ft.FilledButton(
            "Générer le classeur", icon=ft.Icons.TABLE_VIEW_ROUNDED,
            on_click=self._generer,
            style=ft.ButtonStyle(padding=ft.padding.symmetric(14, 22),
                                 shape=ft.RoundedRectangleBorder(radius=RAYON)))
        return ft.Container(
            ft.Row([
                ft.OutlinedButton("Charger config", icon=ft.Icons.FOLDER_OPEN,
                                  on_click=self._charger_config, style=STYLE_BTN),
                ft.OutlinedButton("Enregistrer config", icon=ft.Icons.SAVE_OUTLINED,
                                  on_click=self._enregistrer_config, style=STYLE_BTN),
                self._enregistrer_aussi,
                ft.Container(expand=True),
                self._btn_generer,
            ], vertical_alignment=ft.CrossAxisAlignment.CENTER, spacing=12),
            padding=ft.padding.symmetric(12, 24),
            border=ft.border.only(top=ft.BorderSide(1, C_LINE)))

    def _construire(self) -> ft.Control:
        self._fp_open = ft.FilePicker(on_result=self._on_open_result)
        self._fp_save_cfg = ft.FilePicker(on_result=self._on_save_cfg)
        self._fp_xlsx = ft.FilePicker(on_result=self._on_save_xlsx)
        self.page.overlay.extend([self._fp_open, self._fp_save_cfg, self._fp_xlsx])
        self._build_reglages()

        corps = ft.Row([self._sidebar(), self._zone_principale()],
                       spacing=0, expand=True,
                       vertical_alignment=ft.CrossAxisAlignment.STRETCH)
        return ft.Column([self._entete(), ft.Divider(height=1, color=C_LINE),
                          corps, self._pied()], spacing=0, expand=True)


def main(page: ft.Page):
    AppLoyers(page)


if __name__ == "__main__":
    ft.app(main)
