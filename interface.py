#!/usr/bin/env python3
"""Interface graphique pour générer les classeurs de suivi des loyers.

Destinée à un usage non technique (Windows) : on remplit un formulaire, on clique
« Générer le fichier Excel ». Aucune connaissance de Python n'est requise.

Lancement (développement) :  python interface.py
Distribution : packagée en .exe autonome via build.bat (PyInstaller).
"""

from __future__ import annotations

import json
import math
import datetime as dt
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import generer_suivi_loyers as moteur

# Calendrier optionnel : si tkcalendar est dispo, on ajoute un bouton 📅 ; la saisie
# clavier (AAAA-MM-JJ) reste toujours possible, calendrier ou pas.
try:
    from tkcalendar import Calendar
    HAS_CAL = True
except Exception:  # noqa: BLE001
    HAS_CAL = False

# Thème moderne optionnel (look Windows 11). Absent → repli sur le thème ttk « clam ».
# L'interface reste pleinement fonctionnelle dans les deux cas.
try:
    import sv_ttk
    HAS_SVTTK = True
except Exception:  # noqa: BLE001
    HAS_SVTTK = False

# Détection du mode clair/sombre du système (Windows/macOS/Linux). Optionnel : absent →
# le mode « Système » retombe sur clair.
try:
    import darkdetect
    HAS_DARKDETECT = True
except Exception:  # noqa: BLE001
    HAS_DARKDETECT = False

APP_TITRE = "Suivi des loyers — générateur"
ANNEE = dt.date.today().year
TYPES_BIEN = moteur.TYPES_BIEN
OBSERVATIONS = moteur.OBSERVATIONS

MODES = [("comprises", "Loyer charges comprises"),
         ("separees", "Loyer + charges séparés"),
         ("sans", "Loyer seul (sans charges)")]
MODE_KEY = {lbl: k for k, lbl in MODES}
MODE_LABEL = {k: lbl for k, lbl in MODES}

# Apparence : thèmes pilotés par le registre du moteur (label affiché <-> identifiant).
# Aucune liste en dur : ajouter un thème côté moteur le rend disponible ici.
THEME_LABEL = {tid: spec["label"] for tid, spec in moteur.THEMES.items()}
THEME_KEY = {lbl: tid for tid, lbl in THEME_LABEL.items()}
# Polices d'origine sur Windows (rendu garanti Excel + LibreOffice). Défaut = moteur.
POLICES = ["Tahoma", "Calibri", "Arial", "Verdana", "Segoe UI", "Georgia", "Times New Roman"]

# Affichage de la fenêtre (distinct du thème du classeur Excel). « Système » suit le mode
# clair/sombre de Windows ; n'a d'effet que si sv-ttk est présent.
WINDOW_THEMES = [("systeme", "Système"), ("clair", "Clair"), ("sombre", "Sombre")]
WINDOW_KEY = {lbl: k for k, lbl in WINDOW_THEMES}
WINDOW_LABEL = {k: lbl for k, lbl in WINDOW_THEMES}


def _parse_nombre(txt: str) -> float:
    txt = (txt or "").strip().replace(",", ".").replace("€", "").replace(" ", "")
    if not txt:
        return 0.0
    val = float(txt)
    if not math.isfinite(val):
        raise ValueError("Montant invalide.")
    return val


# --------------------------------------------------------------------------- #
# Champ date : calendrier (tkcalendar) ou saisie texte AAAA-MM-JJ en repli
# --------------------------------------------------------------------------- #

class ChampDate:
    """Saisie de date : champ texte (placeholder « AAAA-MM-JJ », tapable au clavier) + 📅."""

    _PH = "AAAA-MM-JJ"

    def __init__(self, parent, iso: str = ""):
        self.frame = ttk.Frame(parent)
        self.var = tk.StringVar()
        self.entry = ttk.Entry(self.frame, textvariable=self.var)
        self.entry.pack(side="left", fill="x", expand=True)
        self.bouton = None
        if HAS_CAL:
            self.bouton = ttk.Button(self.frame, text="📅", width=3,
                                     command=self._ouvrir, takefocus=False)
            self.bouton.pack(side="left", padx=(5, 0))
        # Couleur de texte normale (selon thème) vs placeholder grisé.
        dark = HAS_SVTTK and sv_ttk.get_theme() == "dark"
        self._fg = ttk.Style().lookup("TEntry", "foreground") or ("#f0f0f0" if dark else "#000000")
        self._ph_fg = "#8a8a8a"
        self._ph_on = False
        self.entry.bind("<FocusIn>", self._ph_hide)
        self.entry.bind("<FocusOut>", self._ph_show)
        self.set(iso or "")

    def _ph_show(self, _e=None) -> None:
        if not self.var.get().strip():
            self._ph_on = True
            self.entry.configure(foreground=self._ph_fg)
            self.var.set(self._PH)

    def _ph_hide(self, _e=None) -> None:
        if self._ph_on:
            self._ph_on = False
            self.var.set("")
            self.entry.configure(foreground=self._fg)

    def grid(self, **kw):
        self.frame.grid(**kw)

    def _ouvrir(self) -> None:
        # Calendrier en popup déroulant ancré SOUS le champ (pas une fenêtre séparée).
        # Sélection en un clic ; ferme au clic ailleurs ou sur Échap. Couleurs clair/sombre.
        dark = HAS_SVTTK and sv_ttk.get_theme() == "dark"
        if dark:
            couleurs = dict(background="#2b2b2b", foreground="#f0f0f0",
                            headersbackground="#333333", headersforeground="#f0f0f0",
                            normalbackground="#2b2b2b", normalforeground="#f0f0f0",
                            weekendbackground="#313131", weekendforeground="#f0f0f0",
                            othermonthbackground="#262626", othermonthwebackground="#262626",
                            selectbackground="#4cc2ff", selectforeground="#06243a",
                            bordercolor="#454545")
        else:
            couleurs = dict(background="#ffffff", foreground="#1a1a1a",
                            headersbackground="#f0f0f0", headersforeground="#1a1a1a",
                            normalbackground="#ffffff", normalforeground="#1a1a1a",
                            weekendbackground="#fafafa", weekendforeground="#1a1a1a",
                            othermonthbackground="#f5f5f5", othermonthwebackground="#f5f5f5",
                            selectbackground="#0067c0", selectforeground="#ffffff",
                            bordercolor="#d9d9d9")
        val = "" if self._ph_on else self.var.get().strip()
        try:
            d = dt.date.fromisoformat(val)
        except ValueError:
            d = dt.date.today()

        racine = self.frame.winfo_toplevel()             # dialogue modal parent
        pop = tk.Toplevel(self.frame)
        pop.wm_overrideredirect(True)                    # sans bordure ni barre de titre
        self.entry.update_idletasks()
        x = self.entry.winfo_rootx()
        y = self.entry.winfo_rooty() + self.entry.winfo_height() + 2
        pop.wm_geometry(f"+{x}+{y}")
        # sv-ttk ne thématise pas le Calendar (widget non-ttk) : on lui passe couleurs + police.
        # Une police plus grande agrandit les cases des jours et le rend plus lisible.
        cal = Calendar(pop, selectmode="day", year=d.year, month=d.month, day=d.day,
                       date_pattern="yyyy-mm-dd", showweeknumbers=False, firstweekday="monday",
                       font="TkDefaultFont 11", **couleurs)
        cal.pack(padx=2, pady=2)
        # Remonter au-dessus du dialogue modal et prendre la main (sinon invisible/inerte).
        pop.update_idletasks()
        pop.lift()
        pop.attributes("-topmost", True)

        def fermer():
            try:
                pop.grab_release()
            except Exception:  # noqa: BLE001
                pass
            pop.destroy()
            try:
                racine.grab_set()                        # rend la modalité au dialogue
            except Exception:  # noqa: BLE001
                pass

        def choisir(_e=None):
            self.set(cal.get_date())
            fermer()

        def dehors(e):
            dans = (pop.winfo_rootx() <= e.x_root <= pop.winfo_rootx() + pop.winfo_width()
                    and pop.winfo_rooty() <= e.y_root <= pop.winfo_rooty() + pop.winfo_height())
            if not dans:
                fermer()

        cal.bind("<<CalendarSelected>>", choisir)
        pop.bind("<Escape>", lambda e: fermer())
        pop.bind("<Button-1>", dehors, add="+")
        pop.grab_set()
        cal.focus_set()

    def get(self) -> str:
        txt = "" if self._ph_on else self.var.get().strip()
        if txt:
            dt.date.fromisoformat(txt)  # lève ValueError si format invalide
        return txt

    def set(self, iso: str) -> None:
        self._ph_on = False
        self.entry.configure(foreground=self._fg)
        self.var.set(iso or "")
        if not (iso or "").strip():
            self._ph_show()

    def set_state(self, state: str) -> None:
        for w in (self.entry, self.bouton):
            if w is not None:
                try:
                    w.configure(state=state)
                except Exception:  # noqa: BLE001
                    pass


# --------------------------------------------------------------------------- #
# Boîte de dialogue : ajout / modification d'un locataire
# --------------------------------------------------------------------------- #

class DialogueLocataire(tk.Toplevel):
    def __init__(self, parent, modules: dict, adresses: list[str], valeurs: dict | None = None):
        super().__init__(parent)
        self.title("Locataire")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        self.resultat: dict | None = None
        self._modules = modules

        v = valeurs or {}
        # Loyer nu + charges saisis séparément dès qu'il y a des charges (modes « comprises »
        # ET « separees ») ; un seul champ loyer en mode « sans ». Aligné sur _flags_charges.
        mode = modules.get("mode_charges") or (
            "separees" if modules.get("loyer_nu_charges", True) else "sans")
        split = mode in ("comprises", "separees")
        caf, depot = modules["caf"], modules["depot_garantie"]
        self._vars: dict[str, tk.StringVar] = {}
        self._champs_num: set[str] = set()
        self._depot = depot
        dark = HAS_SVTTK and sv_ttk.get_theme() == "dark"
        muted = "#AAAAAA" if dark else "#666666"
        ligne = 0

        # Champs alignés : label en colonne 0, champ étiré en colonne 1 (sticky ew) pour
        # que tous les bords gauche/droit coïncident quelle que soit la longueur du label.
        self.columnconfigure(1, weight=1)

        def libelle(texte: str) -> None:
            nonlocal ligne
            ttk.Label(self, text=texte).grid(row=ligne, column=0, sticky="w", padx=(10, 6), pady=3)

        def entree(cle: str, *, num=False) -> None:
            nonlocal ligne
            var = tk.StringVar(value=str(v[cle]) if v.get(cle) not in (None, "") else "")
            ttk.Entry(self, textvariable=var, width=30).grid(
                row=ligne, column=1, sticky="ew", padx=(0, 10), pady=3)
            self._vars[cle] = var
            if num:
                self._champs_num.add(cle)
            ligne += 1

        def combo(cle: str, valeurs_combo: list[str], *, readonly=False, defaut="") -> None:
            nonlocal ligne
            var = tk.StringVar(value=str(v.get(cle, defaut)) or defaut)
            cb = ttk.Combobox(self, textvariable=var, values=valeurs_combo, width=30,
                              state="readonly" if readonly else "normal")
            cb.grid(row=ligne, column=1, sticky="ew", padx=(0, 10), pady=3)
            self._vars[cle] = var
            ligne += 1
            return cb

        def section(texte: str) -> None:
            nonlocal ligne
            ttk.Label(self, text=texte, font=("Segoe UI", 9, "bold"), foreground=muted).grid(
                row=ligne, column=0, columnspan=2, sticky="w", padx=10, pady=(12, 2))
            ligne += 1

        section("Identité")
        libelle("Nom *"); entree("nom")
        libelle("Prénom"); entree("prenom")
        section("Logement")
        libelle("Type de bien"); combo("type_bien", TYPES_BIEN, readonly=True,
                                       defaut=v.get("type_bien", TYPES_BIEN[0]))
        libelle("N° appart. / Nom du bien"); entree("identifiant")
        libelle("Adresse du logement"); combo("adresse", adresses)
        section("Loyer & charges")
        if split:
            libelle("Loyer nu (€)"); entree("loyer_nu", num=True)
            libelle("Charges (€)"); entree("charges", num=True)
        else:
            libelle("Loyer (€)"); entree("loyer", num=True)
        if caf:
            libelle("Part CAF / APL (€)"); entree("part_caf", num=True)
        if depot:
            libelle("Dépôt de garantie (€)"); entree("depot_garantie", num=True)

        section("Bail")
        libelle("Date d'entrée")
        self.date_entree = ChampDate(self, v.get("date_entree", ""))
        self.date_entree.grid(row=ligne, column=1, sticky="w", padx=(0, 10), pady=3)
        ligne += 1

        # Bloc « locataire parti » : active date de sortie + caution + observation.
        self.var_parti = tk.BooleanVar(value=bool(v.get("date_sortie")))
        ttk.Checkbutton(self, text="Le locataire est parti", variable=self.var_parti,
                        command=self._maj_sortie, style="Switch.TCheckbutton").grid(
                            row=ligne, column=0, columnspan=2, sticky="w", padx=8, pady=(8, 2))
        ligne += 1

        libelle("Date de sortie")
        self.date_sortie = ChampDate(self, v.get("date_sortie", ""))
        self.date_sortie.grid(row=ligne, column=1, sticky="w", padx=(0, 10), pady=3)
        ligne += 1

        self.var_caution = tk.BooleanVar(value=bool(v.get("caution_rendue")))
        self.chk_caution = ttk.Checkbutton(self, text="Caution rendue au locataire",
                                           variable=self.var_caution, style="Switch.TCheckbutton")
        if depot:
            self.chk_caution.grid(row=ligne, column=0, columnspan=2, sticky="w", padx=8, pady=2)
            ligne += 1

        libelle("Observation (motif de départ)")
        self.cb_observation = combo("observation", OBSERVATIONS, defaut=v.get("observation", ""))

        # Champs propres au bail, repris sur les documents de ce locataire.
        section("Documents (bail)")
        libelle("Date du bail")
        self.date_bail = ChampDate(self, v.get("date_bail", ""))
        self.date_bail.grid(row=ligne, column=1, sticky="w", padx=(0, 10), pady=3)
        ligne += 1
        libelle("Mode de paiement"); entree("mode_paiement")
        libelle("Jour d'échéance (ex. 5)"); entree("jour_echeance")

        ttk.Label(self, text="Adresse : choisissez-en une déjà saisie ou tapez-en une nouvelle.",
                  foreground=muted, wraplength=360, justify="left").grid(
            row=ligne, column=0, columnspan=2, padx=10, pady=(10, 4), sticky="w")
        ligne += 1

        barre = ttk.Frame(self)
        barre.grid(row=ligne, column=0, columnspan=2, sticky="e", padx=10, pady=(6, 12))
        ttk.Button(barre, text="Valider", command=self._valider,
                   style="Accent.TButton").pack(side="left", padx=6)
        ttk.Button(barre, text="Annuler", command=self.destroy).pack(side="left")

        self._maj_sortie()

    def _maj_sortie(self) -> None:
        actif = self.var_parti.get()
        etat = "normal" if actif else "disabled"
        self.date_sortie.set_state(etat)
        if self._depot:
            self.chk_caution.configure(state=etat)
        self.cb_observation.configure(state=("normal" if actif else "disabled"))

    def _valider(self) -> None:
        data: dict = {}
        try:
            nom = self._vars["nom"].get().strip()
            if not nom:
                raise ValueError("Le nom du locataire est obligatoire.")
            data["nom"] = nom
            for cle, var in self._vars.items():
                if cle in ("nom", "observation"):
                    continue
                val = var.get().strip()
                data[cle] = _parse_nombre(val) if cle in self._champs_num else val
            data["date_entree"] = self.date_entree.get()
            data["date_bail"] = self.date_bail.get()
            if self.var_parti.get():
                data["date_sortie"] = self.date_sortie.get()
                data["caution_rendue"] = bool(self.var_caution.get()) if self._depot else False
                data["observation"] = self._vars["observation"].get().strip()
            else:
                data["date_sortie"] = ""
                data["caution_rendue"] = False
                data["observation"] = ""
        except ValueError as e:
            messagebox.showerror("Saisie invalide", str(e), parent=self)
            return
        self.resultat = data
        self.destroy()


# --------------------------------------------------------------------------- #
# Fenêtre principale
# --------------------------------------------------------------------------- #

class Application(tk.Tk):
    COLS = [("nom", "Nom", 110), ("prenom", "Prénom", 90), ("type_bien", "Type", 90),
            ("identifiant", "N° / Nom", 120), ("adresse", "Adresse", 150),
            ("loyer", "Loyer", 65), ("date_entree", "Entrée", 90), ("date_sortie", "Sortie", 90)]

    def __init__(self):
        super().__init__()
        self.title(APP_TITRE)
        self.minsize(940, 720)  # widgets sv-ttk un peu plus hauts que clam
        self.locataires: list[dict] = []
        self.var_apparence = tk.StringVar(value=WINDOW_LABEL["systeme"])
        self._init_style()
        self._construire()

    def _init_style(self) -> None:
        self._style = ttk.Style(self)
        if not HAS_SVTTK:
            try:
                self._style.theme_use("clam")
            except tk.TclError:
                pass
        self._appliquer_theme_fenetre()

    def _appliquer_theme_fenetre(self, *_) -> None:
        """Applique le mode d'affichage (Système / Clair / Sombre) via sv-ttk si présent."""
        dark = False
        if HAS_SVTTK:
            choix = WINDOW_KEY.get(self.var_apparence.get(), "systeme")
            if choix == "systeme":
                sys_mode = darkdetect.theme() if HAS_DARKDETECT else None
                dark = (sys_mode or "Light").lower() == "dark"
            else:
                dark = choix == "sombre"
            sv_ttk.set_theme("dark" if dark else "light")
        self._config_styles(dark)

    def _config_styles(self, dark: bool = False) -> None:
        # set_theme réinitialise les styles : on (re)pose nos styles custom après chaque bascule.
        # Accent de marque dérivé du registre de thèmes (clair) ; bleu clair lisible en sombre.
        accent = "#7FB3FF" if dark else "#" + moteur.THEMES[moteur.THEME_DEFAUT]["primaire"]
        sous = "#AAAAAA" if dark else "#666666"
        self._style.configure("Titre.TLabel", font=("Segoe UI", 16, "bold"), foreground=accent)
        self._style.configure("SousTitre.TLabel", font=("Segoe UI", 9), foreground=sous)
        self._style.configure("Section.TLabelframe.Label", font=("Segoe UI", 10, "bold"))
        # sv-ttk ne thématise pas la LISTE déroulante des Combobox (popdown = tk.Listbox natif).
        # On force ses couleurs via l'option database (appliquée à toute l'app, dialogues compris).
        if dark:
            lb_bg, lb_fg, lb_sb, lb_sf = "#2b2b2b", "#f0f0f0", "#4cc2ff", "#06243a"
        else:
            lb_bg, lb_fg, lb_sb, lb_sf = "#ffffff", "#1a1a1a", "#0067c0", "#ffffff"
        self.option_add("*TCombobox*Listbox.background", lb_bg)
        self.option_add("*TCombobox*Listbox.foreground", lb_fg)
        self.option_add("*TCombobox*Listbox.selectBackground", lb_sb)
        self.option_add("*TCombobox*Listbox.selectForeground", lb_sf)

    def _construire(self) -> None:
        # ---- En-tête : titre + sous-titre à gauche, bascule d'affichage à droite ----
        header = ttk.Frame(self)
        header.pack(side="top", fill="x", padx=14, pady=(12, 4))
        titres = ttk.Frame(header)
        titres.pack(side="left", anchor="w")
        ttk.Label(titres, text="Générateur de suivi des loyers", style="Titre.TLabel").pack(anchor="w")
        ttk.Label(titres, style="SousTitre.TLabel",
                  text="Remplissez le formulaire puis cliquez « Générer le fichier Excel ». "
                       "Aucune connaissance technique requise.").pack(anchor="w")
        # Bascule d'affichage (segmented control via le style Toggle de sv-ttk ; radio classiques
        # en repli). value = libellé, car var_apparence stocke le libellé (cf. _appliquer_theme_fenetre).
        seg = ttk.Frame(header)
        seg.pack(side="right", anchor="e")
        ttk.Label(seg, text="Affichage :", style="SousTitre.TLabel").pack(side="left", padx=(0, 6))
        for _cle, label in WINDOW_THEMES:
            ttk.Radiobutton(seg, text=label, value=label, variable=self.var_apparence,
                            command=self._appliquer_theme_fenetre, style="Toggle.TButton",
                            takefocus=False).pack(side="left")

        # ---- Pied collant (posé avant le corps pour rester ancré en bas) ----
        af = ttk.Frame(self)
        af.pack(side="bottom", fill="x", padx=14, pady=(6, 12))
        ttk.Button(af, text="Charger une config…", command=self._charger).pack(side="left")
        ttk.Button(af, text="Enregistrer la config…", command=self._enregistrer).pack(side="left", padx=6)
        self.var_save_config = tk.BooleanVar(value=True)
        ttk.Checkbutton(af, text="Enregistrer aussi la configuration",
                        variable=self.var_save_config,
                        style="Switch.TCheckbutton").pack(side="left", padx=(16, 0))
        ttk.Button(af, text="Générer le fichier Excel", command=self._generer,
                   style="Accent.TButton").pack(side="right")

        # ---- Corps : bandeau de réglages (haut) + liste locataires (bas, extensible) ----
        body = ttk.Frame(self)
        body.pack(side="top", fill="both", expand=True, padx=14, pady=4)

        self.var_nom = tk.StringVar()
        self.var_prenom = tk.StringVar()
        self.var_sci = tk.BooleanVar(value=False)
        self.var_sci_nom = tk.StringVar()
        self.var_adresse = tk.StringVar()
        self.var_tel = tk.StringVar()
        self.var_email = tk.StringVar()
        self.var_iban = tk.StringVar()
        self.var_debut = tk.IntVar(value=ANNEE)
        self.var_fin = tk.IntVar(value=ANNEE + 2)
        self.var_caf = tk.BooleanVar(value=True)
        self.var_depot = tk.BooleanVar(value=True)
        self.var_documents = tk.BooleanVar(value=True)
        self.var_regul = tk.BooleanVar(value=True)
        self.var_irl = tk.BooleanVar(value=True)
        self.var_tableau = tk.BooleanVar(value=True)
        self.var_mode = tk.StringVar(value=MODE_LABEL["comprises"])
        self.var_theme = tk.StringVar(value=THEME_LABEL[moteur.THEME_DEFAUT])
        self.var_police = tk.StringVar(value=moteur.POLICE_DEFAUT)
        self.var_style_excel = tk.BooleanVar(value=True)

        def ligne(parent, r, label, widget):
            """Ligne label/champ alignée : labels en colonne 0, champ étiré en colonne 1."""
            ttk.Label(parent, text=label).grid(row=r, column=0, sticky="w", padx=(8, 6), pady=4)
            widget.grid(row=r, column=1, sticky="ew", padx=(0, 8), pady=4)
            parent.columnconfigure(1, weight=1)
            return widget

        # Bandeau de réglages : 3 colonnes (Bailleur · Contenu · Période+Apparence).
        reglages = ttk.Frame(body)
        reglages.pack(fill="x")
        for col, poids in ((0, 3), (1, 3), (2, 2)):
            reglages.columnconfigure(col, weight=poids, uniform="reglages")

        # -- Bailleur (colonne 0) --
        bf = ttk.LabelFrame(reglages, text="Bailleur", style="Section.TLabelframe")
        bf.grid(row=0, column=0, sticky="nsew", padx=(0, 7))
        ligne(bf, 0, "Nom *", ttk.Entry(bf, textvariable=self.var_nom))
        ligne(bf, 1, "Prénom", ttk.Entry(bf, textvariable=self.var_prenom))
        ttk.Checkbutton(bf, text="Le bailleur est une SCI", variable=self.var_sci,
                        command=self._maj_sci, style="Switch.TCheckbutton").grid(
                            row=2, column=0, columnspan=2, sticky="w", padx=8, pady=(6, 2))
        self.ent_sci = ligne(bf, 3, "Nom de la SCI", ttk.Entry(bf, textvariable=self.var_sci_nom))
        ligne(bf, 4, "Adresse", ttk.Entry(bf, textvariable=self.var_adresse))
        ligne(bf, 5, "Téléphone", ttk.Entry(bf, textvariable=self.var_tel))
        ligne(bf, 6, "E-mail", ttk.Entry(bf, textvariable=self.var_email))
        # IBAN du bailleur (compte de réception) repris sur les documents.
        ligne(bf, 7, "IBAN", ttk.Entry(bf, textvariable=self.var_iban))
        self._maj_sci()

        # -- Contenu du classeur (colonne 1) --
        mf = ttk.LabelFrame(reglages, text="Contenu du classeur", style="Section.TLabelframe")
        mf.grid(row=0, column=1, sticky="nsew", padx=7)
        ligne(mf, 0, "Loyer / charges",
              ttk.Combobox(mf, textvariable=self.var_mode, values=[lbl for _, lbl in MODES],
                           state="readonly"))
        options = [("Suivre la part CAF (tiers payant)", self.var_caf),
                   ("Suivre le dépôt de garantie", self.var_depot),
                   ("Documents (quittance, avis, relance)", self.var_documents),
                   ("Tableau de bord (graphiques)", self.var_tableau),
                   ("Régularisation annuelle des charges", self.var_regul),
                   ("Révision IRL (loyer revalorisé répercuté)", self.var_irl)]
        for i, (txt, var) in enumerate(options, start=1):
            ttk.Checkbutton(mf, text=txt, variable=var, style="Switch.TCheckbutton").grid(
                row=i, column=0, columnspan=2, sticky="w", padx=8, pady=3)

        # -- Période + Apparence du classeur (colonne 2, empilées) --
        col2 = ttk.Frame(reglages)
        col2.grid(row=0, column=2, sticky="nsew", padx=(7, 0))
        col2.columnconfigure(0, weight=1)
        pf = ttk.LabelFrame(col2, text="Période", style="Section.TLabelframe")
        pf.grid(row=0, column=0, sticky="ew")
        ligne(pf, 0, "Année début",
              ttk.Spinbox(pf, from_=2000, to=2100, textvariable=self.var_debut))
        ligne(pf, 1, "Année fin",
              ttk.Spinbox(pf, from_=2000, to=2100, textvariable=self.var_fin))
        apf = ttk.LabelFrame(col2, text="Apparence du classeur", style="Section.TLabelframe")
        apf.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        ligne(apf, 0, "Thème",
              ttk.Combobox(apf, textvariable=self.var_theme,
                           values=[THEME_LABEL[t] for t in moteur.THEMES], state="readonly"))
        ligne(apf, 1, "Police",
              ttk.Combobox(apf, textvariable=self.var_police, values=POLICES, state="readonly"))
        ttk.Checkbutton(apf, text="Générer pour Microsoft Excel",
                        variable=self.var_style_excel, style="Switch.TCheckbutton").grid(
                            row=2, column=0, columnspan=2, sticky="w", padx=8, pady=(6, 0))
        ttk.Label(apf, foreground="#888888", wraplength=300, justify="left",
                  text="Applique le style de graphique natif d'Excel au tableau de bord "
                       "(recommandé pour Excel). Décochez pour un rendu neutre (LibreOffice).").grid(
                            row=3, column=0, columnspan=2, sticky="w", padx=8, pady=(0, 4))

        # -- Locataires : pleine largeur, sous le bandeau, extensible en hauteur --
        lf = ttk.LabelFrame(body, text="Locataires", style="Section.TLabelframe")
        lf.pack(fill="both", expand=True, pady=(10, 0))
        barre = ttk.Frame(lf)
        barre.pack(fill="x", padx=8, pady=(8, 4))
        ttk.Button(barre, text="Ajouter", command=self._ajouter_locataire,
                   style="Accent.TButton").pack(side="left")
        ttk.Button(barre, text="Modifier", command=self._modifier_locataire).pack(side="left", padx=6)
        ttk.Button(barre, text="Supprimer", command=self._supprimer_locataire).pack(side="left")

        zone = ttk.Frame(lf)
        zone.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        cols = [c[0] for c in self.COLS]
        self.tree = ttk.Treeview(zone, columns=cols, show="headings", height=8)
        for cle, titre, larg in self.COLS:
            self.tree.heading(cle, text=titre)
            self.tree.column(cle, width=larg, anchor="e" if cle == "loyer" else "w")
        self.tree.pack(side="left", fill="both", expand=True)
        self.tree.bind("<Double-1>", lambda e: self._modifier_locataire())
        sb = ttk.Scrollbar(zone, orient="vertical", command=self.tree.yview)
        sb.pack(side="left", fill="y")
        self.tree.configure(yscrollcommand=sb.set)

    # ------------------------- gestion locataires ------------------------- #

    def _maj_sci(self) -> None:
        self.ent_sci.configure(state="normal" if self.var_sci.get() else "disabled")

    @staticmethod
    def _lire_annee(var: tk.IntVar, label: str) -> int:
        try:
            return int(var.get())
        except (tk.TclError, ValueError):
            raise ValueError(f"{label} invalide (saisissez une année).") from None

    def _modules(self) -> dict:
        return {"mode_charges": MODE_KEY.get(self.var_mode.get(), "comprises"),
                "caf": self.var_caf.get(), "depot_garantie": self.var_depot.get(),
                "documents": self.var_documents.get(), "tableau_bord": self.var_tableau.get(),
                "regularisation_charges": self.var_regul.get(), "irl": self.var_irl.get()}

    def _adresses(self) -> list[str]:
        vues, ordre = set(), []
        for loc in self.locataires:
            a = (loc.get("adresse") or "").strip()
            if a and a not in vues:
                vues.add(a)
                ordre.append(a)
        return ordre

    def _rafraichir_tree(self) -> None:
        self.tree.delete(*self.tree.get_children())
        for loc in self.locataires:
            loyer = loc.get("loyer_nu", loc.get("loyer", ""))
            self.tree.insert("", "end", values=(
                loc.get("nom", ""), loc.get("prenom", ""), loc.get("type_bien", ""),
                loc.get("identifiant", ""), loc.get("adresse", ""), loyer,
                loc.get("date_entree", ""), loc.get("date_sortie", "")))

    def _ajouter_locataire(self) -> None:
        d = DialogueLocataire(self, self._modules(), self._adresses())
        self.wait_window(d)
        if d.resultat:
            self.locataires.append(d.resultat)
            self._rafraichir_tree()

    def _selection(self) -> int | None:
        sel = self.tree.selection()
        return self.tree.index(sel[0]) if sel else None

    def _modifier_locataire(self) -> None:
        i = self._selection()
        if i is None:
            messagebox.showinfo(APP_TITRE, "Sélectionnez un locataire à modifier.")
            return
        d = DialogueLocataire(self, self._modules(), self._adresses(), self.locataires[i])
        self.wait_window(d)
        if d.resultat:
            self.locataires[i] = d.resultat
            self._rafraichir_tree()

    def _supprimer_locataire(self) -> None:
        i = self._selection()
        if i is None:
            return
        nom = self.locataires[i].get("nom", "")
        if messagebox.askyesno(APP_TITRE, f"Supprimer le locataire « {nom} » ?"):
            del self.locataires[i]
            self._rafraichir_tree()

    # ------------------------- config <-> formulaire ------------------------- #

    def _config(self) -> dict:
        return {
            "version": moteur.CONFIG_VERSION,
            "bailleur": {"nom": self.var_nom.get().strip(), "prenom": self.var_prenom.get().strip(),
                         "sci": self.var_sci.get(), "sci_nom": self.var_sci_nom.get().strip(),
                         "adresse": self.var_adresse.get(), "tel": self.var_tel.get(),
                         "email": self.var_email.get(),
                         "iban": self.var_iban.get().strip()},
            "periode": {"annee_debut": self._lire_annee(self.var_debut, "Année de début"),
                        "annee_fin": self._lire_annee(self.var_fin, "Année de fin")},
            "modules": self._modules(),
            "theme": THEME_KEY.get(self.var_theme.get(), moteur.THEME_DEFAUT),
            "police": self.var_police.get() or moteur.POLICE_DEFAUT,
            "style_excel": self.var_style_excel.get(),
            "locataires": self.locataires,
        }

    def _appliquer_config(self, cfg: dict) -> None:
        b = cfg.get("bailleur", {})
        self.var_nom.set(b.get("nom", ""))
        self.var_prenom.set(b.get("prenom", ""))
        self.var_sci.set(bool(b.get("sci", False)))
        self.var_sci_nom.set(b.get("sci_nom", ""))
        self._maj_sci()
        self.var_adresse.set(b.get("adresse", ""))
        self.var_tel.set(b.get("tel", ""))
        self.var_email.set(b.get("email", ""))
        self.var_iban.set(b.get("iban", ""))
        p = cfg.get("periode", {})
        self.var_debut.set(int(p.get("annee_debut", ANNEE)))
        self.var_fin.set(int(p.get("annee_fin", ANNEE)))
        m = cfg.get("modules", {})
        mode = m.get("mode_charges")
        if mode not in MODE_LABEL:
            mode = "separees" if m.get("loyer_nu_charges", True) else "sans"
        self.var_mode.set(MODE_LABEL[mode])
        self.var_caf.set(bool(m.get("caf", True)))
        self.var_depot.set(bool(m.get("depot_garantie", True)))
        self.var_documents.set(bool(m.get("documents", m.get("quittances", True))))
        self.var_tableau.set(bool(m.get("tableau_bord", True)))
        self.var_regul.set(bool(m.get("regularisation_charges", True)))
        self.var_irl.set(bool(m.get("irl", True)))
        theme = cfg.get("theme") if cfg.get("theme") in THEME_LABEL else moteur.THEME_DEFAUT
        self.var_theme.set(THEME_LABEL[theme])
        self.var_police.set(cfg.get("police") or moteur.POLICE_DEFAUT)
        self.var_style_excel.set(bool(cfg.get("style_excel", True)))
        self.locataires = list(cfg.get("locataires", []))
        self._rafraichir_tree()

    def _charger(self) -> None:
        chemin = filedialog.askopenfilename(
            title="Charger une configuration",
            filetypes=[("Configuration", "*.json *.yaml *.yml"), ("Tous", "*.*")])
        if not chemin:
            return
        try:
            p = Path(chemin)
            if p.suffix.lower() in (".yaml", ".yml"):
                import yaml
                brut = yaml.safe_load(p.read_text(encoding="utf-8"))
            else:
                brut = json.loads(p.read_text(encoding="utf-8"))
            cfg, avertis = moteur.migrer_config(brut)
            self._appliquer_config(cfg)
            if avertis:
                messagebox.showwarning(
                    APP_TITRE,
                    "Configuration chargée, avec quelques adaptations :\n\n• "
                    + "\n• ".join(avertis)
                    + "\n\nVérifiez les locataires et les options avant de générer.")
        except Exception as e:  # noqa: BLE001 - retour utilisateur
            messagebox.showerror(
                APP_TITRE,
                f"Impossible de charger ce fichier :\n{e}\n\n"
                "Le fichier est peut-être corrompu ou n'est pas une configuration.")

    def _enregistrer(self) -> None:
        try:
            cfg = moteur.valider_config(self._config())
        except ValueError as e:
            messagebox.showerror(APP_TITRE, str(e))
            return
        chemin = filedialog.asksaveasfilename(
            title="Enregistrer la configuration", defaultextension=".json",
            initialfile=f"config_{moteur.base_slug(cfg['bailleur'])}.json",
            filetypes=[("Configuration JSON", "*.json")])
        if not chemin:
            return
        self._ecrire_config(Path(chemin))
        messagebox.showinfo(APP_TITRE, "Configuration enregistrée.")

    def _ecrire_config(self, chemin: Path) -> None:
        chemin.write_text(
            json.dumps(self._config(), ensure_ascii=False, indent=2), encoding="utf-8")

    # ------------------------- génération ------------------------- #

    def _generer(self) -> None:
        try:
            cfg = moteur.valider_config(self._config())
        except ValueError as e:
            messagebox.showerror(APP_TITRE, str(e))
            return

        slug = moteur.base_slug(cfg["bailleur"])
        chemin = filedialog.asksaveasfilename(
            title="Enregistrer le classeur de suivi", defaultextension=".xlsx",
            initialfile=f"Suivi_{slug}.xlsx", filetypes=[("Classeur Excel", "*.xlsx")])
        if not chemin:
            return
        chemin = Path(chemin)

        # Par défaut, dossier dédié (xlsx + config). En régénération, on reste dans le dossier.
        if chemin.exists():
            xlsx = chemin
        else:
            dossier = chemin.parent / chemin.stem
            dossier.mkdir(parents=True, exist_ok=True)
            xlsx = dossier / chemin.name

        preserver = True
        if xlsx.exists():
            rep = messagebox.askyesnocancel(
                APP_TITRE,
                "Ce fichier existe déjà.\n\n"
                "Voulez-vous CONSERVER les loyers déjà saisis dedans ?\n\n"
                "• Oui : on met à jour la structure (locataires, options) en gardant vos saisies.\n"
                "• Non : on repart d'un fichier vierge (les saisies seront perdues).\n"
                "• Annuler : ne rien faire.")
            if rep is None:   # Annuler / fermeture de la fenêtre : on n'écrase rien.
                return
            preserver = rep
        orphelins: list[str] = []
        try:
            sortie = moteur.generer_workbook(cfg, xlsx, preserver=preserver,
                                             orphelins_out=orphelins)
        except Exception as e:  # noqa: BLE001 - retour utilisateur
            messagebox.showerror(APP_TITRE, f"Erreur pendant la génération :\n{e}")
            return

        msg = f"Fichier généré :\n{sortie}"
        if orphelins:
            msg += ("\n\n⚠ Saisies non réattribuées (locataire renommé ou supprimé) :\n"
                    + ", ".join(orphelins))
        if self.var_save_config.get():
            try:
                config_json = sortie.with_name(f"{sortie.stem}.config.json")
                self._ecrire_config(config_json)
                msg += f"\n\nConfiguration enregistrée à côté :\n{config_json.name}"
            except Exception as e:  # noqa: BLE001 - retour utilisateur
                msg += f"\n\n(La configuration n'a pas pu être enregistrée : {e})"

        if messagebox.askyesno(APP_TITRE, msg + "\n\nOuvrir le dossier ?"):
            self._ouvrir_dossier(sortie.parent)

    @staticmethod
    def _ouvrir_dossier(dossier: Path) -> None:
        import sys
        import subprocess
        try:
            if sys.platform.startswith("win"):
                import os
                os.startfile(dossier)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.run(["open", str(dossier)], check=False)
            else:
                subprocess.run(["xdg-open", str(dossier)], check=False)
        except Exception:  # noqa: BLE001
            pass


def main() -> None:
    Application().mainloop()


if __name__ == "__main__":
    main()
