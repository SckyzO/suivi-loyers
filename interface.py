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
    """Saisie de date : zone texte AAAA-MM-JJ (toujours tapable au clavier) + bouton 📅."""

    def __init__(self, parent, iso: str = ""):
        self.frame = ttk.Frame(parent)
        self.var = tk.StringVar(value=iso or "")
        self.entry = ttk.Entry(self.frame, textvariable=self.var, width=12)
        self.entry.pack(side="left")
        ttk.Label(self.frame, text="AAAA-MM-JJ", foreground="#888").pack(side="left", padx=(4, 0))
        self.bouton = None
        if HAS_CAL:
            self.bouton = ttk.Button(self.frame, text="📅", width=3, command=self._ouvrir)
            self.bouton.pack(side="left", padx=(4, 0))

    def grid(self, **kw):
        self.frame.grid(**kw)

    def _ouvrir(self) -> None:
        top = tk.Toplevel(self.frame)
        top.title("Choisir une date")
        top.transient(self.frame.winfo_toplevel())
        top.grab_set()
        try:
            d = dt.date.fromisoformat(self.var.get().strip())
        except ValueError:
            d = dt.date.today()
        cal = Calendar(top, selectmode="day", year=d.year, month=d.month, day=d.day,
                       date_pattern="yyyy-mm-dd")
        cal.pack(padx=8, pady=8)

        def valider():
            self.var.set(cal.get_date())
            top.destroy()

        ttk.Button(top, text="Valider", command=valider).pack(pady=(0, 8))
        self.entry.focus_set()

    def get(self) -> str:
        txt = self.var.get().strip()
        if txt:
            dt.date.fromisoformat(txt)  # lève ValueError si format invalide
        return txt

    def set(self, iso: str) -> None:
        self.var.set(iso or "")

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
        ligne = 0

        def libelle(texte: str) -> None:
            nonlocal ligne
            ttk.Label(self, text=texte).grid(row=ligne, column=0, sticky="w", padx=8, pady=3)

        def entree(cle: str, *, num=False) -> None:
            nonlocal ligne
            var = tk.StringVar(value=str(v[cle]) if v.get(cle) not in (None, "") else "")
            ttk.Entry(self, textvariable=var, width=34).grid(row=ligne, column=1, padx=8, pady=3)
            self._vars[cle] = var
            if num:
                self._champs_num.add(cle)
            ligne += 1

        def combo(cle: str, valeurs_combo: list[str], *, readonly=False, defaut="") -> None:
            nonlocal ligne
            var = tk.StringVar(value=str(v.get(cle, defaut)) or defaut)
            cb = ttk.Combobox(self, textvariable=var, values=valeurs_combo, width=32,
                              state="readonly" if readonly else "normal")
            cb.grid(row=ligne, column=1, padx=8, pady=3)
            self._vars[cle] = var
            ligne += 1
            return cb

        libelle("Nom *"); entree("nom")
        libelle("Prénom"); entree("prenom")
        libelle("Type de bien"); combo("type_bien", TYPES_BIEN, readonly=True,
                                       defaut=v.get("type_bien", TYPES_BIEN[0]))
        libelle("N° d'appartement / Nom de la maison"); entree("identifiant")
        libelle("Adresse du logement"); combo("adresse", adresses)
        if split:
            libelle("Loyer nu (€)"); entree("loyer_nu", num=True)
            libelle("Charges (€)"); entree("charges", num=True)
        else:
            libelle("Loyer (€)"); entree("loyer", num=True)
        if caf:
            libelle("Part CAF / APL (€)"); entree("part_caf", num=True)
        if depot:
            libelle("Dépôt de garantie (€)"); entree("depot_garantie", num=True)

        libelle("Date d'entrée")
        self.date_entree = ChampDate(self, v.get("date_entree", ""))
        self.date_entree.grid(row=ligne, column=1, sticky="w", padx=8, pady=3)
        ligne += 1

        # Bloc « locataire parti » : active date de sortie + caution + observation.
        self.var_parti = tk.BooleanVar(value=bool(v.get("date_sortie")))
        ttk.Checkbutton(self, text="Le locataire est parti", variable=self.var_parti,
                        command=self._maj_sortie).grid(row=ligne, column=0, columnspan=2,
                                                       sticky="w", padx=8, pady=(8, 2))
        ligne += 1

        libelle("Date de sortie")
        self.date_sortie = ChampDate(self, v.get("date_sortie", ""))
        self.date_sortie.grid(row=ligne, column=1, sticky="w", padx=8, pady=3)
        ligne += 1

        self.var_caution = tk.BooleanVar(value=bool(v.get("caution_rendue")))
        self.chk_caution = ttk.Checkbutton(self, text="Caution rendue au locataire",
                                           variable=self.var_caution)
        if depot:
            self.chk_caution.grid(row=ligne, column=0, columnspan=2, sticky="w", padx=8, pady=2)
            ligne += 1

        libelle("Observation (motif de départ)")
        self.cb_observation = combo("observation", OBSERVATIONS, defaut=v.get("observation", ""))

        ttk.Label(self, text="Adresse : choisissez-en une déjà saisie ou tapez-en une nouvelle.",
                  foreground="#666", wraplength=320, justify="left").grid(
            row=ligne, column=0, columnspan=2, padx=8, pady=(4, 6), sticky="w")
        ligne += 1

        barre = ttk.Frame(self)
        barre.grid(row=ligne, column=0, columnspan=2, pady=8)
        ttk.Button(barre, text="Valider", command=self._valider).pack(side="left", padx=6)
        ttk.Button(barre, text="Annuler", command=self.destroy).pack(side="left", padx=6)

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
        self._init_style()
        self._construire()

    def _init_style(self) -> None:
        style = ttk.Style(self)
        if HAS_SVTTK:
            sv_ttk.set_theme("light")          # look Windows 11 (mode clair)
        else:
            try:
                style.theme_use("clam")
            except tk.TclError:
                pass
        # Accent de marque dérivé du registre de thèmes (pas de hex en dur).
        accent = "#" + moteur.THEMES[moteur.THEME_DEFAUT]["primaire"]
        style.configure("Titre.TLabel", font=("Segoe UI", 16, "bold"), foreground=accent)
        style.configure("SousTitre.TLabel", font=("Segoe UI", 9), foreground="#666")
        style.configure("Section.TLabelframe.Label", font=("Segoe UI", 10, "bold"))

    def _construire(self) -> None:
        ttk.Label(self, text="Générateur de suivi des loyers", style="Titre.TLabel").pack(
            anchor="w", padx=14, pady=(12, 0))
        ttk.Label(self, style="SousTitre.TLabel",
                  text="Remplissez le formulaire puis cliquez « Générer le fichier Excel ». "
                       "Aucune connaissance technique requise.").pack(anchor="w", padx=14, pady=(0, 6))

        bf = ttk.LabelFrame(self, text="Bailleur", style="Section.TLabelframe")
        bf.pack(fill="x", padx=14, pady=6)
        self.var_nom = tk.StringVar()
        self.var_prenom = tk.StringVar()
        self.var_sci = tk.BooleanVar(value=False)
        self.var_sci_nom = tk.StringVar()
        self.var_adresse = tk.StringVar()
        self.var_tel = tk.StringVar()
        self.var_email = tk.StringVar()

        def champ(r, c, lab, var):
            ttk.Label(bf, text=lab).grid(row=r, column=c * 2, sticky="w", padx=8, pady=4)
            ttk.Entry(bf, textvariable=var, width=30).grid(row=r, column=c * 2 + 1, padx=8, pady=4)

        champ(0, 0, "Nom *", self.var_nom)
        champ(0, 1, "Prénom", self.var_prenom)
        ttk.Checkbutton(bf, text="SCI", variable=self.var_sci, command=self._maj_sci).grid(
            row=1, column=0, sticky="w", padx=8, pady=4)
        self.ent_sci = ttk.Entry(bf, textvariable=self.var_sci_nom, width=30)
        self.ent_sci.grid(row=1, column=1, padx=8, pady=4)
        ttk.Label(bf, text="nom de la SCI", foreground="#666").grid(
            row=1, column=2, sticky="w", padx=4)
        champ(2, 0, "Adresse", self.var_adresse)
        champ(2, 1, "Téléphone", self.var_tel)
        champ(3, 0, "E-mail", self.var_email)
        self._maj_sci()

        pm = ttk.Frame(self)
        pm.pack(fill="x", padx=14, pady=6)

        pf = ttk.LabelFrame(pm, text="Période", style="Section.TLabelframe")
        pf.pack(side="left", fill="y")
        self.var_debut = tk.IntVar(value=ANNEE)
        self.var_fin = tk.IntVar(value=ANNEE + 2)
        ttk.Label(pf, text="Année début").grid(row=0, column=0, padx=8, pady=4, sticky="w")
        ttk.Spinbox(pf, from_=2000, to=2100, textvariable=self.var_debut, width=8).grid(
            row=0, column=1, padx=8, pady=4)
        ttk.Label(pf, text="Année fin").grid(row=1, column=0, padx=8, pady=4, sticky="w")
        ttk.Spinbox(pf, from_=2000, to=2100, textvariable=self.var_fin, width=8).grid(
            row=1, column=1, padx=8, pady=4)

        mf = ttk.LabelFrame(pm, text="Options à inclure", style="Section.TLabelframe")
        mf.pack(side="left", fill="both", expand=True, padx=(10, 0))
        self.var_caf = tk.BooleanVar(value=True)
        self.var_depot = tk.BooleanVar(value=True)
        self.var_documents = tk.BooleanVar(value=True)
        self.var_regul = tk.BooleanVar(value=True)
        self.var_irl = tk.BooleanVar(value=True)
        self.var_tableau = tk.BooleanVar(value=True)
        self.var_mode = tk.StringVar(value=MODE_LABEL["comprises"])

        frm = ttk.Frame(mf)
        frm.pack(anchor="w", padx=8, pady=2, fill="x")
        ttk.Label(frm, text="Loyer / charges :").pack(side="left")
        ttk.Combobox(frm, textvariable=self.var_mode, values=[lbl for _, lbl in MODES],
                     state="readonly", width=26).pack(side="left", padx=4)
        ttk.Checkbutton(mf, text="Suivre la part CAF (tiers payant)",
                        variable=self.var_caf).pack(anchor="w", padx=8, pady=1)
        ttk.Checkbutton(mf, text="Suivre le dépôt de garantie",
                        variable=self.var_depot).pack(anchor="w", padx=8, pady=1)
        ttk.Checkbutton(mf, text="Documents à imprimer (quittance, avis, relance)",
                        variable=self.var_documents).pack(anchor="w", padx=8, pady=1)
        ttk.Checkbutton(mf, text="Tableau de bord (graphiques)",
                        variable=self.var_tableau).pack(anchor="w", padx=8, pady=1)
        ttk.Checkbutton(mf, text="Régularisation annuelle des charges",
                        variable=self.var_regul).pack(anchor="w", padx=8, pady=1)
        ttk.Checkbutton(mf, text="Révision IRL (loyer revalorisé répercuté dans le suivi)",
                        variable=self.var_irl).pack(anchor="w", padx=8, pady=1)

        apf = ttk.LabelFrame(self, text="Apparence", style="Section.TLabelframe")
        apf.pack(fill="x", padx=14, pady=6)
        self.var_theme = tk.StringVar(value=THEME_LABEL[moteur.THEME_DEFAUT])
        self.var_police = tk.StringVar(value=moteur.POLICE_DEFAUT)
        ttk.Label(apf, text="Thème (couleurs)").grid(row=0, column=0, padx=8, pady=6, sticky="w")
        ttk.Combobox(apf, textvariable=self.var_theme,
                     values=[THEME_LABEL[t] for t in moteur.THEMES],
                     state="readonly", width=22).grid(row=0, column=1, padx=8, pady=6, sticky="w")
        ttk.Label(apf, text="Police").grid(row=0, column=2, padx=(24, 8), pady=6, sticky="w")
        ttk.Combobox(apf, textvariable=self.var_police, values=POLICES,
                     state="readonly", width=18).grid(row=0, column=3, padx=8, pady=6, sticky="w")

        lf = ttk.LabelFrame(self, text="Locataires", style="Section.TLabelframe")
        lf.pack(fill="both", expand=True, padx=14, pady=6)
        cols = [c[0] for c in self.COLS]
        self.tree = ttk.Treeview(lf, columns=cols, show="headings", height=8)
        for cle, titre, larg in self.COLS:
            self.tree.heading(cle, text=titre)
            self.tree.column(cle, width=larg, anchor="w")
        self.tree.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=8)
        self.tree.bind("<Double-1>", lambda e: self._modifier_locataire())
        sb = ttk.Scrollbar(lf, orient="vertical", command=self.tree.yview)
        sb.pack(side="left", fill="y", pady=8)
        self.tree.configure(yscrollcommand=sb.set)

        bl = ttk.Frame(lf)
        bl.pack(side="left", fill="y", padx=8, pady=8)
        ttk.Button(bl, text="Ajouter", command=self._ajouter_locataire).pack(fill="x", pady=3)
        ttk.Button(bl, text="Modifier", command=self._modifier_locataire).pack(fill="x", pady=3)
        ttk.Button(bl, text="Supprimer", command=self._supprimer_locataire).pack(fill="x", pady=3)

        af = ttk.Frame(self)
        af.pack(fill="x", padx=14, pady=(4, 14))
        ttk.Button(af, text="Charger une config…", command=self._charger).pack(side="left")
        ttk.Button(af, text="Enregistrer la config…", command=self._enregistrer).pack(
            side="left", padx=6)
        self.var_save_config = tk.BooleanVar(value=True)
        ttk.Checkbutton(af, text="Enregistrer aussi la configuration",
                        variable=self.var_save_config).pack(side="left", padx=(16, 0))
        # Action principale mise en avant (style accentué si sv-ttk est présent).
        ttk.Button(af, text="Générer le fichier Excel", command=self._generer,
                   style="Accent.TButton").pack(side="right")

    # ------------------------- gestion locataires ------------------------- #

    def _maj_sci(self) -> None:
        self.ent_sci.configure(state="normal" if self.var_sci.get() else "disabled")

    @staticmethod
    def _lire_annee(var: tk.IntVar, label: str) -> int:
        try:
            return int(var.get())
        except (tk.TclError, ValueError):
            raise ValueError(f"{label} invalide (saisissez une année).")

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
                         "email": self.var_email.get()},
            "periode": {"annee_debut": self._lire_annee(self.var_debut, "Année de début"),
                        "annee_fin": self._lire_annee(self.var_fin, "Année de fin")},
            "modules": self._modules(),
            "theme": THEME_KEY.get(self.var_theme.get(), moteur.THEME_DEFAUT),
            "police": self.var_police.get() or moteur.POLICE_DEFAUT,
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
