# CLAUDE.md : Suivi des loyers

Générateur de classeurs Excel de suivi des loyers (part locataire + part CAF), avec une
interface graphique pour un usage non technique sous Windows.

## Public et contraintes

- **Utilisatrice finale** : non technicienne, sous Windows. Elle ne lance que l'`.exe`, jamais
  Python. Toute l'ergonomie passe par `interface.py`.
- **Mainteneur** : sous WSL, **tout en Docker** (pas de venv local). Le `.exe` se construit
  côté Windows avec `build.bat` (PyInstaller n'est pas multi-plateforme).
- **Langue** : interface, classeur généré et documentation en **français** (audience française).

## Architecture

Deux couches, à garder séparées :

- `generer_suivi_loyers.py` : **moteur** pur (openpyxl), sans dépendance à Tkinter.
  Construit le classeur à partir d'une config validée. Réutilisable en CLI / Docker.
- `interface.py` : **surcouche graphique** (Tkinter). Ne contient pas de logique de
  construction ; elle assemble une config et appelle `moteur.generer_workbook`.

Point d'entrée commun : `valider_config(dict) -> cfg`, utilisé par le YAML (`charger_config`)
et par l'interface. Garder cette frontière : pas de logique métier dans l'interface.

## Onglets du classeur

`Guide` · `Locataires` (référentiel + plages nommées) · `Suivi` (saisie mensuelle) · `Bilan`
(synthèse `SUMIFS`). Le `Suivi` et le `Bilan` lisent le référentiel via `RECHERCHEV` /
plages nommées. Ne pas dupliquer les montants en dur.

## Règles à respecter

- **Formules en noms anglais** (`SUMIFS`, `VLOOKUP`, `IF`, `IFERROR`). Le format de fichier
  les stocke ainsi ; Excel FR les affiche localisées et LibreOffice les lit. Ne pas écrire
  `RECHERCHEV` dans une formule openpyxl.
- **Colonnes pilotées par les modules** (`loyer_nu_charges`, `caf`, `depot_garantie`). Une
  option désactivée doit retirer les colonnes correspondantes sans casser les formules ni les
  plages nommées. Tester systématiquement une config « minimale » (tout à `false`).
- **Période d'activité** : `construire_suivi` ne génère des lignes que sur les mois compris
  entre `date_entree` et `date_sortie` (`_mois_actifs`). C'est ce qui gère les rotations de
  locataires. Ne pas revenir à une grille pleine.
- **Préservation des saisies** : toute régénération sur un fichier existant doit passer par
  `recolter_saisies` puis réinjection dans `construire_suivi(saisies=...)`. La clé est
  `(nom, année, mois)`. Ne jamais écraser un fichier sans cette reprise quand `preserver=True`.

## Workflow de fin de modification

Lancer **`make`** après chaque modification. Cela enchaîne : build de l'image, génération des
classeurs d'exemple, `tests/smoke.py`, puis sync (code vers le dossier Windows, `.xlsx` vers
`~/Downloads`). Tout passe par Docker, rien d'installé en local.

```bash
make            # build + gen + test + sync (workflow complet)
make test       # smoke test seul (structure, modularité, préservation)
make sync       # code -> dossier Windows, xlsx -> ~/Downloads
```

Le smoke test (`tests/smoke.py`) valide systématiquement : modularité (pas de colonne CAF en
config minimale), période d'activité (rotation), et reprise des saisies après ajout d'un
locataire. Tout fichier à visualiser doit finir dans `~/Downloads` (cf. mémoire projet).

Un hook git `post-commit` relance `make sync-win` à chaque commit pour que `build.bat` parte
toujours de la dernière version côté Windows.

## Reste à faire (phase 2)

Modules câblés dans la config mais inactifs (défaut `false`) : `quittances`, `irl`,
`regularisation_charges`. Le module IRL devra introduire une table loyer-par-année (le montant
attendu suit aujourd'hui la valeur courante de la fiche locataire).
