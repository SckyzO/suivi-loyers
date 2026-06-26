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

`Guide` · `Locataires` (référentiel) · **une feuille par locataire** (saisie, nommée par
l'identifiant du bien) · `Données` (consolidée, **masquée**) · `Bilan` · documents (`Quittance`,
`Avis d'échéance`, `Lettre de relance`).

Flux de données clé :
- La saisie a lieu dans les **feuilles locataire** (`construire_feuilles_locataires`). Leurs
  colonnes calculées référencent la ligne du locataire dans `Locataires` (`_ref`).
- `Données` (`construire_donnees`, masquée) recopie chaque ligne par formule depuis les
  feuilles locataire et porte les **plages nommées `Suivi_*`**. Aucune double saisie.
- `Bilan` et les **documents** consomment uniquement ces plages `Suivi_*` via `SUMIFS`
  (sélection dynamique du locataire = `SUMIFS`, surtout pas d'`INDIRECT`).

## Règles à respecter

- **Formules en noms anglais** (`SUMIFS`, `VLOOKUP`, `IF`, `IFERROR`). Le format de fichier
  les stocke ainsi ; Excel FR les affiche localisées et LibreOffice les lit. Ne pas écrire
  `RECHERCHEV` dans une formule openpyxl.
- **Colonnes pilotées par les modules** (`loyer_nu_charges`, `caf`, `depot_garantie`). Une
  option désactivée doit retirer les colonnes correspondantes sans casser les formules ni les
  plages nommées. Tester systématiquement une config « minimale » (tout à `false`).
- **Période d'activité** : on ne génère des lignes que sur les mois compris entre `date_entree`
  et `date_sortie` (`_mois_actifs`). C'est ce qui gère les rotations de locataires. Ne pas
  revenir à une grille pleine.
- **Préservation des saisies** : toute régénération sur un fichier existant doit passer par
  `recolter_saisies` (qui lit **chaque feuille locataire**, pas un onglet `Suivi` unique) puis
  réinjection via `construire_feuilles_locataires(saisies=...)`. Clé `(nom, année, mois)`.
  Ne jamais écraser un fichier sans cette reprise quand `preserver=True`.
- **Nom de feuille locataire** : passe par `_nom_feuille` (≤ 31 car., caractères interdits,
  unicité), au format « identifiant - Nom ». Toute référence inter-feuilles passe par `_ref`
  (gère espaces et apostrophes).
- **Identité locataire** : `_identite(loc)` = « Nom Prénom » (clé unique partout : feuilles,
  Données, Bilan, documents, IRL, préservation). Le bail stocke `nom` + `prenom` séparés.
- **Couleurs conditionnelles** : toujours via `_fill_cf` (start+end color = bgColor). Un fill
  conditionnel avec `fgColor` seul N'APPARAÎT PAS dans Excel (bug corrigé).
- **Texte d'origine utilisateur** : l'écrire via `ecrire_texte` (anti-injection de formule :
  openpyxl traite une chaîne commençant par `=` comme une formule). Ne jamais écrire un champ
  saisi (nom, adresse, identifiant, observation, bailleur…) par un `ws.cell(...)` brut.
- **Totaux annuels** dans les feuilles locataire : ligne « Total <année> » + ligne vide entre
  années. `rows_map` ne contient que les lignes de mois, donc `Données`/IRL restent corrects.

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

## Compatibilité des configs (anciennes versions)

`migrer_config(raw) -> (cfg, avertissements)` est le **point unique** de rétro-compatibilité :
tolérant (ne lève pas pour un champ manquant), il convertit les anciens schémas (`bien` ->
`identifiant`, module `quittances` -> `documents`) et renvoie la liste des adaptations.
`valider_config` l'appelle puis impose les champs requis ; l'interface l'utilise au chargement
et affiche les avertissements. Toute évolution de schéma incrémente `CONFIG_VERSION` et ajoute
sa transition ici.

Fait : sauvegarde `.bak` du classeur avant écrasement (`generer_workbook`), coercition douce des
montants (virgule décimale) + validation stricte dans `valider_config` (montants/dates/unicité),
et avertissement explicite si la config vient d'une version **plus récente** que le binaire.

## Modules (tous implémentés)

`documents` (quittance + avis d'échéance + lettre de relance, builder générique
`construire_document`), `regularisation_charges` (`construire_regularisation`), et `irl`
(`construire_irl` : table d'indices IRL + calcul du loyer révisé par locataire). Les saisies
propres à ces onglets sont préservées par `recolter_regularisation` et `recolter_irl`, en plus
de `recolter_saisies` (mensuel).

Limite assumée IRL : c'est un **calculateur d'aide**. Le loyer attendu du suivi suit toujours
la fiche locataire courante ; l'intégration loyer-par-année (révision répercutée
automatiquement mois par mois) est reportée à la passe design.
