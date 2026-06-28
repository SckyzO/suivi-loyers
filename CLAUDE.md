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
l'identifiant du bien) · `Données` (consolidée, **masquée**) · `Bilan` · `Tableau de bord`
(graphiques) · `Régularisation charges` · `Révision IRL` · documents (`Quittance`,
`Avis d'échéance`, `Lettre de relance`, `Mise en demeure`).

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
- **Identité graphique** : toutes les couleurs viennent de l'objet `CHARTE` (jamais de hex en
  dur). `CHARTE` est résolu par `resoudre_charte(theme, police)` en tête de `generer_workbook`
  depuis le **registre `THEMES`** (id → couleurs + clé `table` = style de tableau Excel
  intégré pour l'onglet Locataires, dans la famille de teinte du thème ; fallback
  `TableStyleMedium2` si absente). Ajouter un thème = une entrée du registre, rien d'autre à
  toucher (l'interface lit `THEMES` pour sa liste). Défaut = `classique` (look
  historique). Thème inconnu → fallback `classique` + avertissement (via `migrer_config`). La
  **police** est appliquée en dernière passe par `appliquer_police(wb, police)` (réécrit le nom
  sur chaque cellule en préservant gras/taille/couleur ; modifier le style « Normal » ne se
  propage pas). Défaut = Tahoma (Aptos n'est pas une police d'origine Windows). Clés config
  `theme`/`police` tolérantes (`migrer_config`/`valider_config`).
- **Mode charges** : `_flags_charges(cfg)` → `(a_charges, charges_separees, mode)` avec
  `mode_charges` ∈ {`comprises`, `separees`, `sans`} (par bailleur). En `comprises`, loyer nu et
  charges restent calculés mais **masqués** (Données/régularisation en ont besoin). Ancien
  booléen `loyer_nu_charges` mappé en rétro-compat. Ne plus lire `loyer_nu_charges` directement.
- **Prorata** : `_prorata_suffixe(loc, année, mois)` renvoie `*jours/jours_du_mois` (jours réels,
  février géré) pour les mois d'entrée/sortie partiels ; appliqué au loyer/charges, pas à la CAF.
- **Texte d'origine utilisateur** : l'écrire via `ecrire_texte` (anti-injection de formule :
  openpyxl traite une chaîne commençant par `=` comme une formule). Ne jamais écrire un champ
  saisi (nom, adresse, identifiant, observation, bailleur…) par un `ws.cell(...)` brut.
- **Totaux annuels** dans les feuilles locataire : ligne « Total <année> » + ligne vide entre
  années. `rows_map` ne contient que les lignes de mois, donc `Données`/IRL restent corrects.

## Workflow de fin de modification

Lancer **`make`** après chaque modification. Cela enchaîne : build de l'image, **lint ruff**,
génération des classeurs d'exemple, `tests/smoke.py`, puis sync (code vers le dossier Windows,
`.xlsx` vers `~/Downloads`). Tout passe par Docker, rien d'installé en local.

```bash
make            # build + lint + gen + test + sync (workflow complet)
make lint       # lint ruff seul (bugs/style ; config dans ruff.toml)
make test       # smoke test seul (structure, modularité, préservation)
make sync       # code -> dossier Windows, xlsx -> ~/Downloads
```

Le lint (`ruff.toml`) volontairement n'impose ni la longueur de ligne (E501) ni le tri d'imports
(isort) : wrapping et regroupement sont manuels. `interface.py` tolère les one-liners `;` (E702).

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

`documents` (`construire_document` ; 4 `kind` : `quittance`, `avis`, `relance`,
`mise_en_demeure`), `tableau_bord` (`construire_tableau_bord` : cartes de
synthèse KPI + graphiques openpyxl légendés/commentés basés sur le Bilan),
`regularisation_charges` (`construire_regularisation` : filtre locataire + pré-remplissage =
provisions en mode comprises) et `irl` (`construire_irl`). Les saisies propres à ces onglets
sont préservées par `recolter_regularisation` et `recolter_irl`, en plus de `recolter_saisies`
(mensuel).

- **Documents conformes (art. 21 loi 1989)** : la **quittance** bascule en « REÇU DE PAIEMENT
  PARTIEL » (titre + corps par formule `IF`) quand le reçu ne solde pas le dû — comparaison
  sur l'**écart arrondi au centime** (`ROUND(du-recu,2)<=0`) car `Suivi_TotalDu` est calculé en
  direct avec prorata (fractions de centime) alors que `Suivi_TotalRecu` est une saisie arrondie.
  La **relance** n'annonce plus de « frais de relance » (non imputables) et la **mise en demeure**
  (`kind` dédié) cite le bail, la clause résolutoire et le commandement de payer par commissaire
  de justice. Relance + mise en demeure affichent le **reste dû cumulé** (`SUMIFS` sur le seul
  locataire, toutes périodes) ; pas de détail mois par mois (modèle à sélecteur unique). Champs
  optionnels repris sur les documents, **au bon niveau** : `iban` = **bailleur** (compte de
  réception, statique) ; `mode_paiement`, `jour_echeance`, `date_bail` = **locataire** (clauses du
  bail) — ajoutés comme colonnes du référentiel Locataires quand le module `documents` est actif,
  lus par `VLOOKUP` selon le locataire sélectionné (jamais en dur). Montants formatés via
  **`FIXED(x,2,TRUE)`** et non `TEXT(...,"0.00")` (la chaîne de format de `TEXT` est localisée :
  en Excel FR le `.` est mal interprété → 4,62 au lieu de 461,61). `date_bail` saisi en ISO est
  affiché en JJ/MM/AAAA dans le référentiel.

- **Bilan structuré** (`construire_bilan`) : titre + **évolution annuelle** (1 ligne/année,
  portefeuille) + **synthèse par locataire** (toutes années) + **un bloc détail par année**.
  Tous les blocs partagent le helper interne `bloc()` (en-tête + lignes `SUMIFS` + TOTAL surligné
  `calc` + mise en forme conditionnelle sur le Solde). Le filtrage par année passe par un critère
  `Suivi_Annee` (jamais une table figée). `construire_bilan` **renvoie les coordonnées** des blocs
  (`global`/`annuel` : hdr/first/last/total) que `construire_tableau_bord` consomme — pas de
  positions codées en dur entre les deux. Le tableau de bord lit le TOTAL global pour ses cartes
  KPI et ajoute un graphe « évolution annuelle » en plus des graphes par locataire.

- **IRL répercutée mois par mois** (`construire_irl`, Section 2 « Loyer applicable par année ») :
  modèle fermé `loyer(Y) = loyer_base × IRL_Tref(Y) / IRL_Tref(A0)`, où `A0` = 1ʳᵉ année active du
  locataire (loyer de base = fiche Locataires, colonne `loyer_nu`/`loyer_total`), `Tref` = trimestre
  dérivé de la date d'entrée (`_trimestre_de`). Indices absents → `IFERROR` retombe sur le loyer de
  base (aucune révision). La table est entièrement recalculée (aucune saisie de révision à préserver ;
  `recolter_irl` ne préserve que les indices). Elle est exposée en plages nommées `LoyerAn_Loc/Annee/
  Valeur` ; quand `modules.irl` est actif, le loyer attendu de chaque feuille locataire lit ces plages
  par `SUMIFS` (au lieu du renvoi direct à `Locataires`). **Seul le loyer est révisé**, pas les charges
  ni la CAF. Module IRL off (défaut) → renvoi direct inchangé (non-régression).
