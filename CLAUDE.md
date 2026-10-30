# CLAUDE.md : Suivi des loyers

Générateur de classeurs Excel de suivi des loyers (part locataire + part CAF), avec une
interface graphique pour un usage non technique sous Windows (et Linux depuis la 1.2.0).

## Public et contraintes

- **Utilisatrice finale** : non technicienne, sous Windows. Elle ne lance que l'`.exe`, jamais
  Python. Toute l'ergonomie passe par `interface_flet.py`. Elle ne fait **aucune sauvegarde**
  elle-même : l'application doit garantir qu'aucune saisie n'est perdue (cf. règle
  « Préservation des saisies »).
- **Mainteneur** : sous WSL, **tout en Docker** (pas de venv local). Le `.exe` se construit
  côté Windows avec `build-flet.bat` (PyInstaller n'est pas multi-plateforme). Le binaire
  **Linux** se construit dans Docker : `make linux` (`Dockerfile.linux`, base bookworm = glibc
  2.36 minimale). Runtime hook `packaging/pyi_rth_linux.py` : (1) fige `FLET_LINUX_DISTRO` à
  la distro du build (sinon flet-desktop cherche l'archive client sous le nom de la distro de
  l'**hôte** et tente un téléchargement GitHub) ; (2) rend aux sous-processus le
  `LD_LIBRARY_PATH` d'origine (sinon le client Flutter hérite des libs PyInstaller, Mesa de
  l'hôte échoue : « No provider of eglGetPlatformDisplayEXT », le client meurt). CI :
  `build.yml` (jobs `windows` + `linux` avec test de lancement sous Xvfb, puis `release`
  unique sur tag `v*`) ; `ci.yml` lance lint + tests sous Linux **et** sous Windows (job
  `windows` : smoke, tests de l'interface, et `tests/verrou_windows.py` qui reproduit le
  verrou d'un classeur ouvert dans Excel avec `CreateFileW`).
- **Langue** : interface, classeur généré et documentation en **français** (audience française).

## Architecture

Deux couches, à garder séparées :

- `generer_suivi_loyers.py` : **moteur** pur (openpyxl), sans dépendance à l'interface (Flet).
  Construit le classeur à partir d'une config validée. Réutilisable en CLI / Docker.
- `interface_flet.py` : **surcouche graphique** (Flet). Ne contient pas de logique de
  construction ; elle assemble une config et appelle `moteur.generer_workbook`.

Point d'entrée commun : `valider_config(dict) -> cfg`, utilisé par le YAML (`charger_config`)
et par l'interface. Garder cette frontière : pas de logique métier dans l'interface.

**Conventions Flet (1.0.5)** : dialogues via `page.show_dialog(dlg)` / `page.pop_dialog()` (les
anciens `page.open`/`close` n'existent plus) ; `FilePicker` est un **service** (`page.services`),
et ses `pick_files`/`save_file` sont des **coroutines** — les handlers qui les appellent sont
`async`. Les helpers de layout minuscules ont disparu : utiliser `ft.Padding/Margin/Border.*` et
`ft.Alignment.CENTER*` (majuscules), pas `ft.padding.*`. `page.launch_url()` n'existe plus
(1.0) : ouvrir une URL via le service `ft.UrlLauncher()` (coroutine, garder une référence sinon
le registre le purge). Libellé d'un bouton = `content` : l'ancien attribut `text` est ignoré
sans erreur. Bordure des champs (`TextField`/`Dropdown`) : `border=BORDURE_CHAMP` (dict
`ControlState` → `OutlineInputBorder`) ; ne pas revenir à `border_radius`/`border_color`/
`focused_border_color`, dépréciés en 1.0 et retirés en 1.3. Migration validée par captures
avant/après identiques au pixel (le rendu web a un défaut de police intermittent : refaire une
capture avant de conclure à un écart). Toute montée mineure de Flet casse l'API : la faire à
la main avec rendu (Dependabot n'accepte que les patchs). Un `TextStyle` de `text_theme` doit
porter une couleur explicite (`ft.Colors.ON_SURFACE`) sinon le texte reste blanc en thème clair.
Réglages persistés par `_charger_reglages`/`_enregistrer_reglages` (→ `%APPDATA%/SuiviLoyers/
settings.json`, écriture atomique sous verrou ; l'interface passe par `_sauver_reglages`, qui
signale un échec au lieu de l'avaler) ; la clé `verifier_maj` pilote la **vérification de mise à jour** au démarrage
(thread, API GitHub `releases/latest`, silencieuse hors-ligne, notif SnackBar + lien). La clé
`afficher_partis` (défaut `true`) mémorise le toggle « Afficher les partis » de la
table : **filtre d'affichage seulement**, le classeur garde tous les locataires. Le filtrage
vit dans des fonctions pures de module (`_filtrer_locataires`, `_partis_masques`,
`_libelle_compteur`, `_est_parti`), testées par `tests/ui_filtres.py`, comme le choix des
colonnes selon la largeur du tableau (`_colonnes_pour_largeur` : chaque colonne a une largeur
lisible `min` et un ordre de retrait `masquable` ; calibrés pour que tout tienne à la largeur
par défaut, 1180 px) et la disposition de la barre au-dessus (`_barre_sur_une_ligne` : toggle,
recherche et « Ajouter » sur la ligne du titre, ou dessous en fenêtre étroite ; le compteur
prend la place libre pour que les commandes ne bougent pas quand son texte change ;
`_disposer_barre` déplace les mêmes contrôles, ce que Flet 1.0 gère). La clé `classeurs`
(slug du bailleur → chemin) mémorise le **dernier classeur généré**, proposé d'office à la
génération ; si le fichier choisi ne porte pas les saisies attendues (neuf, vide, autre
classeur), l'interface demande d'où les reprendre (`_cas_reprise`) puis affiche un compte
rendu (`_lignes_rapport`) : fonctions pures testées par `tests/ui_reprise.py`. La clé
`aide_vue` dit si le mode d'emploi intégré (liste `AIDE`, dont les libellés cités doivent
rester ceux de l'interface) a été lu ; sinon il s'ouvre au démarrage.

**Saisie en cours** : recopiée dans `session.json` (à côté de `settings.json`, écriture
atomique) toutes les 3 s et après chaque `_rafraichir_table`, avec un drapeau `propre` =
« déjà dans un .json ». Tout code qui charge ou écrit la config dans un .json doit appeler
`_marquer_propre(cfg)`, sinon le prochain lancement proposerait à tort une reprise. Au
démarrage, seule une session non propre et non vide est proposée (sinon formulaire vierge,
comme avant) ; la sauvegarde automatique ne démarre qu'après la réponse. Une session écartée
(refusée, ou remplacée par un chargement confirmé) part dans `saisies-mises-de-cote/`, un
fichier daté par saisie (`_archiver_session`), jamais supprimé ni écrasé ; si ce déplacement
échoue, rien n'est remplacé et la sauvegarde automatique ne démarre pas (elle effacerait la
session). Les Réglages listent ces saisies (`_saisies_mises_de_cote`, section masquée s'il n'y
en a pas) avec « Reprendre », qui passe par la même confirmation que « Charger config »
(`_confirmer_remplacement`) ; le fichier repris reste sur le disque. « Charger config » lit
le fichier via `_lire_config` : accepte une saisie mise de côté, refuse un fichier sans
bailleur ni locataire (`migrer_config`, tolérant, rendait sinon un formulaire vide).
Supprimer un locataire passe par `_demander_suppression` (confirmation). Règles pures testées
par `tests/ui_protection.py`.

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
  Ne jamais écraser un fichier sans cette reprise quand `preserver=True`. Avant d'écraser,
  `generer_workbook` copie l'ancien classeur en `<nom>.bak-<horodatage>.xlsx` : ne jamais
  retirer cette copie. **Exigence produit : aucune perte de saisie, même « au cas où »** ;
  toute fonction qui supprime ou écarte des données (locataire supprimé ou renommé, période
  réduite) doit soit les reprendre, soit prévenir l'utilisateur et lui laisser une copie.
  `generer_workbook(..., source=, rapport=)` : `source` = classeur dont lire les saisies
  (générer sous un nouveau nom sans repartir de zéro ; jamais modifié ; absent → erreur,
  jamais un classeur vide). `rapport` (`RapportGeneration`) reçoit ce qui est repris, écarté
  (`Ecartee` avec motif et unité) et la copie `.bak`, calculés par **relecture** du nouveau
  classeur (`_bilan_reprise`) : toute nouvelle saisie préservée doit avoir son `recolter_*`
  pour être vérifiée. Copie `.bak` impossible → la génération s'arrête (ne jamais avaler cette
  erreur) ; classeur non remplaçable (ouvert dans Excel, lecture seule) → `_verifier_remplacable`
  échoue avant tout travail. Le pré-remplissage `=C<ligne>` de la régularisation n'est pas une
  saisie (`recolter_regularisation` l'ignore, sinon il pointerait une autre ligne).
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
make test-ui    # fonctions pures de l'interface (filtrage, reprise, protection) — image Dockerfile.linux
make sync       # code -> dossier Windows, xlsx -> ~/Downloads
make linux      # binaire Linux -> dist/linux/SuiviLoyers
```

Chaque changement visible par l'utilisateur ajoute une ligne sous `[Non publié]` dans
`CHANGELOG.md` (Keep a Changelog, en français) ; à la release, cette section prend le numéro
du tag. Les captures du README (`docs/screenshots/`) se refont avec le harnais `_flet_web.py`
(états `readme`, `readme_dialog`, `readme_dark`, `readme_reprise`, `readme_focus`, `reglages`)
+ Playwright ;
l'état `generation` rejoue une vraie génération sous un nouveau nom (sans dialogue natif),
l'état `session` un redémarrage sur une saisie non enregistrée, l'état `reglages_mdc` les
Réglages avec deux saisies mises de côté. Le classeur se rend via
LibreOffice (xlsx -> PDF -> PNG, cf. mémoire projet).

Le lint (`ruff.toml`) volontairement n'impose ni la longueur de ligne (E501) ni le tri d'imports
(isort) : wrapping et regroupement sont manuels.

Le smoke test (`tests/smoke.py`) valide systématiquement : modularité (pas de colonne CAF en
config minimale), période d'activité (rotation), et reprise des saisies après ajout d'un
locataire. Tout fichier à visualiser doit finir dans `~/Downloads` (cf. mémoire projet).

Un hook git `post-commit` relance `make sync-win` à chaque commit pour que `build-flet.bat` parte
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

- **Lisibilité des graphes du tableau de bord** (`graphe()` / `_police_graphique`) : couleurs de
  séries explicites `spPr` (palette Office `4472C4`/`ED7D31`/…), axe des montants `#,##0`
  (« 25 000 », pas « 25,000.00 € »), polices 10 pt axes/légende et 14 pt gras pour le titre,
  quadrillage clair. Honoré par Excel ET LibreOffice. NB : le **« Style 1 » natif d'Excel
  (galerie 2013+) n'est PAS reproductible par programme** — openpyxl/xlsxwriter ne gèrent que les
  styles 2007 (`<c:style>` 1-48), et l'injection des parties `chartStyle` modernes ne fonctionne
  pas sur les graphes openpyxl (format « ancien »). Ne pas réessayer ; pour le Style 1 exact,
  c'est un clic manuel dans Excel (pinceau). Feature `style_excel` retirée (cul-de-sac documenté).

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
