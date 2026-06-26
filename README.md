# Suivi des loyers

Génère un classeur Excel pour suivre, mois par mois, les loyers encaissés auprès de chaque
locataire et de la CAF. On remplit un formulaire (bailleur, locataires, options), on clique
sur un bouton, et le fichier Excel est prêt. Chaque bailleur a son propre fichier.

Le classeur produit contient quatre onglets : un **Guide**, la fiche des **Locataires**, le
**Suivi** mensuel (avec alertes de couleur sur les impayés et les trop-perçus) et un **Bilan**
qui se calcule tout seul.

---

## Utiliser le logiciel (rien à installer)

Une fois que vous avez le fichier `SuiviLoyers.exe` :

1. Double-cliquez sur `SuiviLoyers.exe`. Une fenêtre s'ouvre.
2. Renseignez le **bailleur** (au minimum son nom) et la **période** (année de début et de fin).
3. Cochez les **options** utiles à ce bailleur :
   - *Séparer loyer nu / charges* : si vous suivez le loyer et les charges séparément.
   - *Suivre la part CAF* : si une partie du loyer est versée directement par la CAF.
   - *Suivre le dépôt de garantie*.
4. Ajoutez les **locataires** (bouton *Ajouter*). Pour chacun : nom, logement, montants, date
   d'entrée, et date de sortie s'il est parti.
5. Cliquez sur **Générer le fichier Excel** et choisissez où l'enregistrer.

Le logiciel crée un **dossier dédié** à ce bailleur. Dedans, vous trouvez le classeur Excel
**et** un fichier de configuration (`.config.json`). Cette configuration sert à retrouver vos
réglages la fois suivante, même si vous oubliez de l'exporter à la main. La case *Enregistrer
aussi la configuration* est cochée par défaut ; décochez-la si vous ne voulez que le classeur.

Ensuite, chaque mois, ouvrez le fichier Excel et saisissez les montants **reçus** dans les
cases jaunes (CAF reçue, part locataire reçue, dates). Tout le reste (totaux, écarts, statut,
bilan) se calcule automatiquement. Ne touchez pas aux cases bleutées.

### Les couleurs du suivi

| Statut | Couleur | Signification |
|---|---|---|
| Soldé | vert | Le total reçu couvre le total dû. |
| Trop-perçu | orange | Reçu supérieur au dû (avance, régularisation à prévoir). |
| Partiel | rouge | Reçu inférieur au dû (impayé partiel). |
| À encaisser | gris | Aucun paiement saisi pour ce mois. |

### Changer de locataire plus tard

Les locataires changent souvent : c'est prévu. Relancez `SuiviLoyers.exe`, cliquez sur
*Charger une config…* et ouvrez le fichier `.config.json` du dossier de ce bailleur. Ajoutez
ou retirez un locataire, puis régénérez **par-dessus le même fichier Excel**. Le logiciel vous
demande alors de **conserver les loyers déjà saisis** : répondez *Oui*, vos saisies sont
reprises et seule la structure est mise à jour.

Pour repartir sur un locataire qui arrive dans un logement déjà occupé avant : mettez une
**date de sortie** à l'ancien et une **date d'entrée** au nouveau. Chacun n'apparaît dans le
suivi que sur ses mois de présence.

---

## Fabriquer l'exécutable (une fois, sous Windows)

L'`.exe` se construit sous Windows.

1. Installez [Python](https://www.python.org/downloads/) (3.10 ou plus). À l'installation,
   cochez **« Add Python to PATH »**.
2. Double-cliquez sur `build.bat`.
3. À la fin, l'exécutable se trouve dans `dist\SuiviLoyers.exe`. Copiez-le où vous voulez ;
   il fonctionne seul, sans Python.

---

## Pour le mainteneur

Le projet se teste sans rien installer localement, via Docker.

```bash
# Construire l'image
docker compose build

# Générer un classeur depuis une config YAML (sortie dans ./sorties)
docker compose run --rm suivi configs/exemple.yaml
docker compose run --rm suivi configs/minimal.yaml
```

Le moteur (`generer_suivi_loyers.py`) lit une config et écrit le `.xlsx`. L'interface
(`interface.py`) est une surcouche graphique qui appelle ce même moteur. Les formules sont
écrites avec les noms de fonctions anglais (`SUMIFS`, `VLOOKUP`, `IF`) : Excel les affiche en
français et LibreOffice les lit sans souci.

### Structure du projet

| Fichier | Rôle |
|---|---|
| `generer_suivi_loyers.py` | Moteur : construit le classeur à partir d'une config. |
| `interface.py` | Interface graphique (Tkinter) pour l'utilisateur final. |
| `build.bat` | Construit `SuiviLoyers.exe` sous Windows (PyInstaller). |
| `configs/` | Exemples de configurations (`exemple.yaml`, `minimal.yaml`). |
| `Dockerfile`, `docker-compose.yml` | Environnement de test du moteur. |
| `requirements.txt` | Dépendances Python (openpyxl, PyYAML). |

### Modules disponibles

Trois options pilotent les colonnes et les calculs. Quand une option est désactivée, ses
colonnes disparaissent du classeur.

| Option (config) | Effet |
|---|---|
| `loyer_nu_charges` | Sépare le loyer nu et les charges ; sinon une seule colonne « Loyer dû ». |
| `caf` | Ajoute les colonnes part CAF attendue / reçue et la date. |
| `depot_garantie` | Ajoute le dépôt de garantie à la fiche locataire. |

Trois autres modules sont prévus mais pas encore actifs : quittances, révision IRL, et
régularisation des charges.

### Limites connues

- Le montant **attendu** suit la valeur courante de la fiche locataire. Un loyer révisé en
  cours d'année (IRL) demandera une table loyer-par-année, prévue avec le module IRL.
- La reprise des saisies associe les lignes par **nom de locataire + année + mois**. Renommer
  un locataire rompt ce lien ; ajouter ou retirer un locataire ne pose pas de problème.
