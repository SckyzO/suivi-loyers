# Journal des modifications

Les changements notables de chaque version. Le format suit
[Keep a Changelog](https://keepachangelog.com/fr/1.1.0/) et les numéros de version suivent
[Semantic Versioning](https://semver.org/lang/fr/).

## [Non publié]

### Ajouté

- Ce qui est saisi dans le logiciel (bailleur, locataires, options) est sauvegardé en continu.
  Si le logiciel est fermé avant d'avoir généré le classeur ou enregistré la configuration,
  il propose de reprendre cette saisie au lancement suivant.
- Supprimer un locataire demande une confirmation, qui rappelle que ses saisies ne seront
  plus reportées dans le classeur et propose plutôt d'indiquer son départ.
- Charger une configuration alors que la saisie en cours n'est pas enregistrée demande une
  confirmation. Une saisie écartée est mise de côté, jamais effacée.

### Modifié

- Dans une fenêtre étroite, la liste des locataires masque les colonnes les moins utiles
  (Adresse, puis CAF, Période et Bien) au lieu de toutes les tronquer : le nom, le loyer et
  l'état restent lisibles. À la taille de fenêtre par défaut, rien ne change.
- L'interrupteur des partis, la recherche et le bouton « Ajouter » sont sur la ligne du titre
  de la liste, qui gagne une ligne de hauteur. L'interrupteur s'appelle maintenant « Afficher
  les partis ». Dans une fenêtre étroite, ces commandes passent sous le titre.

## [1.2.1] - 2026-10-29

### Ajouté

- Le logiciel retient le dernier classeur généré pour chaque bailleur et le propose d'office
  au moment de générer : le remplacer reprend les saisies sans aucune question.
- Avant de générer dans un fichier qui ne contient pas les saisies attendues (nouveau nom,
  fichier vide, autre classeur), le logiciel demande d'où les reprendre, en indiquant combien
  de mois chaque fichier contient. Le classeur d'origine n'est jamais modifié.
- Compte rendu après chaque génération : saisies reprises, nom de la copie de sécurité, et
  toute saisie non reprise (locataire renommé ou supprimé, mois hors de la période, option
  désactivée) avec le fichier où elle reste.
- Mode d'emploi intégré (bouton « Aide »), ouvert tout seul au premier lancement.
- Onglet Guide du classeur : section « Mettre à jour le classeur ».

### Modifié

- Les erreurs de génération s'affichent dans une fenêtre au lieu d'un bandeau éphémère.

### Corrigé

- Générer sous un nouveau nom partait d'un classeur vide sans prévenir : les saisies
  restaient dans l'ancien fichier et semblaient perdues.
- Si la copie de sécurité ne peut pas être créée, la génération s'arrête au lieu de remplacer
  le classeur.
- Un classeur qui ne peut pas être remplacé (ouvert dans Excel, en lecture seule) donne un
  message clair avant tout travail, au lieu d'une erreur « Errno 13 ».
- Régularisation des charges en mode charges comprises : après l'ajout d'un locataire, les
  charges réelles pré-remplies d'un locataire pouvaient pointer les provisions d'un autre.
- Le bouton « Générer le classeur » affiche de nouveau « Génération… » pendant la génération.
- L'onglet Guide cite la mise en demeure parmi les documents à imprimer.

## [1.2.0] - 2026-10-29

### Ajouté

- Version Linux : chaque release contient `SuiviLoyers-linux-x86_64.tar.gz` (binaire et
  notice). Il fonctionne sur les distributions avec une glibc 2.36 ou plus récente (Debian 12,
  Ubuntu 24.04, Fedora 37 et suivantes). Pour le construire soi-même : `make linux`.
- Interrupteur « Afficher les locataires partis » dans la liste des locataires. Il masque les
  anciens locataires de la liste sans les retirer du classeur ; le choix est mémorisé.
- Le compteur de la liste indique combien de locataires partis sont masqués.

### Modifié

- Interface passée à Flet 1.0.5 et exécutables construits avec Python 3.14.
- La recherche et l'interrupteur des partis occupent une seconde ligne au-dessus de la liste,
  pour que le bouton « Ajouter » reste visible à la largeur par défaut de la fenêtre.
- Une même release publie désormais les versions Windows et Linux.

### Corrigé

- Le bouton « Supprimer » de la liste des locataires était coupé au bord droit, et le badge
  État l'était aussi sur une fenêtre étroite.
- La version Linux s'arrêtait au démarrage sur certaines distributions avec l'erreur
  « No provider of eglGetPlatformDisplayEXT found ».

### Retiré

- L'ancienne interface Tkinter (`interface.py`, `build.bat`), remplacée par l'interface Flet
  depuis la 1.1.0.

## [1.1.0] - 2026-07-03

### Ajouté

- Vérification d'une nouvelle version au démarrage, avec un lien pour la télécharger.
  Désactivable dans les Réglages.
- Proposition d'ouvrir le dossier du classeur une fois celui-ci généré.

### Modifié

- Interface migrée vers Flet 0.85.
- La case « Enregistrer la config à côté » est cochée par défaut.
- La période proposée par défaut va de l'année en cours à deux ans plus tard.

### Corrigé

- Texte blanc, donc illisible, en thème clair.
- Taille des fenêtres Réglages et Locataire.
- Adresse e-mail du bailleur tronquée.

## [1.0.0] - 2026-07-01

Première version publiée : interface graphique et classeur Excel par bailleur, avec une
feuille de saisie par locataire, bilan, tableau de bord, régularisation des charges, révision
IRL et documents à imprimer (quittance, avis d'échéance, lettre de relance, mise en demeure).

[Non publié]: https://github.com/SckyzO/suivi-loyers/compare/v1.2.1...HEAD
[1.2.1]: https://github.com/SckyzO/suivi-loyers/compare/v1.2.0...v1.2.1
[1.2.0]: https://github.com/SckyzO/suivi-loyers/compare/v1.1.0...v1.2.0
[1.1.0]: https://github.com/SckyzO/suivi-loyers/compare/v1.0.0...v1.1.0
[1.0.0]: https://github.com/SckyzO/suivi-loyers/releases/tag/v1.0.0
