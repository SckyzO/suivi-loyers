# Journal des modifications

Les changements notables de chaque version. Le format suit
[Keep a Changelog](https://keepachangelog.com/fr/1.1.0/) et les numéros de version suivent
[Semantic Versioning](https://semver.org/lang/fr/).

## [Non publié]

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

[Non publié]: https://github.com/SckyzO/suivi-loyers/compare/v1.2.0...HEAD
[1.2.0]: https://github.com/SckyzO/suivi-loyers/compare/v1.1.0...v1.2.0
[1.1.0]: https://github.com/SckyzO/suivi-loyers/compare/v1.0.0...v1.1.0
[1.0.0]: https://github.com/SckyzO/suivi-loyers/releases/tag/v1.0.0
