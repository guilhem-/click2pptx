# click2pptx

Ce projet fournit un petit utilitaire permettant de convertir un export
HTML de Freeplane contenant une image map en une présentation PPTX où
chaque zone cliquable est recréée sous la forme d'un rectangle
transparent pointant vers le lien associé.

## Installation

```bash
pip install .
```

Cette commande installe les dépendances `beautifulsoup4` et
`python-pptx` nécessaires au fonctionnement du script.

## Utilisation

```
click2pptx [-i SOURCE.html] [-o DESTINATION.pptx]
```

- `-i`, `--input` : fichier HTML source. S'il est omis, le script prend
  le premier fichier `*.html` présent dans le répertoire courant.
- `-o`, `--output` : chemin du fichier PPTX à produire. Si ce paramètre
  est omis, un dossier `output` est créé (s'il n'existe pas déjà) et le
  fichier `mind_map_clickable_YYYYMMDD_HHMMSS.pptx` y est écrit.

Exemple :

```bash
click2pptx -i mon_export.html -o presentation.pptx
```

Le programme lit le fichier HTML, extrait les zones cliquables ainsi que
l'image utilisée, puis génère un fichier PPTX équivalent. Chaque zone
active de l'image est recouverte par un rectangle invisible possédant le
hyperlien défini dans l'export Freeplane.
