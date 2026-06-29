# Générateur de Fiches PPI

Générateur de fiches PPI (Phrases Préfabriquées des Interactions) à partir de grilles d'analyse de PPI pré-remplies.

## 📋 Description

Cet outil permet de générer automatiquement des fiches d'analyse PPI à partir de paires de grilles Excel (oral/écrit) pré-remplies. Les fiches produites suivent la grille d'analyse développée dans le cadre du projet ANR PREFAB et résument certaines informations pour faciliter la saisie des fiches finales.

Les fonctionnalités principales :

- **Mode simple** : génération d'une fiche à partir d'une paire de fichiers (oral + écrit)
- **Mode batch** : génération de plusieurs fiches à partir d'un dossier contenant des paires de fichiers `*_Or.xlsx` / `*_Ph.xlsx`
- **Interface graphique** intuitive (Tkinter)
- **Formatage automatique** des cellules (gras, couleurs, tags XML)

## 🚀 Installation

### Depuis PyPI (recommandé)

```bash
pip install fiche-ppi
```

### Depuis les sources

```bash
git clone https://github.com/yourusername/fiche-ppi.git
cd fiche-ppi
pip install -e .
```

### Dépendances

- Python ≥ 3.8
- pandas ≥ 1.5.0
- numpy ≥ 1.24.0
- requests ≥ 2.28.0
- beautifulsoup4 ≥ 4.12.0
- ppi_analyser (dépendance interne)

**Note** : Sur Linux, `tkinter` peut nécessiter l'installation d'un paquet système :

```bash
sudo apt-get install python3-tk   # Debian/Ubuntu
sudo dnf install python3-tkinter  # Fedora
```

## 💻 Utilisation

### Interface graphique

Lancez l'interface graphique :

```bash
fiche-ppi-gui
```

### Mode graphique - Onglet Simple

1. Sélectionnez le fichier oral (`*_Or.xlsx`)
2. Sélectionnez le fichier écrit (`*_Ph.xlsx`)
3. Le chemin de sortie est automatiquement généré (modifiable)
4. Cliquez sur "⚡ Générer la fiche"

### Mode graphique - Onglet Batch

1. Sélectionnez un dossier contenant des paires de fichiers
2. Les paires `*_Or.xlsx` / `*_Ph.xlsx` sont automatiquement détectées
3. Cliquez sur "⚡ Générer toutes les fiches"

### Ligne de commande

```bash
fiche-ppi [-h] [-o OUTPUT] [--batch DOSSIER] [file_oral] [file_ecrit]

Génère une fiche PPI consolidée à partir de deux fichiers Excel.

positional arguments:
  file_oral             Fichier Excel oral (.xlsx)
  file_ecrit            Fichier Excel écrit (.xlsx)

options:
  -h, --help            show this help message and exit
  -o OUTPUT, --output OUTPUT
                        Fichier de sortie (mode simple ; défaut : <stem>_fiche.xlsx).
  --batch DOSSIER       Traite toutes les paires *_Or.xlsx / *_Ph.xlsx du dossier.

```

## 📁 Structure des fichiers d'entrée

Les fichiers Excel doivent contenir les colonnes suivantes :

- `Lemme`, `Forme PPI`, `Acception`, `Type de phrase`
- `Modalité d'énonciation`, `Expansion`, `Modifieurs`
- `Cooccurrents`, `Fonction globale`, `Fonctions spécifiques`
- `milieu`, `secteur`, `Remarques`,`node`

## Génération des colonnes de la fiche PPI

| Propriété | Comment c'est généré |
|-----------|---------------------|
| **Fe_1a PPI** | Prend la valeur de la colonne `Forme PPI` du premier enregistrement du DataFrame combiné (oral + écrit) |
| **Fe_1b Acception** | Joint les valeurs uniques non-vides de la colonne `Acception` (triées, séparées par ", ") |
| **Fe_1c Variantes formelles** | Calcule les variantes formelles pour l'oral et l'écrit séparément :<br>- Extrait les modifieurs de la colonne `Modifieurs`<br>- Nettoie les tokens de modifieurs<br>- Applique `remove_modifier()` pour retirer les modifieurs de la colonne `node`<br>- Formate : `\n- Oral : var1, var2\n- Écrit : var3, var4` |
| **Fe_1e Prononciation** | Pour le lemme (premier `Lemme` du DataFrame oral) :Scrape la page Wiktionary du lemme pour extraire les URLs des fichiers audio MP3 - Extrait les URLs des fichiers MP3<br>- Retourne une liste d'URLs séparées par des sauts de ligne |
| **Fe_2a Statut syntaxique phrase** | Joint les valeurs uniques de la colonne `Type de phrase` |
| **Fe_2c Modalité de phrase** | Joint les valeurs uniques de la colonne `Modalité d'énonciation` |
| **Fe_2e Expansion éventuelle** | Joint les valeurs uniques de la colonne `Expansion` |
| **Fe_3a Fonction globale** | Joint les valeurs uniques de la colonne `Fonction globale` |
| **Fe_3b Fonctions spécifiques** | Joint les valeurs uniques de la colonne `Fonctions spécifiques` |
| **Fe_3c Codes Fonction globale** | Joint les valeurs uniques de la colonne `Fonction globale` (identique à Fe_3a) |
| **Fe_3d Codes Fonctions spécifiques** | Joint les valeurs uniques de la colonne `Fonctions spécifiques` (identique à Fe_3b) |
| **Fe_3f Structure interactionnelle** | Pour l'oral, l'écrit et le combiné :<br>- Calcule les fréquences des colonnes `Déclenchement`, `Portée`, `Position`<br>- Formate : `Oral :\n\t- Déclenchement : val1 (n), val2 (n)\n\t- Portée : ...` |
| **Fe_3g Contexte spécifique** | Joint les valeurs uniques de la colonne `milieu` |
| **Fe_3h Modalité écrite et orale** | Joint les valeurs uniques de la colonne `secteur` |
| **Fe_4a Cooccurrents privilégiés** | Pour l'oral, l'écrit et le combiné :<br>- Parse la colonne `Cooccurrents`<br>- Sépare les éléments antéposés (a) et postposés (p)<br>- Compte les fréquences avec `Counter()`<br>- Calcule les pourcentages d'antéposés/postposés<br>- Formate : `Oral :\n\t- Cooccurrents antéposés (X%) : ...\n\t- Cooccurrents postposés (Y%) : ...` |
| **Fe_4b Modifieurs de la PPI** | Joint les valeurs uniques de la colonne `Modifieurs` |
| **Fe_9a Remarques** | Joint les valeurs uniques de la colonne `Remarques` |

## Normalisation des données d'entrée

L'outil procède à une normalisation de certaines valeurs afin d'obtenir des calculs corrects, ex. ante:mais--> mais (a); post:mais --> mais (p)
Un minimum de normalisation est toutefois attendu.

## Contrôle d'intégrité

L'outil affiche des avertissements lorsque des lignes vides sont détectées:
ex. 
	[warn] [Oral] Portée (33) ≠ Position (34) — ligne(s) incomplète(s)
	[warn] [Écrit] Portée (44) ≠ Position (43) — ligne(s) incomplète(s)

---

**LIDILEM · ANR PREFAB**
