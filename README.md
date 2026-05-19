# RCS Madagascar — Extracteur d'entreprises

Outil pour extraire la liste des entreprises immatriculées au Registre du Commerce et des Sociétés de Madagascar (https://www.rcsmada.mg) et générer un fichier Excel formaté.

## Pour les utilisateurs (non-technique)

Aucune installation requise. Aucun terminal, aucune commande, aucun Python.

### macOS

1. Télécharger **RCS Madagascar.app**
2. Double-cliquer dessus

> Si macOS affiche « impossible d'ouvrir car le développeur n'est pas vérifié » : clic droit sur l'app → « Ouvrir » → confirmer. C'est à faire une seule fois.

### Windows

1. Télécharger **RCS Madagascar.exe**
2. Double-cliquer dessus

> Si Windows Defender affiche un avertissement : cliquer sur « Informations complémentaires » → « Exécuter quand même ». C'est à faire une seule fois.

### Utilisation

L'interface graphique s'ouvre directement avec les champs à remplir :

- **Ville (greffe)** : ville d'immatriculation
- **Année** : année d'immatriculation
- **Forme juridique** : SARL, SA-AG, EI, GIE, etc.
- **Type d'assujetti** : personne physique, morale, GIE, etc.
- **Dossier de sortie** : où enregistrer le fichier Excel

Cliquer sur **« Lancer l'extraction »**. Une barre de progression affiche l'avancement et l'estimation du temps restant. Une boîte de dialogue confirme la fin.

---

## Pour le mainteneur (génération des exécutables)

Cette section explique comment produire les fichiers `.app` (Mac) et `.exe` (Windows) à distribuer aux utilisateurs.

### Prérequis (une fois)

Installer un Python 3.10+ avec Tkinter fonctionnel :

**macOS**
- Télécharger depuis https://www.python.org/downloads/ (recommandé)
- OU `brew install python-tk@3.12`
- ⚠️ **Ne PAS utiliser** le Python d'Apple (`/usr/bin/python3` ou Command Line Tools) : son Tkinter est incompatible

**Windows**
- Télécharger depuis https://www.python.org/downloads/
- À l'installation, cocher **« Add python.exe to PATH »** et **« tcl/tk and IDLE »**

### Générer le .app (sur Mac)

Dans un terminal, depuis le dossier du projet :

```bash
./build_mac.sh
```

Le script :
1. Détecte automatiquement un Python compatible (ignore celui d'Apple)
2. Vérifie que Tkinter marche
3. Installe les dépendances et PyInstaller
4. Compile l'app
5. Ouvre le dossier `dist/` contenant `RCS Madagascar.app`

L'app produite (~30-50 Mo) est totalement autonome : elle contient Python, Tkinter et toutes les libs. Elle peut être copiée sur n'importe quel Mac et lancée sans rien installer.

### Générer le .exe (sur Windows)

Sur une machine Windows, depuis le dossier du projet :

```cmd
build_windows.bat
```

Ou simplement double-cliquer sur `build_windows.bat` dans l'Explorateur.

Le script fait la même chose et produit `dist\RCS Madagascar.exe`.

> **À noter** : il n'est pas possible de générer un `.exe` Windows depuis un Mac ni l'inverse. Si tu n'as pas de machine Windows, options possibles : VM (Parallels, VMware), Boot Camp, ou GitHub Actions pour build dans le cloud.

### Distribution

Une fois compilées, les deux applications peuvent être distribuées par :
- E-mail (attention à la taille, 30-50 Mo)
- Google Drive / Dropbox / WeTransfer
- Page GitHub Releases
- Site web

L'utilisateur final n'a **rien à installer**, juste à double-cliquer sur le fichier reçu.

### Mode développement / CLI

Pour tester en mode développement sans compiler :

```bash
# macOS
python3.12 rcs_scraper.py

# Windows
py rcs_scraper.py
```

Trois modes sont disponibles selon les arguments :
- Sans argument → interface graphique
- `--interactive` ou `-i` → mode interactif texte (questions/réponses dans le terminal)
- Avec des paramètres (`--greffe`, `--annee`, ...) → mode ligne de commande direct

Exemples CLI :

```bash
python3.12 rcs_scraper.py --greffe "Nosy Be" --annee 2025 --forme SARL
python3.12 rcs_scraper.py --greffe "Antananarivo" --annee 2026 --forme SA-AG --type B
python3.12 rcs_scraper.py --greffe "Fianarantsoa" --limit 5  # test rapide
```

| Argument | Description | Défaut |
|---|---|---|
| `--greffe` | Nom du greffe (ville) | `Nosy Be` |
| `--annee` | Année d'immatriculation | `2025` |
| `--forme` | Code de la forme juridique | `SARL` |
| `--type` | Type d'assujetti (A/B/C/D/E) | `B` |
| `--out` | Chemin du fichier Excel de sortie | `~/Desktop/rcs_<ville>_<année>_<forme>.xlsx` |
| `--sleep` | Pause entre requêtes (secondes) | `1.0` |
| `--limit` | Limiter à N entreprises (0 = tout) | `0` |
| `--interactive` / `-i` | Forcer le mode interactif texte | |
| `--gui` | Forcer l'interface graphique | |

## Valeurs disponibles

### Greffes (39)

Ambanja, Ambatolampy, Ambatondrazaka, Ambilobe, Ambositra, Ambovombe, Ampanihy, Analalava, Ankazoabo, Ankazobe, Antalaha, Antananarivo, Antsirabe, Antsiranana, Antsohihy, Arivonimamo, Avaradrano, Betroka, Farafangana, Fianarantsoa, Ihosy, Ikongo, Maevatanana, Mahajanga, Maintirano, Mampikony, Manakara, Mananjary, Mandritsara, Maroantsetra, Miandrivazo, Miarinarivo, Moramanga, Morombe, Morondava, Nosy Be, Port Berger, Sainte Marie, Taolagnaro, Toamasina, Toliara, Tsiroanomandidy, Vatomandry.

### Formes juridiques courantes

`SARL`, `SARLU`, `SA-AG`, `SA-CA`, `SAU`, `EI`, `SNC`, `SCS`, `GIE`, `ONG`, `ASSOCIATION`.

### Types d'assujetti

| Code | Description |
|---|---|
| A | Personne physique |
| B | Personne morale *(défaut)* |
| C | Groupement d'intérêt économique |
| D | Personne morale autre qu'un GIE |
| E | Institution de microfinance |

## Format du fichier Excel généré

Deux feuilles :

**Feuille « <Nom du greffe> »** avec les colonnes :

- Immatriculation, Dénomination, Adresse
- Date de début, Date d'immatriculation
- Capital, Forme juridique, Activité, Type d'entreprise
- Dirigeant — Fonction, Dirigeant — Nom et prénoms

**Feuille « Info »** avec les métadonnées (paramètres de recherche, nombre d'entreprises, date d'extraction, source).

## Structure du projet

```
rcs/
├── rcs_scraper.py          Script principal
├── requirements.txt        Dépendances Python
├── build_mac.sh            Script de build macOS (→ .app)
├── build_windows.bat       Script de build Windows (→ .exe)
└── README.md               Ce fichier
```

## Source

Site officiel du RCS Madagascar : https://www.rcsmada.mg
