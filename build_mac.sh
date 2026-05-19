#!/bin/bash
# Script de build pour macOS — produit RCS Madagascar.app
# Usage : ./build_mac.sh

set -e
cd "$(dirname "$0")"

echo "=========================================="
echo "  Build RCS Madagascar pour macOS"
echo "=========================================="
echo ""

# Cherche un Python compatible (évite le Python d'Apple qui a Tkinter cassé)
PYTHON=""
for cmd in python3.13 python3.12 python3.11 python3.10 python3; do
    if command -v "$cmd" >/dev/null 2>&1; then
        path=$(command -v "$cmd")
        if [[ "$path" != *"/CommandLineTools/"* ]] && [[ "$path" != "/usr/bin/$cmd" ]]; then
            PYTHON="$cmd"
            break
        fi
    fi
done

if [ -z "$PYTHON" ]; then
    echo "ERREUR : Aucun Python compatible trouvé sur ce Mac."
    echo ""
    echo "Le Python d'Apple (/usr/bin/python3 ou Command Line Tools) ne marche pas"
    echo "pour compiler une app avec Tkinter."
    echo ""
    echo "Installer un vrai Python :"
    echo "  - Télécharger depuis https://www.python.org/downloads/"
    echo "  - Ou via Homebrew : brew install python-tk@3.12"
    echo ""
    exit 1
fi

PY_PATH=$(command -v "$PYTHON")
echo "Python utilisé : $($PYTHON --version) ($PY_PATH)"
echo ""

# Vérifie que Tkinter marche sur ce Python
if ! $PYTHON -c "import tkinter; r = tkinter.Tk(); r.withdraw(); r.destroy()" 2>/dev/null; then
    echo "ERREUR : Tkinter ne fonctionne pas avec ce Python ($PY_PATH)."
    echo "Installer un Python avec Tkinter fonctionnel (voir ci-dessus)."
    exit 1
fi

echo "Tkinter : OK"
echo ""

echo "Installation des dépendances..."
$PYTHON -m pip install --upgrade pip --quiet
$PYTHON -m pip install -r requirements.txt pyinstaller --quiet
echo "  Dépendances installées."
echo ""

echo "Nettoyage des builds précédents..."
rm -rf build dist
rm -f "RCS Madagascar.spec" rcs_scraper.spec
echo ""

echo "Compilation en cours (peut prendre 1-2 minutes)..."
$PYTHON -m PyInstaller \
    --windowed \
    --onefile \
    --name="RCS Madagascar" \
    --osx-bundle-identifier="mg.rcsmada.extracteur" \
    --noconfirm \
    rcs_scraper.py

echo ""
echo "=========================================="
echo "  Compilation terminée !"
echo "=========================================="
echo ""
echo "Application produite :"
echo "  $(pwd)/dist/RCS Madagascar.app"
echo ""
echo "Cette app peut être distribuée à n'importe quel utilisateur Mac."
echo "Il lui suffit de la copier dans Applications et de double-cliquer."
echo ""

open dist/
