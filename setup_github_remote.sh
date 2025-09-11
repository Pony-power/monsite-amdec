#!/bin/bash

echo "🔧 Configuration du remote GitHub"
echo "================================"

# Charger les variables
if [ -f .env ]; then
    source .env
else
    echo "❌ Fichier .env introuvable !"
    exit 1
fi

# Vérifier les variables
if [ -z "$GITHUB_USERNAME" ] || [ -z "$GITHUB_TOKEN" ]; then
    echo "❌ Variables GITHUB_USERNAME ou GITHUB_TOKEN manquantes dans .env"
    exit 1
fi

echo "✅ Variables chargées"
echo "   Username: $GITHUB_USERNAME"
echo "   Token: ${GITHUB_TOKEN:0:10}..."

# Supprimer l'ancien remote s'il existe
if git remote | grep -q origin; then
    echo "⚠️  Remote origin existe, suppression..."
    git remote remove origin
fi

# Ajouter le nouveau remote
echo "📎 Ajout du remote origin..."
git remote add origin https://github.com/$GITHUB_USERNAME/monsite-amdec.git

# Vérifier
echo ""
echo "✅ Configuration terminée !"
echo "Remote configuré :"
git remote -v

echo ""
echo "📤 Pour faire votre premier push :"
echo "   git add ."
echo "   git commit -m 'Initial commit'"
echo "   git push -u origin main"
