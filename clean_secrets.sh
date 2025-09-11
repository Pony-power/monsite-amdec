#!/bin/bash

echo "🔒 Nettoyage des secrets Git..."

# 1. Retirer tous les fichiers sensibles
echo "📝 Retrait des fichiers sensibles..."
git rm --cached .env 2>/dev/null
git rm --cached *.key 2>/dev/null
git rm --cached *.token 2>/dev/null

# 2. Créer/mettre à jour .gitignore
echo "📋 Mise à jour .gitignore..."
cat > .gitignore << 'EOF'
# SECRETS - NE JAMAIS COMMIT
.env
.env.*
*.key
*.token
*.pem
secrets/
credentials/

# Python
__pycache__/
*.py[cod]
*.pyc
venv/
env/

# Django
*.log
db.sqlite3
media/
staticfiles/

# IDE
.vscode/
.idea/
*.swp
*.swo

# OS
.DS_Store
Thumbs.db
EOF

# 3. Commit les changements
echo "💾 Création d'un commit propre..."
git add .gitignore
git add -A
git status

echo ""
echo "⚠️  VÉRIFIEZ que .env n'apparaît PAS dans les fichiers à commit !"
echo "Continuer ? (o/n)"
read response

if [ "$response" = "o" ]; then
    git commit -m "Configuration projet AMDEC Django (sans secrets)"
    echo "✅ Commit propre créé"
    echo ""
    echo "📤 Pour envoyer : git push -u origin main --force"
else
    echo "❌ Annulé. Vérifiez et recommencez."
fi
