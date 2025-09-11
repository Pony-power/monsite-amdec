#!/bin/bash
# Helper Git pour le projet AMDEC

# Charger les variables d'environnement
export $(cat .env | xargs)

# Fonction pour configurer Git
setup_git() {
    echo "🔧 Configuration Git..."
    git init
    git remote add origin $GITHUB_REPO_URL 2>/dev/null || git remote set-url origin $GITHUB_REPO_URL
    echo "✅ Remote configuré : $GITHUB_REPO_URL"
}

# Fonction pour sauvegarder
save() {
    MESSAGE=${1:-"Mise à jour automatique"}
    echo "💾 Sauvegarde : $MESSAGE"
    git add .
    git commit -m "$MESSAGE"
    echo "✅ Sauvegarde locale effectuée"
}

# Fonction pour envoyer sur GitHub
push() {
    echo "📤 Envoi sur GitHub..."
    git push -u origin main
    echo "✅ Envoyé sur GitHub : $GITHUB_REPO_URL"
}

# Fonction pour tout faire d'un coup
deploy() {
    MESSAGE=${1:-"Déploiement $(date +%Y-%m-%d_%H:%M)"}
    save "$MESSAGE"
    push
}

# Menu
case "$1" in
    setup)
        setup_git
        ;;
    save)
        save "$2"
        ;;
    push)
        push
        ;;
    deploy)
        deploy "$2"
        ;;
    *)
        echo "Usage: ./git_helper.sh {setup|save|push|deploy} [message]"
        echo ""
        echo "  setup  : Configure Git avec GitHub"
        echo "  save   : Sauvegarde locale (commit)"
        echo "  push   : Envoie sur GitHub"
        echo "  deploy : Save + Push en une commande"
        echo ""
        echo "Exemple: ./git_helper.sh deploy 'Ajout API REST'"
        ;;
esac
