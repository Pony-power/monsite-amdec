#!/usr/bin/env python3
"""
Automatisation Git pour le projet AMDEC Django.
Gère les commits et push automatiques vers GitHub.
"""

import os
import subprocess
import sys
from pathlib import Path
from datetime import datetime
from typing import Optional, List, Tuple

class Colors:
    """Couleurs pour l'affichage terminal"""
    GREEN = '\033[92m'
    RED = '\033[91m'
    YELLOW = '\033[93m'
    BLUE = '\033[94m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'

class GitAutomation:
    """Gestionnaire automatique Git/GitHub pour le projet AMDEC."""
    
    def __init__(self, repo_path: str = "."):
        """
        Initialise le gestionnaire Git.
        
        Args:
            repo_path: Chemin vers le repository (défaut: dossier actuel)
        """
        self.repo_path = Path(repo_path).resolve()
        self.username = None
        self.token = None
        self.repo_name = "monsite-amdec"
        
        # Charger la configuration
        self.load_env()
        self.verify_git_repo()
        
    def load_env(self):
        """Charge les variables depuis .env"""
        env_file = self.repo_path / ".env"
        
        if not env_file.exists():
            self.print_error("Fichier .env introuvable!")
            self.print_info("Créez un fichier .env avec:")
            print("  GITHUB_USERNAME=votre-username")
            print("  GITHUB_TOKEN=ghp_votre_token")
            sys.exit(1)
        
        # Parser le fichier .env
        with open(env_file) as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith('#'):
                    if '=' in line:
                        key, value = line.split('=', 1)
                        os.environ[key] = value
        
        # Récupérer les variables
        self.username = os.getenv('GITHUB_USERNAME')
        self.token = os.getenv('GITHUB_TOKEN')
        self.repo_name = os.getenv('GITHUB_REPO_NAME', 'monsite-amdec')
        
        # Vérifier les credentials
        if not self.username or not self.token:
            self.print_error("Credentials GitHub manquants dans .env!")
            sys.exit(1)
        
        self.print_success(f"Credentials chargés pour {self.username}")
        
    def verify_git_repo(self):
        """Vérifie que nous sommes dans un repo Git."""
        git_dir = self.repo_path / ".git"
        
        if not git_dir.exists():
            self.print_error(f"Pas de repository Git dans {self.repo_path}")
            if input("Initialiser Git ? (o/n): ").lower() == 'o':
                self.git_init()
            else:
                sys.exit(1)
    
    def run_command(self, cmd: List[str], silent: bool = False) -> Tuple[bool, str, str]:
        """
        Exécute une commande et retourne le résultat.
        
        Args:
            cmd: Commande à exécuter
            silent: Si True, n'affiche pas la sortie
            
        Returns:
            Tuple (succès, stdout, stderr)
        """
        try:
            result = subprocess.run(
                cmd,
                cwd=self.repo_path,
                capture_output=True,
                text=True,
                check=False
            )
            
            if not silent:
                if result.stdout:
                    print(result.stdout)
                if result.stderr and result.returncode != 0:
                    print(result.stderr, file=sys.stderr)
            
            return (result.returncode == 0, result.stdout, result.stderr)
            
        except Exception as e:
            return (False, "", str(e))
    
    def git_init(self):
        """Initialise un nouveau repo Git."""
        self.print_info("Initialisation de Git...")
        success, _, _ = self.run_command(['git', 'init'])
        
        if success:
            self.print_success("Repository Git initialisé")
            # Configuration de base
            self.run_command(['git', 'config', 'user.name', self.username])
            self.run_command(['git', 'config', 'user.email', f'{self.username}@users.noreply.github.com'])
        else:
            self.print_error("Échec de l'initialisation Git")
    
    def git_status(self) -> dict:
        """
        Récupère le statut Git actuel.
        
        Returns:
            Dict avec les fichiers modifiés, ajoutés, supprimés
        """
        self.print_info("Vérification du statut...")
        
        # Statut porcelain pour parsing facile
        success, stdout, _ = self.run_command(['git', 'status', '--porcelain'], silent=True)
        
        status = {
            'modified': [],
            'added': [],
            'deleted': [],
            'untracked': []
        }
        
        if success and stdout:
            for line in stdout.strip().split('\n'):
                if line:
                    status_code = line[:2]
                    filename = line[3:]
                    
                    if status_code == ' M' or status_code == 'M ':
                        status['modified'].append(filename)
                    elif status_code == 'A ':
                        status['added'].append(filename)
                    elif status_code == 'D ':
                        status['deleted'].append(filename)
                    elif status_code == '??':
                        status['untracked'].append(filename)
        
        # Afficher le résumé
        total = sum(len(v) for v in status.values())
        if total > 0:
            print(f"\n📊 Résumé: {total} fichiers modifiés")
            if status['modified']:
                print(f"  📝 Modifiés: {len(status['modified'])}")
            if status['added']:
                print(f"  ➕ Ajoutés: {len(status['added'])}")
            if status['deleted']:
                print(f"  ➖ Supprimés: {len(status['deleted'])}")
            if status['untracked']:
                print(f"  ❓ Non suivis: {len(status['untracked'])}")
        else:
            print("✨ Aucune modification")
        
        return status
    
    def git_add(self, files: Optional[List[str]] = None):
        """
        Ajoute des fichiers au staging.
        
        Args:
            files: Liste des fichiers à ajouter (None = tous)
        """
        if files:
            self.print_info(f"Ajout de {len(files)} fichiers...")
            cmd = ['git', 'add'] + files
        else:
            self.print_info("Ajout de tous les fichiers modifiés...")
            cmd = ['git', 'add', '.']
        
        success, _, _ = self.run_command(cmd)
        
        if success:
            self.print_success("Fichiers ajoutés au staging")
        else:
            self.print_error("Erreur lors de l'ajout des fichiers")
    
    def git_commit(self, message: str = None) -> bool:
        """
        Créé un commit.
        
        Args:
            message: Message de commit
            
        Returns:
            True si succès
        """
        if not message:
            # Message automatique avec timestamp
            message = f"🤖 Auto-commit AMDEC - {datetime.now().strftime('%Y-%m-%d %H:%M')}"
        
        self.print_info(f"Commit: {message}")
        
        success, stdout, stderr = self.run_command(['git', 'commit', '-m', message])
        
        if success:
            # Extraire le hash du commit
            for line in stdout.split('\n'):
                if 'commit' in line.lower():
                    self.print_success(f"Commit créé: {line}")
                    break
            return True
        elif "nothing to commit" in stderr or "nothing to commit" in stdout:
            self.print_warning("Rien à commiter")
            return False
        else:
            self.print_error("Échec du commit")
            return False
    
    def git_push(self, branch: str = "main", force: bool = False) -> bool:
        """
        Push vers GitHub avec authentification automatique.
        
        Args:
            branch: Branche à pusher
            force: Force push si True
            
        Returns:
            True si succès
        """
        self.print_info(f"Push vers GitHub ({branch})...")
        
        # Construire l'URL avec authentification
        remote_url = f"https://{self.username}:{self.token}@github.com/{self.username}/{self.repo_name}.git"
        
        # Commande push
        cmd = ['git', 'push', remote_url, branch]
        if force:
            cmd.insert(2, '-f')
        
        # Masquer le token dans la sortie
        success, stdout, stderr = self.run_command(cmd, silent=True)
        
        # Nettoyer la sortie du token
        clean_output = stderr.replace(self.token, '***TOKEN***')
        
        if success:
            self.print_success(f"✅ Push réussi vers {branch}")
            print(f"🔗 Voir sur: https://github.com/{self.username}/{self.repo_name}")
            return True
        else:
            if "Everything up-to-date" in clean_output:
                self.print_info("Déjà à jour")
                return True
            else:
                self.print_error("Échec du push")
                print(clean_output)
                return False
    
    def deploy(self, message: str = None, files: List[str] = None) -> bool:
        """
        Workflow complet: add, commit, push.
        
        Args:
            message: Message de commit
            files: Fichiers spécifiques (None = tous)
            
        Returns:
            True si succès complet
        """
        print("\n" + "="*50)
        self.print_bold("🚀 DÉPLOIEMENT AUTOMATIQUE")
        print("="*50)
        
        # 1. Vérifier le statut
        status = self.git_status()
        
        if not any(status.values()):
            self.print_warning("Aucune modification à déployer")
            return False
        
        # 2. Ajouter les fichiers
        self.git_add(files)
        
        # 3. Créer le commit
        if not self.git_commit(message):
            return False
        
        # 4. Push vers GitHub
        return self.git_push()
    
    def rollback(self, steps: int = 1):
        """
        Annule les derniers commits locaux.
        
        Args:
            steps: Nombre de commits à annuler
        """
        self.print_warning(f"⚠️ Annulation de {steps} commit(s)")
        
        if input("Confirmer le rollback ? (o/n): ").lower() != 'o':
            print("Annulé")
            return
        
        success, _, _ = self.run_command(['git', 'reset', '--hard', f'HEAD~{steps}'])
        
        if success:
            self.print_success(f"Rollback de {steps} commit(s) effectué")
        else:
            self.print_error("Échec du rollback")
    
    # Méthodes d'affichage
    def print_success(self, msg: str):
        print(f"{Colors.GREEN}✅ {msg}{Colors.ENDC}")
    
    def print_error(self, msg: str):
        print(f"{Colors.RED}❌ {msg}{Colors.ENDC}")
    
    def print_warning(self, msg: str):
        print(f"{Colors.YELLOW}⚠️ {msg}{Colors.ENDC}")
    
    def print_info(self, msg: str):
        print(f"{Colors.BLUE}ℹ️ {msg}{Colors.ENDC}")
    
    def print_bold(self, msg: str):
        print(f"{Colors.BOLD}{msg}{Colors.ENDC}")


def main():
    """Point d'entrée principal du script."""
    import argparse
    
    parser = argparse.ArgumentParser(description='Automatisation Git pour AMDEC')
    parser.add_argument('action', nargs='?', default='deploy',
                       choices=['deploy', 'status', 'push', 'rollback'],
                       help='Action à effectuer')
    parser.add_argument('-m', '--message', help='Message de commit')
    parser.add_argument('-f', '--files', nargs='+', help='Fichiers spécifiques')
    parser.add_argument('--force', action='store_true', help='Force push')
    
    args = parser.parse_args()
    
    # Initialiser le gestionnaire
    try:
        git = GitAutomation()
        
        # Exécuter l'action
        if args.action == 'deploy':
            git.deploy(message=args.message, files=args.files)
        elif args.action == 'status':
            git.git_status()
        elif args.action == 'push':
            git.git_push(force=args.force)
        elif args.action == 'rollback':
            git.rollback()
            
    except KeyboardInterrupt:
        print("\n\n⛔ Interrompu par l'utilisateur")
        sys.exit(1)
    except Exception as e:
        print(f"\n❌ Erreur: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
