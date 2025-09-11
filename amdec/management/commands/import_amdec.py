"""
Commande Django pour importer des fichiers Excel AMDEC.

Cette commande permet d'importer des projets AMDEC depuis des fichiers Excel
en ligne de commande, avec validation et rapport détaillé.

Usage:
    python manage.py import_amdec fichier.xlsx --user=admin
    python manage.py import_amdec *.xlsx --user=admin --batch
    python manage.py import_amdec fichier.xlsx --user=admin --dry-run
    python manage.py import_amdec dossier/ --user=admin --recursive

Standards:
    - Messages colorisés pour la lisibilité
    - Gestion d'erreurs robuste
    - Mode dry-run pour tester sans importer
    - Support du batch processing
    - Rapport détaillé des imports
"""

import os
import sys
from pathlib import Path
from typing import List, Dict, Any, Optional
import logging
from datetime import datetime

from django.core.management.base import BaseCommand, CommandError
from django.contrib.auth.models import User
from django.db import transaction
from django.utils import timezone

# Pour la colorisation des messages
from django.core.management.color import make_style
from django.core.management import color

# Import des modules locaux
try:
    from ...utils.excel_handler import ExcelImporter
    from ...utils.validators import validate_excel_file
    from ...models import AMDECProject
except ImportError as e:
    # Pour le développement, si les modules ne sont pas encore créés
    print(f"⚠️  Import error: {e}")
    print("Assurez-vous que les modules excel_handler et validators sont créés.")


logger = logging.getLogger(__name__)


class Command(BaseCommand):
    """
    Commande pour importer des fichiers Excel AMDEC.
    
    Cette commande supporte l'import unitaire et en batch,
    avec validation complète et rapport détaillé.
    """
    
    help = """
    Importe des projets AMDEC depuis des fichiers Excel.
    
    Exemples:
        # Import simple
        python manage.py import_amdec mon_fichier.xlsx --user=admin
        
        # Import multiple
        python manage.py import_amdec *.xlsx --user=admin --batch
        
        # Test sans import réel
        python manage.py import_amdec fichier.xlsx --user=admin --dry-run
        
        # Import récursif d'un dossier
        python manage.py import_amdec /chemin/dossier/ --user=admin --recursive
        
        # Mode verbeux avec détails
        python manage.py import_amdec fichier.xlsx --user=admin --verbose
    """
    
    def __init__(self, *args, **kwargs):
        """Initialise la commande avec les styles de couleur."""
        super().__init__(*args, **kwargs)
        self.style_success = make_style(opts=('bold',), fg='green')
        self.style_error = make_style(opts=('bold',), fg='red')
        self.style_warning = make_style(opts=('bold',), fg='yellow')
        self.style_info = make_style(opts=('bold',), fg='blue')
        self.style_notice = make_style(fg='cyan')
        
        # Statistiques globales
        self.stats = {
            'files_processed': 0,
            'files_success': 0,
            'files_failed': 0,
            'projects_created': 0,
            'failures_imported': 0,
            'warnings': [],
            'errors': []
        }
    
    def add_arguments(self, parser):
        """
        Définit les arguments de la commande.
        
        Args:
            parser: ArgumentParser Django
        """
        # Arguments positionnels
        parser.add_argument(
            'paths',
            nargs='+',
            type=str,
            help='Chemin(s) vers le(s) fichier(s) Excel ou dossier(s) à importer'
        )
        
        # Arguments requis
        parser.add_argument(
            '--user',
            type=str,
            required=True,
            help='Username de l\'utilisateur qui effectue l\'import (requis)'
        )
        
        # Options
        parser.add_argument(
            '--dry-run',
            action='store_true',
            help='Mode test: valide sans importer réellement'
        )
        
        parser.add_argument(
            '--batch',
            action='store_true',
            help='Mode batch: continue même si un fichier échoue'
        )
        
        parser.add_argument(
            '--recursive',
            action='store_true',
            help='Import récursif des sous-dossiers'
        )
        
        parser.add_argument(
            '--verbose',
            action='store_true',
            help='Affiche les détails de chaque import'
        )
        
        parser.add_argument(
            '--skip-validation',
            action='store_true',
            help='Ignore la validation stricte (non recommandé)'
        )
        
        parser.add_argument(
            '--client',
            type=str,
            default=None,
            help='Force le nom du client pour tous les imports'
        )
        
        parser.add_argument(
            '--encoding',
            type=str,
            default='utf-8',
            help='Encodage des fichiers (par défaut: utf-8)'
        )
        
        parser.add_argument(
            '--max-errors',
            type=int,
            default=10,
            help='Nombre maximum d\'erreurs avant arrêt (défaut: 10)'
        )
    
    def handle(self, *args, **options):
        """
        Point d'entrée principal de la commande.
        
        Args:
            *args: Arguments positionnels
            **options: Options de la commande
        """
        # Afficher l'en-tête
        self._print_header()
        
        # Valider l'utilisateur
        try:
            user = self._get_user(options['user'])
        except CommandError as e:
            self.stdout.write(self.style_error(str(e)))
            sys.exit(1)
        
        # Collecter les fichiers à traiter
        files_to_process = self._collect_files(
            options['paths'],
            options.get('recursive', False)
        )
        
        if not files_to_process:
            self.stdout.write(
                self.style_warning("❌ Aucun fichier Excel trouvé à importer")
            )
            sys.exit(1)
        
        # Afficher le résumé avant import
        self._print_pre_import_summary(files_to_process, options)
        
        # Confirmation si pas en dry-run
        if not options['dry_run'] and len(files_to_process) > 5:
            if not self._confirm_import(len(files_to_process)):
                self.stdout.write(self.style_warning("Import annulé"))
                sys.exit(0)
        
        # Traiter chaque fichier
        for filepath in files_to_process:
            try:
                self._process_file(filepath, user, options)
                
                # Vérifier la limite d'erreurs
                if len(self.stats['errors']) >= options['max_errors']:
                    self.stdout.write(
                        self.style_error(
                            f"⛔ Limite d'erreurs atteinte ({options['max_errors']}). Arrêt."
                        )
                    )
                    break
                    
            except Exception as e:
                if not options['batch']:
                    raise
                else:
                    self.stats['errors'].append(str(e))
                    self.stdout.write(
                        self.style_error(f"Erreur sur {filepath}: {e}")
                    )
        
        # Afficher le rapport final
        self._print_final_report(options)
        
        # Code de sortie basé sur le succès
        if self.stats['files_failed'] > 0:
            sys.exit(1)
        sys.exit(0)
    
    def _print_header(self):
        """Affiche l'en-tête de la commande."""
        self.stdout.write("")
        self.stdout.write(self.style_info("=" * 60))
        self.stdout.write(self.style_info("    📊 IMPORT EXCEL AMDEC - DJANGO"))
        self.stdout.write(self.style_info("=" * 60))
        self.stdout.write(
            self.style_notice(
                f"    Date: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}"
            )
        )
        self.stdout.write(self.style_info("=" * 60))
        self.stdout.write("")
    
    def _get_user(self, username: str) -> User:
        """
        Récupère l'utilisateur Django.
        
        Args:
            username: Nom d'utilisateur
            
        Returns:
            User: Instance de l'utilisateur
            
        Raises:
            CommandError: Si l'utilisateur n'existe pas
        """
        try:
            user = User.objects.get(username=username)
            self.stdout.write(
                self.style_success(f"✅ Utilisateur trouvé: {username}")
            )
            return user
        except User.DoesNotExist:
            raise CommandError(
                f"❌ Utilisateur '{username}' introuvable. "
                f"Utilisez un username valide ou créez l'utilisateur d'abord."
            )
    
    def _collect_files(self, paths: List[str], recursive: bool) -> List[Path]:
        """
        Collecte tous les fichiers Excel à traiter.
        
        Args:
            paths: Liste des chemins fournis
            recursive: Si True, parcourt les sous-dossiers
            
        Returns:
            list: Liste des fichiers Path à traiter
        """
        files = []
        
        for path_str in paths:
            path = Path(path_str)
            
            if not path.exists():
                self.stdout.write(
                    self.style_warning(f"⚠️  Chemin inexistant: {path}")
                )
                continue
            
            if path.is_file():
                if path.suffix.lower() in ['.xlsx', '.xls']:
                    files.append(path)
                else:
                    self.stdout.write(
                        self.style_warning(
                            f"⚠️  Fichier ignoré (pas Excel): {path}"
                        )
                    )
            elif path.is_dir():
                # Recherche dans le dossier
                pattern = '**/*.xls*' if recursive else '*.xls*'
                excel_files = list(path.glob(pattern))
                
                if excel_files:
                    files.extend(excel_files)
                    self.stdout.write(
                        self.style_notice(
                            f"📁 {len(excel_files)} fichiers trouvés dans {path}"
                        )
                    )
                else:
                    self.stdout.write(
                        self.style_warning(
                            f"⚠️  Aucun fichier Excel dans {path}"
                        )
                    )
        
        # Éliminer les doublons
        files = list(set(files))
        files.sort()
        
        return files
    
    def _print_pre_import_summary(self, files: List[Path], options: Dict[str, Any]):
        """
        Affiche le résumé avant import.
        
        Args:
            files: Liste des fichiers à traiter
            options: Options de la commande
        """
        self.stdout.write("")
        self.stdout.write(self.style_info("📋 RÉSUMÉ PRÉ-IMPORT"))
        self.stdout.write(self.style_info("-" * 40))
        
        self.stdout.write(f"  • Fichiers à traiter: {len(files)}")
        self.stdout.write(f"  • Mode: {'TEST (dry-run)' if options['dry_run'] else 'IMPORT RÉEL'}")
        self.stdout.write(f"  • Batch: {'Oui' if options['batch'] else 'Non'}")
        
        if options.get('client'):
            self.stdout.write(f"  • Client forcé: {options['client']}")
        
        self.stdout.write("")
        
        # Lister les premiers fichiers
        if len(files) <= 10:
            self.stdout.write("  Fichiers:")
            for f in files:
                self.stdout.write(f"    - {f.name}")
        else:
            self.stdout.write(f"  Premiers fichiers:")
            for f in files[:5]:
                self.stdout.write(f"    - {f.name}")
            self.stdout.write(f"    ... et {len(files) - 5} autres")
        
        self.stdout.write("")
    
    def _confirm_import(self, count: int) -> bool:
        """
        Demande confirmation pour l'import.
        
        Args:
            count: Nombre de fichiers
            
        Returns:
            bool: True si confirmé
        """
        self.stdout.write(
            self.style_warning(
                f"⚠️  Vous allez importer {count} fichiers. Continuer? [O/n] "
            ),
            ending=''
        )
        
        response = input().lower()
        return response in ['', 'o', 'oui', 'y', 'yes']
    
    def _process_file(self, filepath: Path, user: User, options: Dict[str, Any]):
        """
        Traite un fichier Excel.
        
        Args:
            filepath: Chemin du fichier
            user: Utilisateur Django
            options: Options de la commande
        """
        self.stats['files_processed'] += 1
        
        self.stdout.write("")
        self.stdout.write(self.style_info(f"📄 Traitement: {filepath.name}"))
        self.stdout.write("-" * 40)
        
        try:
            # Validation du fichier
            if not options.get('skip_validation'):
                with open(filepath, 'rb') as f:
                    try:
                        validate_excel_file(f)
                        if options['verbose']:
                            self.stdout.write("  ✓ Validation du fichier OK")
                    except Exception as e:
                        raise CommandError(f"Validation échouée: {e}")
            
            # Import du fichier
            importer = ExcelImporter()
            
            with open(filepath, 'rb') as f:
                # Parser le fichier
                data = importer.parse_excel(f)
                
                if options['verbose']:
                    self._print_parsed_data(data)
                
                # Forcer le client si spécifié
                if options.get('client'):
                    data['metadata']['client'] = options['client']
                
                # Mode dry-run: afficher sans importer
                if options['dry_run']:
                    self.stdout.write(
                        self.style_notice("  🔍 MODE TEST - Pas d'import réel")
                    )
                    self._print_dry_run_summary(data)
                    self.stats['files_success'] += 1
                    return
                
                # Import réel avec transaction
                with transaction.atomic():
                    project = importer.create_project(data, user)
                    
                    # Succès
                    self.stats['files_success'] += 1
                    self.stats['projects_created'] += 1
                    self.stats['failures_imported'] += importer.imported_count
                    
                    self.stdout.write(
                        self.style_success(
                            f"  ✅ Projet créé: {project.reference}"
                        )
                    )
                    self.stdout.write(
                        f"     - Nom: {project.name}"
                    )
                    self.stdout.write(
                        f"     - Défaillances: {importer.imported_count}"
                    )
                    self.stdout.write(
                        f"     - ID: {project.id}"
                    )
                    
                    # Ajouter les warnings
                    if importer.warnings:
                        self.stats['warnings'].extend(importer.warnings)
                        if options['verbose']:
                            for warning in importer.warnings:
                                self.stdout.write(
                                    self.style_warning(f"     ⚠️  {warning}")
                                )
                    
        except Exception as e:
            self.stats['files_failed'] += 1
            self.stats['errors'].append(f"{filepath.name}: {str(e)}")
            
            self.stdout.write(
                self.style_error(f"  ❌ Échec: {str(e)}")
            )
            
            if not options['batch']:
                raise CommandError(f"Import échoué pour {filepath.name}: {e}")
    
    def _print_parsed_data(self, data: Dict[str, Any]):
        """
        Affiche les données parsées en mode verbose.
        
        Args:
            data: Données parsées
        """
        metadata = data.get('metadata', {})
        failures = data.get('failures', [])
        summary = data.get('summary', {})
        
        self.stdout.write("  📊 Données extraites:")
        self.stdout.write(f"     - Système: {metadata.get('system_name', 'N/A')}")
        self.stdout.write(f"     - Client: {metadata.get('client', 'N/A')}")
        self.stdout.write(f"     - Date: {metadata.get('analysis_date', 'N/A')}")
        self.stdout.write(f"     - Équipe: {metadata.get('team_members', 'N/A')}")
        self.stdout.write(f"     - Défaillances: {len(failures)}")
        self.stdout.write(f"     - Criticité élevée: {summary.get('high', 0)}")
        self.stdout.write(f"     - Criticité modérée: {summary.get('medium', 0)}")
        self.stdout.write(f"     - Criticité faible: {summary.get('low', 0)}")
    
    def _print_dry_run_summary(self, data: Dict[str, Any]):
        """
        Affiche le résumé en mode dry-run.
        
        Args:
            data: Données qui seraient importées
        """
        metadata = data.get('metadata', {})
        failures = data.get('failures', [])
        
        self.stdout.write("  📝 Données qui seraient importées:")
        self.stdout.write(f"     - Projet: {metadata.get('system_name', 'N/A')}")
        self.stdout.write(f"     - Référence: {metadata.get('reference', 'N/A')}")
        self.stdout.write(f"     - {len(failures)} défaillances")
        
        # Afficher quelques défaillances
        if failures and len(failures) <= 3:
            for idx, failure in enumerate(failures, 1):
                self.stdout.write(
                    f"       {idx}. {failure.get('component', '')} - "
                    f"{failure.get('failure_mode', '')}"
                )
        elif failures:
            for idx, failure in enumerate(failures[:2], 1):
                self.stdout.write(
                    f"       {idx}. {failure.get('component', '')} - "
                    f"{failure.get('failure_mode', '')}"
                )
            self.stdout.write(f"       ... et {len(failures) - 2} autres")
    
    def _print_final_report(self, options: Dict[str, Any]):
        """
        Affiche le rapport final.
        
        Args:
            options: Options de la commande
        """
        self.stdout.write("")
        self.stdout.write(self.style_info("=" * 60))
        self.stdout.write(self.style_info("📈 RAPPORT FINAL"))
        self.stdout.write(self.style_info("=" * 60))
        
        # Statistiques générales
        self.stdout.write("")
        self.stdout.write("📊 Statistiques:")
        self.stdout.write(f"  • Fichiers traités: {self.stats['files_processed']}")
        self.stdout.write(
            self.style_success(f"  • Réussis: {self.stats['files_success']}")
        )
        
        if self.stats['files_failed'] > 0:
            self.stdout.write(
                self.style_error(f"  • Échoués: {self.stats['files_failed']}")
            )
        
        if not options['dry_run']:
            self.stdout.write("")
            self.stdout.write("💾 Données importées:")
            self.stdout.write(f"  • Projets créés: {self.stats['projects_created']}")
            self.stdout.write(f"  • Défaillances importées: {self.stats['failures_imported']}")
        
        # Warnings
        if self.stats['warnings']:
            self.stdout.write("")
            self.stdout.write(
                self.style_warning(f"⚠️  Avertissements ({len(self.stats['warnings'])})")
            )
            for warning in self.stats['warnings'][:5]:
                self.stdout.write(f"  • {warning}")
            if len(self.stats['warnings']) > 5:
                self.stdout.write(
                    f"  ... et {len(self.stats['warnings']) - 5} autres"
                )
        
        # Erreurs
        if self.stats['errors']:
            self.stdout.write("")
            self.stdout.write(
                self.style_error(f"❌ Erreurs ({len(self.stats['errors'])})")
            )
            for error in self.stats['errors'][:5]:
                self.stdout.write(f"  • {error}")
            if len(self.stats['errors']) > 5:
                self.stdout.write(
                    f"  ... et {len(self.stats['errors']) - 5} autres"
                )
        
        # Résumé final
        self.stdout.write("")
        self.stdout.write(self.style_info("=" * 60))
        
        if self.stats['files_failed'] == 0:
            self.stdout.write(
                self.style_success("✅ IMPORT TERMINÉ AVEC SUCCÈS")
            )
        elif self.stats['files_success'] > 0:
            self.stdout.write(
                self.style_warning("⚠️  IMPORT TERMINÉ AVEC DES ERREURS")
            )
        else:
            self.stdout.write(
                self.style_error("❌ IMPORT ÉCHOUÉ")
            )
        
        self.stdout.write(self.style_info("=" * 60))
        self.stdout.write("")
