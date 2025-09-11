from django.db import models
from django.contrib.auth.models import User
from django.core.validators import MinValueValidator, MaxValueValidator
from django.utils.text import slugify
from django.urls import reverse
from django.utils import timezone
import os


class AMDECProject(models.Model):
    """
    Modèle principal pour un projet AMDEC
    Contient les informations générales du projet et ses métadonnées
    """
    # Informations de base du projet
    name = models.CharField(
        max_length=200,
        verbose_name="Nom du projet",
        help_text="Nom descriptif du système ou processus analysé"
    )
    
    reference = models.CharField(
        max_length=50,
        unique=True,
        verbose_name="Référence",
        help_text="Code de référence unique du projet (ex: AMDEC-2025-001)"
    )
    
    slug = models.SlugField(
        max_length=250,
        unique=True,
        verbose_name="URL",
        help_text="Identifiant URL unique généré automatiquement"
    )
    
    client = models.CharField(
        max_length=200,
        verbose_name="Client",
        help_text="Nom du client ou de l'entreprise"
    )
    
    analysis_date = models.DateField(
        default=timezone.now,
        verbose_name="Date d'analyse",
        help_text="Date de réalisation de l'analyse AMDEC"
    )
    
    team_members = models.TextField(
        verbose_name="Équipe AMDEC",
        help_text="Liste des membres de l'équipe d'analyse (noms séparés par des virgules)"
    )
    
    objective = models.TextField(
        verbose_name="Objectif",
        help_text="Objectif principal de cette analyse AMDEC"
    )
    
    description = models.TextField(
        blank=True,
        null=True,
        verbose_name="Description",
        help_text="Description détaillée du projet et de son contexte"
    )
    
    # Logo client (relation OneToOne ou ForeignKey)
    logo = models.ForeignKey(
        'AMDECLogo',
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='projects',
        verbose_name="Logo client"
    )
    
    # Métadonnées et traçabilité
    created_by = models.ForeignKey(
        User,
        on_delete=models.PROTECT,
        related_name='amdec_projects_created',
        verbose_name="Créé par"
    )
    
    modified_by = models.ForeignKey(
        User,
        on_delete=models.PROTECT,
        related_name='amdec_projects_modified',
        verbose_name="Modifié par",
        null=True,
        blank=True
    )
    
    created_at = models.DateTimeField(
        auto_now_add=True,
        verbose_name="Date de création"
    )
    
    updated_at = models.DateTimeField(
        auto_now=True,
        verbose_name="Dernière modification"
    )
    
    is_active = models.BooleanField(
        default=True,
        verbose_name="Actif",
        help_text="Décocher pour archiver le projet"
    )
    
    class Meta:
        ordering = ['-analysis_date', '-created_at']
        verbose_name = "Projet AMDEC"
        verbose_name_plural = "Projets AMDEC"
        indexes = [
            models.Index(fields=['-analysis_date']),
            models.Index(fields=['slug']),
            models.Index(fields=['reference']),
        ]
    
    def __str__(self):
        return f"{self.reference} - {self.name}"
    
    def save(self, *args, **kwargs):
        """Override save pour générer automatiquement le slug"""
        if not self.slug:
            base_slug = slugify(f"{self.reference}-{self.name}")
            slug = base_slug
            counter = 1
            while AMDECProject.objects.filter(slug=slug).exists():
                slug = f"{base_slug}-{counter}"
                counter += 1
            self.slug = slug
        super().save(*args, **kwargs)
    
    def get_absolute_url(self):
        """URL canonique du projet"""
        return reverse('amdec:project_detail', kwargs={'slug': self.slug})
    
    @property
    def total_failures(self):
        """Nombre total de défaillances dans le projet"""
        return self.failures.count()
    
    @property
    def high_criticality_count(self):
        """Nombre de défaillances à criticité élevée"""
        return self.failures.filter(
            models.Q(gravity__gte=1) & models.Q(occurrence__gte=1) & models.Q(detection__gte=1)
        ).filter(
            models.Q(gravity__isnull=False) & models.Q(occurrence__isnull=False) & models.Q(detection__isnull=False)
        ).annotate(
            crit=models.F('gravity') * models.F('occurrence') * models.F('detection')
        ).filter(crit__gt=120).count()
    
    @property
    def medium_criticality_count(self):
        """Nombre de défaillances à criticité modérée"""
        return len([f for f in self.failures.all() if 61 <= f.criticality <= 120])
    
    @property
    def low_criticality_count(self):
        """Nombre de défaillances à criticité faible"""
        return len([f for f in self.failures.all() if f.criticality <= 60])


class FailureMode(models.Model):
    """
    Modèle pour un mode de défaillance dans une analyse AMDEC
    Calcule automatiquement la criticité basée sur G x O x D
    """
    
    project = models.ForeignKey(
        AMDECProject,
        on_delete=models.CASCADE,
        related_name='failures',
        verbose_name="Projet AMDEC"
    )
    
    # Informations de la défaillance
    component = models.CharField(
        max_length=200,
        verbose_name="Composant",
        help_text="Composant ou élément du système concerné"
    )
    
    failure_mode = models.CharField(
        max_length=300,
        verbose_name="Mode de défaillance",
        help_text="Description du mode de défaillance"
    )
    
    potential_cause = models.TextField(
        verbose_name="Cause potentielle",
        help_text="Cause(s) potentielle(s) de la défaillance"
    )
    
    effect = models.TextField(
        verbose_name="Effet",
        help_text="Effet(s) de la défaillance sur le système"
    )
    
    # Évaluation (notes de 1 à 10)
    gravity = models.IntegerField(
        validators=[MinValueValidator(1), MaxValueValidator(10)],
        verbose_name="Gravité (G)",
        help_text="Note de 1 (négligeable) à 10 (catastrophique)"
    )
    
    occurrence = models.IntegerField(
        validators=[MinValueValidator(1), MaxValueValidator(10)],
        verbose_name="Occurrence (O)",
        help_text="Note de 1 (très rare) à 10 (permanente)"
    )
    
    detection = models.IntegerField(
        validators=[MinValueValidator(1), MaxValueValidator(10)],
        verbose_name="Détection (D)",
        help_text="Note de 1 (très facile) à 10 (impossible)"
    )
    
    # Actions et recommandations
    preventive_actions = models.TextField(
        blank=True,
        null=True,
        verbose_name="Actions préventives",
        help_text="Actions recommandées pour prévenir ou atténuer la défaillance"
    )
    
    responsible = models.CharField(
        max_length=200,
        blank=True,
        null=True,
        verbose_name="Responsable",
        help_text="Personne ou service responsable du suivi"
    )
    
    deadline = models.DateField(
        blank=True,
        null=True,
        verbose_name="Échéance",
        help_text="Date limite pour la mise en œuvre des actions"
    )
    
    status = models.CharField(
        max_length=30,
        choices=[
            ('PENDING', 'En attente'),
            ('IN_PROGRESS', 'En cours'),
            ('COMPLETED', 'Terminé'),
            ('CANCELLED', 'Annulé'),
        ],
        default='PENDING',
        verbose_name="Statut"
    )
    
    # Ordre d'affichage
    order = models.IntegerField(
        default=0,
        verbose_name="Ordre",
        help_text="Ordre d'affichage dans le tableau"
    )
    
    # Métadonnées
    created_at = models.DateTimeField(
        auto_now_add=True,
        verbose_name="Date de création"
    )
    
    updated_at = models.DateTimeField(
        auto_now=True,
        verbose_name="Dernière modification"
    )
    
    notes = models.TextField(
        blank=True,
        null=True,
        verbose_name="Notes",
        help_text="Notes et commentaires additionnels"
    )
    
    class Meta:
        ordering = ['-gravity', '-occurrence', '-detection', 'order']  # Tri par criticité décroissante
        verbose_name = "Mode de défaillance"
        verbose_name_plural = "Modes de défaillance"
        indexes = [
            models.Index(fields=['project', 'order']),
            models.Index(fields=['gravity', 'occurrence', 'detection']),
        ]
        unique_together = [['project', 'component', 'failure_mode']]  # Éviter les doublons
    
    def __str__(self):
        return f"{self.component} - {self.failure_mode} (C={self.criticality})"
    
    @property
    def criticality(self):
        """Calcul automatique de la criticité (G × O × D)"""
        return self.gravity * self.occurrence * self.detection
    
    @property
    def criticality_level(self):
        """
        Niveau de criticité basé sur le score
        - FAIBLE: 1-60
        - MODÉRÉE: 61-120
        - ÉLEVÉE: >120
        """
        crit = self.criticality
        if crit <= 60:
            return "FAIBLE"
        elif crit <= 120:
            return "MODÉRÉE"
        else:
            return "ÉLEVÉE"
    
    @property
    def criticality_color(self):
        """Couleur associée au niveau de criticité pour l'affichage"""
        level = self.criticality_level
        colors = {
            "FAIBLE": "#27ae60",    # Vert
            "MODÉRÉE": "#f39c12",   # Orange
            "ÉLEVÉE": "#e74c3c"     # Rouge
        }
        return colors.get(level, "#95a5a6")  # Gris par défaut
    
    @property
    def is_critical(self):
        """Indique si la défaillance est critique (criticité > 120)"""
        return self.criticality > 120
    
    def get_absolute_url(self):
        """URL canonique de la défaillance"""
        return reverse('amdec:failure_detail', kwargs={
            'project_slug': self.project.slug,
            'pk': self.pk
        })
    
    def clean(self):
        """Validation personnalisée"""
        from django.core.exceptions import ValidationError
        
        # Vérifier que les notes sont bien entre 1 et 10
        for field, value in [
            ('gravity', self.gravity),
            ('occurrence', self.occurrence),
            ('detection', self.detection)
        ]:
            if value is not None and not (1 <= value <= 10):
                raise ValidationError({
                    field: f"La valeur doit être comprise entre 1 et 10 (valeur actuelle: {value})"
                })


class AMDECLogo(models.Model):
    """
    Modèle pour stocker les logos des clients
    """
    
    def logo_upload_path(instance, filename):
        """Génère le chemin d'upload pour le logo"""
        ext = filename.split('.')[-1]
        filename = f"logo_{instance.name}_{timezone.now().strftime('%Y%m%d_%H%M%S')}.{ext}"
        return os.path.join('amdec', 'logos', filename)
    
    name = models.CharField(
        max_length=100,
        verbose_name="Nom du logo",
        help_text="Nom descriptif du logo (ex: Logo principal, Logo secondaire)"
    )
    
    image = models.ImageField(
        upload_to=logo_upload_path,
        verbose_name="Fichier image",
        help_text="Formats acceptés: JPG, PNG, SVG (max 5MB)"
    )
    
    client_name = models.CharField(
        max_length=200,
        blank=True,
        null=True,
        verbose_name="Nom du client",
        help_text="Nom du client associé à ce logo"
    )
    
    is_default = models.BooleanField(
        default=False,
        verbose_name="Logo par défaut",
        help_text="Utiliser ce logo par défaut pour les nouveaux projets"
    )
    
    uploaded_by = models.ForeignKey(
        User,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='amdec_logos',
        verbose_name="Uploadé par"
    )
    
    created_at = models.DateTimeField(
        auto_now_add=True,
        verbose_name="Date d'upload"
    )
    
    class Meta:
        ordering = ['-is_default', '-created_at']
        verbose_name = "Logo AMDEC"
        verbose_name_plural = "Logos AMDEC"
    
    def __str__(self):
        return f"{self.name} ({self.client_name or 'Sans client'})"
    
    def save(self, *args, **kwargs):
        """Override save pour gérer le logo par défaut unique"""
        if self.is_default:
            # S'assurer qu'il n'y a qu'un seul logo par défaut
            AMDECLogo.objects.filter(is_default=True).update(is_default=False)
        super().save(*args, **kwargs)
    
    @property
    def image_url(self):
        """URL complète de l'image"""
        if self.image:
            return self.image.url
        return None
