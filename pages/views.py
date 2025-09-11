from django.shortcuts import render, get_object_or_404
from django.views.generic import ListView
from .models import Page

def home(request):
    # Page d'accueil
    pages = Page.objects.filter(is_published=True)
    try:
        home_page = Page.objects.get(slug='accueil')
    except Page.DoesNotExist:
        home_page = None
    
    return render(request, 'pages/home.html', {
        'pages': pages,
        'home_page': home_page
    })

def page_detail(request, slug):
    # Affichage d'une page
    page = get_object_or_404(Page, slug=slug, is_published=True)
    pages = Page.objects.filter(is_published=True)
    
    return render(request, 'pages/page_detail.html', {
        'page': page,
        'pages': pages
    })
