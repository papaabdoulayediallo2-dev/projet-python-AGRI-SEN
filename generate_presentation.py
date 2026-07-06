#!/usr/bin/env python3
"""
Script pour générer une présentation PowerPoint du projet AGRI-SEN
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

# Créer une présentation
prs = Presentation()
prs.slide_width = Inches(10)
prs.slide_height = Inches(7.5)

# Définir les couleurs
COULEUR_PRINCIPALE = RGBColor(34, 139, 34)  # Verde oscuro
COULEUR_TEXTE = RGBColor(0, 0, 0)
COULEUR_TITRE = RGBColor(255, 255, 255)

def ajouter_slide_titre(titre, sous_titre):
    """Ajoute une slide de titre"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = COULEUR_PRINCIPALE
    
    # Titre
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(2.5), Inches(9), Inches(1.5))
    title_frame = title_box.text_frame
    title_frame.word_wrap = True
    p = title_frame.paragraphs[0]
    p.text = titre
    p.font.size = Pt(60)
    p.font.bold = True
    p.font.color.rgb = COULEUR_TITRE
    p.alignment = PP_ALIGN.CENTER
    
    # Sous-titre
    subtitle_box = slide.shapes.add_textbox(Inches(0.5), Inches(4), Inches(9), Inches(1))
    subtitle_frame = subtitle_box.text_frame
    p = subtitle_frame.paragraphs[0]
    p.text = sous_titre
    p.font.size = Pt(28)
    p.font.color.rgb = COULEUR_TITRE
    p.alignment = PP_ALIGN.CENTER

def ajouter_slide_contenu(titre, contenu_list):
    """Ajoute une slide avec du contenu"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(245, 245, 245)
    
    # Titre
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
    title_frame = title_box.text_frame
    p = title_frame.paragraphs[0]
    p.text = titre
    p.font.size = Pt(44)
    p.font.bold = True
    p.font.color.rgb = COULEUR_PRINCIPALE
    
    # Contenu
    content_box = slide.shapes.add_textbox(Inches(1), Inches(1.5), Inches(8), Inches(5.5))
    text_frame = content_box.text_frame
    text_frame.word_wrap = True
    
    for i, ligne in enumerate(contenu_list):
        if i == 0:
            p = text_frame.paragraphs[0]
        else:
            p = text_frame.add_paragraph()
        
        p.text = ligne
        p.font.size = Pt(20)
        p.font.color.rgb = COULEUR_TEXTE
        p.space_before = Pt(8)
        p.space_after = Pt(8)
        p.level = 0

# Slide 1 : Titre
ajouter_slide_titre("AGRI-SEN", "Système de Gestion de Coopérative Agricole")

# Slide 2 : Vue d'ensemble
ajouter_slide_contenu("Vue d'ensemble du Projet", [
    "✓ Système de gestion pour une coopérative agricole sénégalaise",
    "✓ Gestion des producteurs et de leurs récoltes",
    "✓ Calcul des statistiques de production",
    "✓ Interface interactive avec menu utilisateur",
    "✓ Développé en Python par 5 membres d'équipe"
])

# Slide 3 : Objectifs
ajouter_slide_contenu("Objectifs du Projet", [
    "• Centraliser les informations sur les producteurs",
    "• Enregistrer et suivre les récoltes par produit",
    "• Analyser la production totale et par produit",
    "• Identifier le meilleur producteur",
    "• Calculer les moyennes de production"
])

# Slide 4 : Architecture - Membre 1
ajouter_slide_contenu("Architecture du Projet - Membre 1", [
    "📁 membre1.py : Gestion des Producteurs",
    "   • Liste globale 'producteurs'",
    "   • ajouter_producteur() - ajoute un nouveau producteur",
    "   • afficher_producteurs() - liste tous les producteurs",
    "   • rechercher_producteur(nom) - recherche par nom"
])

# Slide 5 : Architecture - Membre 2
ajouter_slide_contenu("Architecture du Projet - Membre 2", [
    "📁 membre2.py : Gestion des Récoltes",
    "   • PRODUITS_VALIDES - liste des produits autorisés",
    "   • produit_valide(produit) - valide un produit",
    "   • enregistrer_recolte() - enregistre une récolte",
    "   • afficher_recoltes(nom) - affiche les récoltes"
])

# Slide 6 : Architecture - Membre 3
ajouter_slide_contenu("Architecture du Projet - Membre 3", [
    "📁 Membre3.py : Statistiques de Production",
    "   • production_totale() - calcule la production totale",
    "   • production_par_produit() - détail par type de produit",
    "   • Utilise les données de tous les producteurs",
    "   • Affiche les résultats en kg"
])

# Slide 7 : Architecture - Membre 4
ajouter_slide_contenu("Architecture du Projet - Membre 4", [
    "📁 membre4.py : Analyses Avancées",
    "   • meilleur_producteur(liste) - identifie le meilleur",
    "   • moyenne_producteur(liste) - calcule la moyenne",
    "   • Utilise pour les statistiques globales",
    "   • Affiche les résultats formatés"
])

# Slide 8 : Architecture - Membre 5
ajouter_slide_contenu("Architecture du Projet - Membre 5", [
    "📁 main.py : Menu Principal et Intégration",
    "   • Importe toutes les fonctions des 4 membres",
    "   • Menu interactif avec 10 options",
    "   • Point d'entrée du programme",
    "   • Permet à l'utilisateur d'accéder à toutes les fonctionnalités"
])

# Slide 9 : Flux du Programme
ajouter_slide_contenu("Flux du Programme", [
    "1️⃣ Lancement : python main.py",
    "2️⃣ Affichage du menu interactif",
    "3️⃣ L'utilisateur choisit une option (1-10)",
    "4️⃣ Exécution de la fonction correspondante",
    "5️⃣ Retour au menu pour une autre action"
])

# Slide 10 : Produits Gérés
ajouter_slide_contenu("Produits Autorisés", [
    "🌾 Mil",
    "🌽 Maïs",
    "🍚 Riz",
    "🥜 Arachide"
])

# Slide 11 : Fonctionnalités Principales
ajouter_slide_contenu("Fonctionnalités Principales", [
    "✅ Ajouter et afficher les producteurs",
    "✅ Rechercher un producteur spécifique",
    "✅ Enregistrer les récoltes avec validation",
    "✅ Afficher les récoltes par producteur",
    "✅ Calculer les statistiques de production",
    "✅ Identifier le meilleur producteur"
])

# Slide 12 : Structure des Données
ajouter_slide_contenu("Structure des Données", [
    "Producteur = {",
    "  'nom': string,",
    "  'recoltes': [",
    "    {'nom': string, 'produit': string, 'quantite': float},",
    "    ...",
    "  ]",
    "}"
])

# Slide 13 : Validation et Contrôle
ajouter_slide_contenu("Validation et Contrôle de Qualité", [
    "✓ Validation des noms (non vide, pas de doublons)",
    "✓ Validation des produits (liste fermée)",
    "✓ Validation des quantités (positives, nombre valide)",
    "✓ Gestion d'erreurs complète",
    "✓ Messages clairs à l'utilisateur"
])

# Slide 14 : Technologies Utilisées
ajouter_slide_contenu("Technologies Utilisées", [
    "🐍 Python 3.x - Langage de programmation",
    "📦 Modules standards Python (pas de dépendances externes)",
    "🔧 Dictionnaires et listes pour la structure de données",
    "💾 Données en mémoire (pas de base de données)",
    "📝 Version contrôle : Git & GitHub"
])

# Slide 15 : Avantages du Système
ajouter_slide_contenu("Avantages du Système", [
    "✨ Simple et intuitif à utiliser",
    "📊 Rapports de production détaillés",
    "🔍 Recherche rapide des producteurs",
    "📈 Suivi de la productivité",
    "🛡️ Données validées et cohérentes"
])

# Slide 16 : Conclusion
ajouter_slide_contenu("Conclusion", [
    "✓ Système complet et fonctionnel",
    "✓ Code bien organisé et commenté",
    "✓ Facile à étendre et maintenir",
    "✓ Prêt pour une utilisation réelle",
    "✓ Collaboration d'équipe réussie !"
])

# Sauvegarder la présentation
output_path = "AGRI-SEN_Presentation.pptx"
prs.save(output_path)
print(f"✅ Présentation créée avec succès : {output_path}")
