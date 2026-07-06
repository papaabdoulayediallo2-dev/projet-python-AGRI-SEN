#!/usr/bin/env python3
"""
Script pour générer un rapport complet du projet AGRI-SEN en PowerPoint
Contient: Introduction, Analyse, Algorithme, Démonstration, Conclusion
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

# Créer la présentation
prs = Presentation()
prs.slide_width = Inches(10)
prs.slide_height = Inches(7.5)

# Couleurs
COULEUR_PRINCIPALE = RGBColor(34, 139, 34)
COULEUR_TEXTE = RGBColor(0, 0, 0)
COULEUR_TITRE = RGBColor(255, 255, 255)

def add_title_slide(title, subtitle):
    """Ajoute une slide de titre"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = COULEUR_PRINCIPALE
    
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(2.5), Inches(9), Inches(1.5))
    title_frame = title_box.text_frame
    title_frame.word_wrap = True
    p = title_frame.paragraphs[0]
    p.text = title
    p.font.size = Pt(54)
    p.font.bold = True
    p.font.color.rgb = COULEUR_TITRE
    p.alignment = PP_ALIGN.CENTER
    
    subtitle_box = slide.shapes.add_textbox(Inches(0.5), Inches(4), Inches(9), Inches(1))
    subtitle_frame = subtitle_box.text_frame
    p = subtitle_frame.paragraphs[0]
    p.text = subtitle
    p.font.size = Pt(24)
    p.font.color.rgb = COULEUR_TITRE
    p.alignment = PP_ALIGN.CENTER

def add_content_slide(title, content_lines):
    """Ajoute une slide de contenu"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(245, 245, 245)
    
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.4), Inches(9), Inches(0.8))
    title_frame = title_box.text_frame
    p = title_frame.paragraphs[0]
    p.text = title
    p.font.size = Pt(40)
    p.font.bold = True
    p.font.color.rgb = COULEUR_PRINCIPALE
    
    content_box = slide.shapes.add_textbox(Inches(0.7), Inches(1.4), Inches(8.6), Inches(5.7))
    text_frame = content_box.text_frame
    text_frame.word_wrap = True
    
    for i, ligne in enumerate(content_lines):
        if i == 0:
            p = text_frame.paragraphs[0]
        else:
            p = text_frame.add_paragraph()
        
        p.text = ligne
        p.font.size = Pt(16)
        p.font.color.rgb = COULEUR_TEXTE
        p.space_before = Pt(3)
        p.space_after = Pt(3)
        p.line_spacing = 1.1

# PAGE DE COUVERTURE
add_title_slide("RAPPORT DE PROJET", "AGRI-SEN : Système de Gestion de Coopérative Agricole")

# SECTION I : INTRODUCTION
add_title_slide("I. INTRODUCTION", "Contexte et Présentation du Problème")

add_content_slide("Contexte Général", [
    "Le secteur agricole en Afrique de l'Ouest joue un rôle fondamental",
    "dans l'économie et la sécurité alimentaire.",
    "",
    "Les producteurs se regroupent dans des coopératives pour :",
    "- Améliorer leur productivité",
    "- Partager les ressources",
    "- Maximiser leurs revenus",
    "",
    "Cependant, la gestion administrative reste rudimentaire et peu",
    "structurée dans la plupart des coopératives."
])

add_content_slide("Problématique Identifiée", [
    "Les coopératives agricoles font face à plusieurs défis majeurs :",
    "",
    "1. Absence de centralisation des données des producteurs",
    "",
    "2. Difficulté à suivre et enregistrer les récoltes de manière organisée",
    "",
    "3. Manque de visibilité sur la production globale et ses performances",
    "",
    "4. Impossibilité d'identifier les producteurs les plus performants",
    "",
    "5. Absence d'outils d'analyse pour optimiser la gestion"
])

add_content_slide("Solution Proposée : AGRI-SEN", [
    "AGRI-SEN est un système informatique développé pour résoudre",
    "ces problèmes critiques.",
    "",
    "Le système permet :",
    "",
    "- Une gestion complète et intégrée de la coopérative",
    "- Centralisation des données des producteurs et récoltes",
    "- Rapports d'analyse détaillés",
    "- Facilitation de la prise de décision",
    "- Optimisation de la gestion administrative"
])

add_content_slide("Objectifs du Projet", [
    "1. Centraliser et organiser les données des producteurs",
    "   de manière sécurisée et accessible",
    "",
    "2. Permettre l'enregistrement structuré des récoltes",
    "   avec validation complète des données",
    "",
    "3. Générer des rapports détaillés sur la production",
    "   globale et par produit",
    "",
    "4. Identifier les producteurs les plus performants",
    "",
    "5. Faciliter la gestion administrative des coopératives"
])

# SECTION II : ANALYSE
add_title_slide("II. ANALYSE", "Description Détaillée des Fonctionnalités")

add_content_slide("Architecture Générale", [
    "AGRI-SEN est construit selon une architecture modulaire",
    "composée de 5 modules Python indépendants mais",
    "interconnectés.",
    "",
    "Architecture modulaire permet :",
    "",
    "- Une maintenance facile",
    "- La réutilisabilité du code",
    "- Une séparation claire des responsabilités",
    "- Une extensibilité future"
])

add_content_slide("Les 5 Modules du Système", [
    "Module 1 (membre1.py) : Gestion des Producteurs",
    "   - Ajouter, afficher et rechercher des producteurs",
    "",
    "Module 2 (membre2.py) : Gestion des Récoltes",
    "   - Enregistrer et afficher les récoltes avec validation",
    "",
    "Module 3 (Membre3.py) : Statistiques de Production",
    "   - Calculer production totale et par produit",
    "",
    "Module 4 (membre4.py) : Analyses Avancées",
    "   - Meilleur producteur et moyennes de production",
    "",
    "Module 5 (main.py) : Interface Principale",
    "   - Menu interactif et intégration de tous les modules"
])

add_content_slide("Module 1 : Gestion des Producteurs", [
    "Responsabilité :",
    "Gérer le cycle de vie des producteurs de la coopérative",
    "",
    "Fonctions principales :",
    "",
    "1. ajouter_producteur()",
    "   - Ajoute un nouveau producteur avec validation du nom",
    "   - Empêche les doublons",
    "",
    "2. afficher_producteurs()",
    "   - Affiche la liste complète avec nombre de récoltes",
    "",
    "3. rechercher_producteur(nom)",
    "   - Trouve un producteur (recherche insensible à la casse)"
])

add_content_slide("Module 2 : Gestion des Récoltes", [
    "Responsabilité :",
    "Gérer l'enregistrement et le suivi des récoltes",
    "",
    "Système de validation stricte :",
    "- Vérification du producteur (doit exister)",
    "- Validation du produit (liste fermée)",
    "- Vérification de la quantité (positive et nombre valide)",
    "",
    "Produits autorisés : Mil, Maïs, Riz, Arachide",
    "",
    "Fonctions : produit_valide(), enregistrer_recolte(),",
    "afficher_recoltes()"
])

add_content_slide("Module 3 : Statistiques de Production", [
    "Responsabilité :",
    "Fournir une vue globale de la production de la coopérative",
    "",
    "Fonctions principales :",
    "",
    "1. production_totale()",
    "   - Calcule la quantité totale produite en kg",
    "   - Agrège les données de tous les producteurs",
    "",
    "2. production_par_produit()",
    "   - Décompose la production par type de produit",
    "   - Affiche les totaux pour chaque catégorie"
])

add_content_slide("Module 4 : Analyses Avancées", [
    "Responsabilité :",
    "Fournir des analyses détaillées des performances",
    "",
    "Fonctions principales :",
    "",
    "1. meilleur_producteur(liste)",
    "   - Identifie le producteur avec production totale",
    "   - la plus élevée",
    "",
    "2. moyenne_producteur(liste)",
    "   - Calcule la production moyenne par producteur",
    "   - Affiche résultat formaté avec 2 décimales"
])

add_content_slide("Module 5 : Interface Principale", [
    "Responsabilité :",
    "Fournir l'interface utilisateur et intégrer tous les modules",
    "",
    "Menu interactif avec 10 options :",
    "",
    "1. Ajouter un producteur",
    "2. Afficher les producteurs",
    "3. Rechercher un producteur",
    "4. Enregistrer une récolte",
    "5. Afficher les récoltes",
    "6. Production totale",
    "7. Production par produit",
    "8. Meilleur producteur",
    "9. Production moyenne",
    "0. Quitter"
])

# SECTION III : ALGORITHMES
add_title_slide("III. ALGORITHME", "Description des Traitements Réalisés")

add_content_slide("Algorithme d'Enregistrement d'une Récolte", [
    "1. Demander le nom du producteur",
    "2. Rechercher le producteur dans la liste",
    "3. SI producteur n'existe pas : Afficher erreur et quitter",
    "4. Demander le produit récolté",
    "5. SI produit non valide : Afficher erreur et quitter",
    "6. Demander la quantité en kg",
    "7. Essayer convertir quantité en nombre décimal",
    "8. SI quantité <= 0 : Afficher erreur et quitter",
    "9. Créer structure récolte avec nom, produit, quantité",
    "10. Ajouter récolte à la liste du producteur",
    "11. Afficher message de succès"
])

add_content_slide("Algorithme : Production Totale", [
    "1. Initialiser total = 0",
    "",
    "2. POUR CHAQUE producteur DANS producteurs",
    "     POUR CHAQUE récolte DANS producteur.recoltes",
    "       Ajouter récolte.quantite au total",
    "",
    "3. Afficher \"Production totale : \" + total + \" kg\"",
    "",
    "4. Retourner total",
    "",
    "Complexité : O(n*m) où n = nombre de producteurs",
    "m = nombre moyen de récoltes par producteur"
])

add_content_slide("Algorithme : Meilleur Producteur", [
    "1. SI liste vide : Afficher erreur et quitter",
    "",
    "2. Initialiser max_production = -1, nom_meilleur = \"\"",
    "",
    "3. POUR CHAQUE producteur DANS liste",
    "     Calculer total_producteur = somme des quantités",
    "     SI total_producteur > max_production",
    "       Mettre à jour max_production et nom_meilleur",
    "",
    "4. Afficher \"Meilleur producteur : \" + nom_meilleur",
    "5. Afficher \"Production : \" + max_production + \" kg\"",
    "",
    "Complexité : O(n*m) - même que production totale"
])

add_content_slide("Flux d'Exécution Principal", [
    "1. Démarrage : python main.py",
    "",
    "2. Importation de tous les modules",
    "",
    "3. Affichage du menu principal",
    "",
    "4. BOUCLE INFINIE",
    "     a. Afficher les options disponibles",
    "     b. Demander choix à l'utilisateur (1-9 ou 0)",
    "     c. Selon le choix : Exécuter fonction correspondante",
    "     d. Retourner à 4.a",
    "",
    "5. Fin du programme si option 0 est sélectionnée"
])

# SECTION IV : DÉMONSTRATION
add_title_slide("IV. DÉMONSTRATION", "Illustration du Fonctionnement")

add_content_slide("Affichage du Menu Principal", [
    "==========================================",
    " GESTION D'UNE COOPÉRATIVE AGRICOLE",
    "==========================================",
    "1. Ajouter un producteur",
    "2. Afficher les producteurs",
    "3. Rechercher un producteur",
    "4. Enregistrer une récolte",
    "5. Afficher les récoltes",
    "6. Production totale",
    "7. Production par produit",
    "8. Meilleur producteur",
    "9. Production moyenne par producteur",
    "0. Quitter",
    "",
    "Votre choix :"
])

add_content_slide("Exemple d'Utilisation Complet", [
    "Scénario : Gestionnaire enregistre récoltes de 3 producteurs",
    "",
    "Étape 1 : Ajouter 3 producteurs",
    "           Amadou, Fatou, et Moussa",
    "",
    "Étape 2 : Enregistrer plusieurs récoltes par producteur",
    "          Combinaisons de produits différents",
    "",
    "Étape 3 : Afficher liste des producteurs avec récoltes",
    "",
    "Étape 4 : Afficher toutes les récoltes",
    "",
    "Étape 5 : Voir production totale"
])

add_content_slide("Exemple Continued", [
    "Étape 6 : Voir production par produit",
    "",
    "Étape 7 : Identifier meilleur producteur",
    "          Exemple : Moussa = 420 kg (Meilleur)",
    "",
    "Étape 8 : Calculer moyenne de production",
    "          Moyenne = (350 + 280 + 420) / 3",
    "          Moyenne = 350 kg par producteur",
    "",
    "Démonstration complète des capacités du système",
    "incluant validation des données et gestion d'erreurs"
])

add_content_slide("Fonctionnalités Clés Démontrées", [
    "1. Validation des données",
    "   - Noms en double rejetés",
    "   - Produits invalides refusés",
    "   - Quantités négatives ou zéro rejetées",
    "",
    "2. Gestion des erreurs",
    "   - Conversion de quantité",
    "   - Vérification de positivité",
    "   - Messages clairs",
    "",
    "3. Recherche efficace",
    "   - Localisation rapide de producteurs",
    "",
    "4. Rapports détaillés",
    "   - Production par producteur et produit",
    "",
    "5. Interface conviviale"
])

# SECTION V : CONCLUSION
add_title_slide("V. CONCLUSION", "Difficultés et Perspectives d'Amélioration")

add_content_slide("Résumé des Réalisations", [
    "Le projet AGRI-SEN a atteint tous ses objectifs :",
    "",
    "Atteint : 5 modules indépendants mais intégrés",
    "Atteint : Interface utilisateur intuitive",
    "Atteint : Système de validation complète",
    "Atteint : Rapports et analyses détaillées",
    "Atteint : Code bien commenté et maintainable",
    "",
    "Le système est complet, fonctionnel et prêt pour",
    "utilisation réelle dans les coopératives agricoles."
])

add_content_slide("Difficultés Rencontrées", [
    "1. Coordination d'équipe",
    "   - Synchronisation entre 5 développeurs",
    "   - Communication et organisation claires nécessaires",
    "",
    "2. Imports circulaires",
    "   - Clarification des dépendances requise",
    "",
    "3. Validation des données",
    "   - Tests multiples pour éviter les failles",
    "",
    "4. Structure des données",
    "   - Plusieurs itérations pour optimiser",
    "   - Choix entre dictionnaires et listes"
])

add_content_slide("Perspectives d'Amélioration - Part 1", [
    "1. Persistance des données",
    "   - Ajouter base de données (SQLite, MySQL)",
    "",
    "2. Gestion des utilisateurs",
    "   - Authentification multi-rôles",
    "",
    "3. Édition et suppression",
    "   - Modification de producteurs et récoltes",
    "   - Historique des changements",
    "",
    "4. Export des rapports",
    "   - Génération en PDF ou Excel",
    "   - Créer graphiques pour visualiser tendances"
])

add_content_slide("Perspectives d'Amélioration - Part 2", [
    "5. Application mobile",
    "   - Enregistrement direct des récoltes sur le terrain",
    "",
    "6. Système de notification",
    "   - Alertes pour récoltes attendues",
    "",
    "7. Intégration météo",
    "   - Analyse impact climat sur production",
    "",
    "8. Interface graphique",
    "   - Remplacer menu texte par GUI",
    "   - Utiliser tkinter ou PyQt"
])

add_content_slide("Évaluation Finale", [
    "AGRI-SEN est un projet réussi qui démontre :",
    "",
    "- Bonne compréhension des principes logiciels",
    "- Modularité et séparation des responsabilités",
    "- Validation rigoureuse des données",
    "- Code bien structuré et extensible",
    "",
    "La collaboration d'équipe a été efficace.",
    "",
    "Défis surmontés grâce à communication claire",
    "et approche méthodique.",
    "",
    "Résultat : Système complet prêt pour déploiement"
])

# Sauvegarder la présentation
output_path = "AGRI-SEN_Rapport_Complet.pptx"
prs.save(output_path)
print(f"Rapport PowerPoint créé avec succès : {output_path}")
