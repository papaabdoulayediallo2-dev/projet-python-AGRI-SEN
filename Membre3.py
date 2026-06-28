production_ferme = {
    "Mil": 1500,
    "Mais": 800,
    "Riz": 1200,
    "Arachide": 600
}
#.Fonction pour calculer la production totale
def production_totale(donnees_production): 
    """
    calculer la somme de toutes les productions.
    """
    total = sum(donnees_production.values())
    return total
#.Fonction pour afficher la production par produit
def production_par_produit(donnees_production):
    """
    Parcourt le dictionnaire pour afficher le detail de chaque produit.
    """
    print("\n--- Détail de la production par produit ---")
    for produit, quantite in donnees_production.items():
        print(f"- {produit} : {quantite} kg")

#=================================================
#EXECUTION DU PROGRAMME
#=================================================

#.Appel de la fonction bonus pour afficher le détail
production_par_produit(production_ferme)

#.Appel de la fonction pour le total et stockage du résultat
total_kg = production_totale(production_ferme)

print("\n--- Résultat Globale ---")  
print(f"Production totale de la ferme es de : {total_kg} kg")