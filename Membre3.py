from membre1 import producteurs
from membre2 import PRODUITS_VALIDES 

def production_totale():
    """Calcule et affiche la production totale."""
    total = 0.0
    
    for p in producteurs:
        for r in p["recoltes"]:
            total += r["quantite"]
            
    print(f"La production totale est : {total} kg")
    return total

def production_par_produit():
    """Calcule et affiche la production par produit."""
    stats_produits = {}
    
    for p in producteurs:
        for r in p["recoltes"]:
            nom_produit = r["produit"]
            if nom_produit in stats_produits:
                stats_produits[nom_produit] += r["quantite"]
            else :
                stats_produits[nom_produit] = r["quantite"]

    print("\n--- Production par Produit ---")
    for produit, quantite in stats_produits.items():
        print(f"{produit} : {quantite} kg")
        
    return stats_produits