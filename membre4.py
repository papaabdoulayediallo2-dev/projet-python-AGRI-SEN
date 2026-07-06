from membre1 import producteurs


def meilleur_producteur(liste_producteurs):
    if not liste_producteurs:
        print("Aucun producteur enregistré.")
        return

    nom_meilleur = ""
    max_production = -1.0
    
    for p in liste_producteurs:
        total_producteur = 0.0
        for r in p["recoltes"]:
            total_producteur += r["quantite"]
            
        if total_producteur > max_production:
            max_production = total_producteur
            nom_meilleur = p["nom"]
            
    print(f"Meilleur producteur : {nom_meilleur}")
    print(f"Production : {max_production} kg")


def moyenne_producteur(liste_producteurs):
    if not liste_producteurs:
        print("Aucun producteur pour calculer la moyenne.")
        return

    total_general = 0.0
    nombre_producteurs = len(liste_producteurs)
    
    for p in liste_producteurs:
        for r in p["recoltes"]:
            total_general += r["quantite"]
            
    moyenne = total_general / nombre_producteurs
    print(f"Production moyenne par producteur : {moyenne:.2f} kg")