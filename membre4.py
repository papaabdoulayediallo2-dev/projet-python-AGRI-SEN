producteurs_data = {
    "Moussa": 2500,
    "Fatou": 1800,
    "Amadou": 2200,
    "Awa": 1500
}

def meilleur_producteur(donnees_cooperative):
    if not donnees_cooperative:
        print("Aucun producteur enregistré.")
        return

    nom_meilleur = ""
    max_production = -1
    
    for producteur, production in donnees_cooperative.items():
        if production > max_production:
            max_production = production
            nom_meilleur = producteur
            
    print(f"Meilleur producteur : {nom_meilleur}")
    print(f"Production : {max_production} kg")


def moyenne_producteur(donnees_cooperative):
    if not donnees_cooperative:
        print("Aucun producteur pour calculer la moyenne.")
        return

    total_general = 0
    nombre_producteurs = 0
    
    for production in donnees_cooperative.values():
        total_general += production
        nombre_producteurs += 1
        
    moyenne = total_general / nombre_producteurs
    print(f"Production moyenne par producteur : {moyenne:.2f} kg")


meilleur_producteur(producteurs_data)
print("-" * 30)
moyenne_producteur(producteurs_data)