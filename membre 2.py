# Liste des produits autorisés dans la coopérative
PRODUITS_VALIDES = ["Mil", "Maïs", "Riz", "Arachide"]


def produit_valide(produit):
    """Vérifie si le produit fait partie des produits autorisés."""
    return produit.capitalize() in PRODUITS_VALIDES


def enregistrer_recolte():
    """Demande les infos d'une récolte et l'associe à un producteur existant."""
    nom = input("Nom du producteur : ")

    producteur = rechercher_producteur(nom)
    if producteur is None:
        print(f"Le producteur '{nom}' n'existe pas. Ajoutez-le d'abord.")
        return

    produit = input(f"Produit récolté {PRODUITS_VALIDES} : ")

    if not produit_valide(produit):
        print(f" Produit invalide. Choisissez parmi : {PRODUITS_VALIDES}")
        return

    quantite_str = input("Quantité récoltée (kg) : ")

    try:
        quantite = float(quantite_str)
    except ValueError:
        print(" La quantité doit être un nombre.")
        return

    if quantite <= 0:
        print(" La quantité doit être supérieure à 0.")
        return

    recolte = {
        "nom": producteur["nom"],
        "produit": produit.capitalize(),
        "quantite": quantite
    }

    producteur["recoltes"].append(recolte)
    print(f" Récolte de {quantite} kg de {produit.capitalize()} enregistrée pour {producteur['nom']}.")


def afficher_recoltes(nom=None):
    """Affiche les récoltes d'un producteur précis, ou de tous si nom=None."""
    if not producteurs:
        print("Aucun producteur enregistré pour le moment.")
        return

    if nom is not None:
        producteur = rechercher_producteur(nom)
        if producteur is None:
            print(f"Le producteur '{nom}' n'existe pas.")
            return
        liste_producteurs = [producteur]
    else:
        liste_producteurs = producteurs

    print("\n--- Récoltes enregistrées ---")
    for p in liste_producteurs:
        if not p["recoltes"]:
            print(f"{p['nom']} : aucune récolte enregistrée.")
            continue

        print(f"{p['nom']} :")
        for r in p["recoltes"]:
            print(f"   - {r['produit']} : {r['quantite']} kg")
    print()