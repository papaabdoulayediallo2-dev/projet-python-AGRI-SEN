# Liste globale qui stockera tous les producteurs
# Chaque producteur est un dictionnaire : {"nom": ..., "recoltes": []}
producteurs = []


def ajouter_producteur():
    """Demande un nom et ajoute un nouveau producteur à la liste."""
    nom = input("Nom du producteur : ")

    if nom == "":
        print(" Le nom ne peut pas être vide.")
        return

    # Vérifier si le producteur existe déjà
    if rechercher_producteur(nom) is not None:
        print(f"Le producteur '{nom}' existe déjà.")
        return

    producteurs.append({"nom": nom, "recoltes": []})
    print(f" Producteur '{nom}' ajouté avec succès.")


def afficher_producteurs():
    """Affiche la liste de tous les producteurs enregistrés."""
    if not producteurs:
        print("Aucun producteur enregistré pour le moment.")
        return

    print("\n--- Liste des producteurs ---")
    for i, p in enumerate(producteurs, start=1):
        nb_recoltes = len(p["recoltes"])
        print(f"{i}. {p['nom']} ({nb_recoltes} récolte(s) enregistrée(s))")
    print()


def rechercher_producteur(nom):
    """Recherche un producteur par son nom et renvoie son dictionnaire (ou None)."""
    for p in producteurs:
        if p["nom"].lower() == nom.lower():
            return p
    return None