# ==========================================
# MEMBRE 5 : Menu principal et intégration
# ==========================================

# Importation des fonctions créées par les autres membres

from membre1 import (
    ajouter_producteur,
    afficher_producteurs,
    rechercher_producteur,
    producteurs
)

from membre2 import (
    enregistrer_recolte,
    afficher_recoltes
)

from Membre3 import (
    production_totale,
    production_par_produit
)

from membre4 import (
    meilleur_producteur,
    moyenne_producteur
)


# ======================================================
# Fonction qui affiche le menu principal du programme
# ======================================================
def menu():

    # Boucle infinie permettant au menu de rester affiché
    # tant que l'utilisateur ne choisit pas de quitter.
    while True:

        print("\n==========================================")
        print(" GESTION D'UNE COOPÉRATIVE AGRICOLE")
        print("==========================================")
        print("1. Ajouter un producteur")
        print("2. Afficher les producteurs")
        print("3. Rechercher un producteur")
        print("4. Enregistrer une récolte")
        print("5. Afficher les récoltes")
        print("6. Production totale")
        print("7. Production par produit")
        print("8. Meilleur producteur")
        print("9. Production moyenne par producteur")
        print("0. Quitter")

        # Lecture du choix de l'utilisateur
        choix = input("\nVotre choix : ")

        # ===============================
        # Gestion des différents choix
        # ===============================

        if choix == "1":
            # Appel de la fonction d'ajout d'un producteur
            ajouter_producteur()

        elif choix == "2":
            # Affichage de tous les producteurs
            afficher_producteurs()

        elif choix == "3":
            # Demande le nom du producteur à rechercher
            nom = input("Nom du producteur : ")

            # Recherche du producteur
            resultat = rechercher_producteur(nom)

            # Vérifie si le producteur existe
            if resultat:
                print(f"\nProducteur trouvé : {resultat['nom']}")
                print(f"Nombre de récoltes : {len(resultat['recoltes'])}")
            else:
                print("Producteur introuvable.")

        elif choix == "4":
            # Enregistrement d'une nouvelle récolte
            enregistrer_recolte()

        elif choix == "5":
            # L'utilisateur peut afficher toutes les récoltes
            # ou celles d'un seul producteur.
            nom = input(
                "Nom du producteur (laisser vide pour tout afficher) : "
            )

            if nom == "":
                afficher_recoltes()
            else:
                afficher_recoltes(nom)

        elif choix == "6":
            # Affichage de la production totale
            production_totale()

        elif choix == "7":
            # Affichage de la production pour chaque produit
            production_par_produit()

        elif choix == "8":
            # Affichage du meilleur producteur
            meilleur_producteur(producteurs)

        elif choix == "9":
            # Calcul de la production moyenne
            moyenne_producteur(producteurs)

        elif choix == "0":
            # Arrêt du programme
            print("\nMerci d'avoir utilisé notre application.")
            break

        else:
            # Message si le choix est invalide
            print("Choix invalide. Veuillez réessayer.")


# =======================================
# Point d'entrée du programme
# =======================================
# Lance le menu principal lorsque le fichier est exécuté.
menu()