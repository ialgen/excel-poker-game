Attribute VB_Name = "Initialisation_partie"
Option Explicit
Option Base 1

Sub initialisation_class_jeu(un_jeu_ As jeu, ws_table As Worksheet, ws_param As Worksheet)

'Fonction initialisant les valeurs des proprietes de l'instance de classe jeu au debut d'une nouvelle partie.

    Dim iter As Integer
    Dim joueur_ As joueur
    Let un_jeu_.let_joueurs(ws_param.Range("Nbre_joueurs").Value) = ws_param.Range("argent_en_jeu").Value
    iter = 1
    For Each joueur_ In un_jeu_.get_joueurs
        Let joueur_.let_name = ws_table.Range("Nom_J" & iter)
        Let joueur_.let_number = iter
        Let joueur_.let_stack = ws_table.Range("Stack_J" & iter)
        Let joueur_.let_mise = ws_table.Range("Mise_J" & iter)
        Let joueur_.let_position = ws_table.Range("Position_J" & iter)
        iter = iter + 1
    Next
    
    Let un_jeu_.let_joueurs_en_jeu = un_jeu_.get_joueurs

    un_jeu_.blind = ws_param.Range("blind").Value
    
End Sub

Sub distribution_cartes(un_jeu As jeu)

'Fonction distribuant les cartes aux joueurs en debut de partie en les tirant du paquet.

    Dim joueur_ As joueur
    Dim iter As Integer
    Dim une_main As hand
    Dim une_carte As carte
    
    iter = 1
    For Each joueur_ In un_jeu.get_joueurs_en_jeu
        Set une_main = New hand
    
        'Distribution carte 1
        Set une_carte = New carte
        Let une_carte.let_couleur = un_jeu.get_paquet.item(-1 + 2 * iter).get_couleur
        Let une_carte.let_valeur = un_jeu.get_paquet.item(-1 + 2 * iter).get_valeur
        Let une_main.let_carte_1 = une_carte
    
        'Distribution carte 2
        Set une_carte = New carte
        Let une_carte.let_couleur = un_jeu.get_paquet.item(2 * iter).get_couleur
        Let une_carte.let_valeur = un_jeu.get_paquet.item(2 * iter).get_valeur
        Let une_main.let_carte_2 = une_carte
        
        Let joueur_.let_une_main = une_main
        iter = iter + 1
    Next
End Sub
