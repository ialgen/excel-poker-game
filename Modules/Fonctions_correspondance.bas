Attribute VB_Name = "Fonctions_correspondance"
Option Explicit
Option Base 1

Function f(i_boucle As Integer, pos_utg As Integer, nb_j As Integer) As Integer

'Fonction bijective qui associe un indice d'iteration a la position de parole du joueur _
en prenant pour point de repere la position de du joueur UTG.

    If i_boucle <= nb_j - pos_utg + 1 Then
        f = i_boucle + pos_utg - 1
    Else
        f = i_boucle - (nb_j - pos_utg + 1)
    End If
End Function

Function hauteur_string(hauteur_ As Integer) As String

'Fonction convertissant la valeur numerique d'une carte en une valeur STRING cmprehensible par l'utilisateur.

Select Case hauteur_
    Case 13
        hauteur_string = "A"
    Case 12
        hauteur_string = "K"
    Case 11
        hauteur_string = "Q"
    Case 10
        hauteur_string = "J"
    Case 9
        hauteur_string = 10
    Case 8
        hauteur_string = 9
    Case 7
        hauteur_string = 8
    Case 6
        hauteur_string = 7
    Case 5
        hauteur_string = 6
    Case 4
        hauteur_string = 5
    Case 3
        hauteur_string = 4
    Case 2
        hauteur_string = 3
    Case 1
        hauteur_string = 2
    Case Else
End Select

End Function
