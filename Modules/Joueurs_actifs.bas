Attribute VB_Name = "Joueurs_actifs"
Option Explicit
Option Base 1

Sub mise_a_jour_joueurs(un_jeu_ As jeu)

'Procedure actualisant les joueurs avec un statut actif, _
soit les joueurs qui n'ont pas ete elimine pour cause de stack nul.

    Dim i As Integer
    Dim item As Variant

    Dim joueurs_en_jeu As Collection
    Set joueurs_en_jeu = New Collection

    For Each item In un_jeu_.get_joueurs
        joueurs_en_jeu.Add item
    Next
    
    For i = joueurs_en_jeu.count To 1 Step -1
        If joueurs_en_jeu(i).get_statut = 0 Then
            joueurs_en_jeu.Remove (i)
        End If
    Next i
    
    Let un_jeu_.let_joueurs_en_jeu = joueurs_en_jeu
    
End Sub

Function nb_joueurs_en_manche(un_jeu As jeu) As Integer

'Fonction renvoyant le nombre de joueurs encore en lice dans la manche, _
soit les joueurs qui n'ont pas encore passe.

Dim joueur As joueur
Dim nb As Integer

For Each joueur In un_jeu.get_joueurs_en_jeu
    If joueur.get_action <> "passe" Then
        nb = nb + 1
    End If
Next joueur
nb_joueurs_en_manche = nb

End Function
