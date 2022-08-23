Attribute VB_Name = "Reveal_cartes"
Option Explicit
Option Base 1

Sub reveal_cartes_communes(un_jeu As jeu, _
                            ensemble_cartes As Scripting.Dictionary, _
                            ws_table As Worksheet, _
                            ws_lancement As Worksheet, _
                            premiere As Integer, _
                            derniere As Integer)
                            
'Procedure affichant les cartes communes tirees.

    Dim j As Integer
    Dim nb_joueurs As Integer
    nb_joueurs = un_jeu.get_joueurs_en_jeu.count
    
    For j = premiere To derniere
        With un_jeu.get_paquet.item(nb_joueurs * 2 + j)
            'Affichage du tirage
            ws_table.Range("valeur_tirage_" & j).Value = .get_valeur_string
            ws_table.Range("couleur_tirage_" & j).Value = ws_lancement.Range(.get_couleur_string)
            ws_table.Range("couleur_tirage_" & j).Font.Color = ws_lancement.Range(.get_couleur_string).Font.Color
            
            'Stockage du tirage dans la variable ensemble_cartes
            Dim une_carte As carte
            Set une_carte = New carte
            Let une_carte.let_couleur = .get_couleur
            Let une_carte.let_valeur = .get_valeur
            ensemble_cartes.Add j, une_carte
        End With
    Next j
End Sub

Sub reveal_all(un_jeu As jeu, ws_table As Worksheet, ws_lancement As Worksheet)

'Procedure revelant les cartes des joueurs en fin de manche pour voir le gagnant.

    Dim i As Integer
    Dim num As Integer
    num = 1
    
    Dim nb_joueurs As Integer
    nb_joueurs = un_jeu.get_joueurs_en_jeu.count
    
    For i = 1 To nb_joueurs
        
        If un_jeu.get_joueurs_en_jeu(i).get_action <> "passe" Then
            num = un_jeu.get_joueurs_en_jeu.item(i).get_number
            
            With un_jeu.get_joueurs_en_jeu.item(i).get_une_main.get_carte_1
                ws_table.Range("valeur_carte_1_J" & num).Value = .get_valeur_string
                ws_table.Range("couleur_carte_1_J" & num).Value = ws_lancement.Range(.get_couleur_string).Value
                ws_table.Range("couleur_carte_1_J" & num).Font.Color = ws_lancement.Range(.get_couleur_string).Font.Color
            End With
            With un_jeu.get_joueurs_en_jeu.item(i).get_une_main.get_carte_2
                ws_table.Range("valeur_carte_2_J" & num).Value = .get_valeur_string
                ws_table.Range("couleur_carte_2_J" & num).Value = ws_lancement.Range(.get_couleur_string).Value
                ws_table.Range("couleur_carte_2_J" & num).Font.Color = ws_lancement.Range(.get_couleur_string).Font.Color
            End With
        End If
        
    Next i

End Sub

