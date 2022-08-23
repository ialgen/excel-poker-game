Attribute VB_Name = "Phases_action"
Option Explicit
Option Base 1

Function test_de_mise(jeu_ As jeu, mise_max_ As Long, changement_phase_ As Boolean) As Boolean

'Fonction qui teste si les conditions sont remplie pour passerr a l'etape suivante de la manche.

If changement_phase_ Then
    test_de_mise = Not changement_phase_
    GoTo fin_fonction
End If

Dim resultat As Boolean
Dim joueur_ As joueur

resultat = True

'Les joueurs avec une mise inferieure ˆ la mise max doivent s'etre couches.
For Each joueur_ In jeu_.get_joueurs_en_jeu
    If joueur_.get_mise < mise_max_ And joueur_.get_action <> "passe" Then
        resultat = False
    End If
Next

'Dans le cas on un joueur est force a faire tapis il peut avoir une mise inferieure a la mise max _
sans pour autant s'etre couche au cours de la manche
For Each joueur_ In jeu_.get_joueurs_en_jeu
    If joueur_.get_mise < mise_max_ Then
        If joueur_.get_action = "passe" Then
            resultat = True
        ElseIf joueur_.get_stack = 0 Then
            resultat = True
        Else
            resultat = False
        End If
    End If
Next

test_de_mise = resultat

fin_fonction:
End Function

Sub phase_actions(un_jeu As jeu, ws_table As Worksheet, ws_param As Worksheet, ws_lancement As Worksheet)

'Procedure qui gere les actions d'un joueur au cours d'une manche.
    
    Dim j As Integer
    Dim nb_joueurs, indice_utg, num_joueur_actif, num_col_joueur_actif As Integer

    indice_utg = CInt(ws_param.Range("indice_utg").Value)
    nb_joueurs = un_jeu.get_joueurs_en_jeu.count
    
    For j = 1 To nb_joueurs
        
        'indice du joueur qui va parler
        num_joueur_actif = f(CInt(j), CInt(indice_utg), CInt(nb_joueurs))
        
        'numŽro(dans le nom) du joueur qui va parler
        num_col_joueur_actif = un_jeu.get_joueurs_en_jeu(num_joueur_actif).get_number
        
        'pour parler il ne faut pas etre couchŽ
        If un_jeu.get_joueurs_en_jeu(num_joueur_actif).get_action <> "passe" Then
            
            ws_param.Range("joueur_actif").Value = un_jeu.get_joueurs_en_jeu(num_joueur_actif).get_number
            
            'message pour communiquer avec les joueurs
            While MsgBox("Vous etes bien le joueur " & num_col_joueur_actif & " ?", vbYesNo) = vbNo
                MsgBox "Faites passer l'ordinateur au joueur " & num_col_joueur_actif & " !"
            Wend
            
            'Affichage de sa main
            With un_jeu.get_joueurs_en_jeu.item(num_joueur_actif).get_une_main.get_carte_1
                ws_table.Range("valeur_carte_1_J" & num_col_joueur_actif).Value = .get_valeur_string
                ws_table.Range("couleur_carte_1_J" & num_col_joueur_actif).Value = ws_lancement.Range(.get_couleur_string).Value
                ws_table.Range("couleur_carte_1_J" & num_col_joueur_actif).Font.Color = ws_lancement.Range(.get_couleur_string).Font.Color
            End With
            With un_jeu.get_joueurs_en_jeu.item(num_joueur_actif).get_une_main.get_carte_2
                ws_table.Range("valeur_carte_2_J" & num_col_joueur_actif).Value = .get_valeur_string
                ws_table.Range("couleur_carte_2_J" & num_col_joueur_actif).Value = ws_lancement.Range(.get_couleur_string).Value
                ws_table.Range("couleur_carte_2_J" & num_col_joueur_actif).Font.Color = ws_lancement.Range(.get_couleur_string).Font.Color
            End With
            
            'Actions du joueur via userform
            Actions.Show
            
            'feuille excel a jour
            Let un_jeu.get_joueurs_en_jeu.item(num_joueur_actif).let_action = ws_table.Range("Action_J" & num_col_joueur_actif)
            un_jeu.get_joueurs_en_jeu.item(num_joueur_actif).miser (ws_table.Range("Mise_J" & num_col_joueur_actif))
            
            'Desaffichage de la main du joueur pour passer au suivant
            ws_table.Range("valeur_carte_1_J" & num_col_joueur_actif).ClearContents
            ws_table.Range("couleur_carte_1_J" & num_col_joueur_actif).ClearContents
            ws_table.Range("valeur_carte_2_J" & num_col_joueur_actif).ClearContents
            ws_table.Range("couleur_carte_2_J" & num_col_joueur_actif).ClearContents
            
        End If
        
    Next j
End Sub

Function resultats(un_jeu As jeu, ensemble_cartes As Scripting.Dictionary) As Collection

'Fonction qui renvoie une collection contenant la ou les mains gagnante ou gagnantes _
et les joueurs a qui elles appartiennent.
'Elle permet de conclure la manche en determinant un gagnant.
    
    Dim joueur_ As joueur
    Dim combi_ As Combinaisons
    Set combi_ = New Combinaisons
    Dim resultats_manche As New Collection
    Set resultats_manche = New Collection
    Dim rang_main As Variant
    Dim meme_hauteur As Boolean
    Dim i As Integer
    
    For Each joueur_ In un_jeu.get_joueurs_en_jeu
        If joueur_.get_action <> "passe" Then
            With joueur_.get_une_main
                Set ensemble_cartes(6) = .get_carte_1
                Set ensemble_cartes(7) = .get_carte_2
                rang_main = combi_.best_main(ensemble_cartes)
                rang_main(3) = joueur_.get_number
            End With
            If resultats_manche.count > 0 Then
                If resultats_manche.item(1)(1) > rang_main(1) Then 'prpp a la hierarchie le numero est faible plus on a une bonne combinaison
                    Set resultats_manche = New Collection
                    resultats_manche.Add rang_main
                ElseIf resultats_manche.item(1)(1) = rang_main(1) Then
                    meme_hauteur = True
                    For i = 1 To UBound(rang_main(2))
                        If resultats_manche.item(1)(2)(i) = rang_main(2)(i) Then
                            meme_hauteur = True
                        ElseIf resultats_manche.item(1)(2)(i) > rang_main(2)(i) Then
                            meme_hauteur = False
                            GoTo next_joueur
                        Else
                            meme_hauteur = False
                            Set resultats_manche = New Collection
                            resultats_manche.Add rang_main
                            GoTo next_joueur
                        End If
                    Next
                    If meme_hauteur = True Then
                        resultats_manche.Add rang_main
                    End If
                End If
            Else
                resultats_manche.Add rang_main
            End If
        End If
next_joueur:
    Next
    Set resultats = resultats_manche
End Function
