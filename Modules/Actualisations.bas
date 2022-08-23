Attribute VB_Name = "Actualisations"
Option Explicit
Option Base 1

Sub actualisation_jeu(table As Worksheet, param As Worksheet, jeu_ As jeu, resultats As Collection)

'Fonction que l'on appelle a la fin de chaque manche pour actualiser les valeurs de _
l'instance de classe qui contient tous les parametres du jeu.

Dim gains As Long
Dim joueur_ As joueur
Dim i As Integer
Dim collection_joueurs As New Collection
Set collection_joueurs = jeu_.get_joueurs_en_jeu

'Calcul du gains par joueur en fin de manche
gains = table.Range("pot").Value / resultats.count

For Each joueur_ In collection_joueurs
    For i = 1 To resultats.count
        'Condition pour filtrer les gagnants
        If resultats.item(i)(3) = joueur_.get_number Then
            'Calcul du stack
            Let joueur_.let_stack = joueur_.get_stack + gains - (param.Range("mise_max").Value - joueur_.get_mise) / resultats.count
            'Actualisation du pot
            jeu_.pot = jeu_.pot - gains + (param.Range("mise_max").Value - joueur_.get_mise) / resultats.count
            'Reinitialisation de la mise
            Let joueur_.let_mise = 0
        End If
    Next i
    'Reinitialisation de la dernière action
    Let joueur_.let_action = ""
Next joueur_

'Si le pot n'est pas nul car un joueur a force le tapis a un autre
For Each joueur_ In collection_joueurs
    'Condition pour filtrerr les joueurs ayant mise plus que le gagnant de la manche
    If joueur_.get_mise = param.Range("mise_max") And jeu_.pot > 0 Then
        'Calcul du stack
        Let joueur_.let_stack = joueur_.get_stack + jeu_.pot / (nb_joueurs_en_manche(jeu_) - 1)
        'Actualisation du pot
        jeu_.pot = jeu_.pot - joueur_.get_mise + gains
    End If
    'Reinitialisation de la mise
    Let joueur_.let_mise = 0
Next

'Gestion des joueurs arrivant a un stack nul, qui ont donc perdu.
For Each joueur_ In collection_joueurs
    If joueur_.get_statut = 0 Then
        MsgBox "Le joueur " & joueur_.get_number & " n'a plus d'argent et a donc perdu.", vbOKOnly + vbExclamation
        'On noircit les cases des joueurs ayant perdu
        table.Cells(4 * joueur_.get_number - 1, 2).Interior.Color = vbBlack
        table.Cells(4 * joueur_.get_number - 1, 6).Resize(2, 2).Interior.Color = vbBlack
    End If
Next

End Sub

Sub actualisation_affichage(table As Worksheet, jeu_ As jeu)

'Fonction que l'on appelle a la fin de chaque manche pour actualiser les valeurs _
concernant les joueurs sur la table de jeu.

Dim i As Integer

'Nettoyage des cases qui affichent les cartes, mises, stacks et actions des joueurs
For i = 1 To jeu_.get_joueurs.count
    table.Range("Action_J" & i).ClearContents
    table.Range("Mise_J" & i).ClearContents
    table.Range("valeur_carte_1_J" & i).ClearContents
    table.Range("couleur_carte_1_J" & i).ClearContents
    table.Range("valeur_carte_2_J" & i).ClearContents
    table.Range("couleur_carte_2_J" & i).ClearContents
    table.Range("Stack_J" & i).Value = jeu_.get_joueurs.item(i).get_stack
Next i

'Nettoyage des cases affichant le tirage des cartes communes
For i = 1 To 5
    table.Range("valeur_tirage_" & i).ClearContents
    table.Range("couleur_tirage_" & i).ClearContents
Next i

'Nettoyage de la case affichant le pot
table.Range("pot").Value = 0

End Sub

Sub affichage_pos_blind(un_jeu As jeu, positions As Collection, ws_table As Worksheet, ws_param As Worksheet)
    
'Fonction que l'on appelle a la fin de chaque manche pour actualiser _
les valeurs de position des joueurs et des blinds obligatoires.

    Dim k As Integer
    Dim i As Integer
    Dim nb_joueurs As Integer
    nb_joueurs = un_jeu.get_joueurs_en_jeu.count
    
    k = 0
    For i = 1 To nb_joueurs
        k = un_jeu.get_joueurs_en_jeu.item(i).get_number
        ws_table.Cells(4 * k - 1, 2).Value = positions(i)
        Let un_jeu.get_joueurs_en_jeu.item(i).let_position = positions(i)
        If (ws_table.Cells(4 * k - 1, 2).Value = "Button / Small Blind" Or ws_table.Cells(4 * k - 1, 2).Value = "Small Blind") Then
            ws_table.Cells(4 * k, 7).Value = CLng(ws_param.Range("blind").Value)
            ws_table.Cells(4 * k - 2, 7).Value = ws_table.Cells(4 * k - 2, 7).Value - CLng(ws_param.Range("blind").Value)
        ElseIf ws_table.Cells(4 * k - 1, 2).Value = "Big Blind" Then
            ws_table.Cells(4 * k, 7).Value = 2 * CLng(ws_param.Range("blind").Value)
            ws_table.Cells(4 * k - 2, 7).Value = ws_table.Cells(4 * k - 2, 7).Value - 2 * CLng(ws_param.Range("blind").Value)
        End If
    Next i
            

End Sub

Function reinit_positions(nb_joueurs As Integer, UTG_pos As Integer) As Collection

'Fonction qui actualise les position de parole des joueurs au changement de manche.

    'Creation d'une collection contenant les position dans l'ordre initial
    Set reinit_positions = New Collection
    Dim i, j As Integer
    'Definition des positions
    If nb_joueurs = 2 Then
        reinit_positions.Add "Button / Small Blind"
        reinit_positions.Add "Big Blind"
    Else
        reinit_positions.Add "Button"
        reinit_positions.Add "Small Blind"
        reinit_positions.Add "Big Blind"
        reinit_positions.Add "UTG"
        reinit_positions.Add "UTG+1"
        reinit_positions.Add "Cut-Off"

        For i = nb_joueurs + 1 To 6
            reinit_positions.Remove (nb_joueurs + 1)
        Next i
    End If

    'Reorganisation de la collection pour mettre la position UTG en premier
        'Distinction de cas selon le nombre de joueurs actifs pendant la prochaine manche
    If nb_joueurs = 2 Then
        If UTG_pos = 2 Then
            reinit_positions.Add reinit_positions.item(nb_joueurs), Before:=1
            reinit_positions.Remove (nb_joueurs + 1)
        End If
    ElseIf nb_joueurs = 3 Then
        For j = 3 To 3 - UTG_pos + 2 Step -1
            reinit_positions.Add reinit_positions.item(nb_joueurs), Before:=1
            reinit_positions.Remove (nb_joueurs + 1)
        Next j
    ElseIf nb_joueurs = 4 Then
        For j = nb_joueurs To nb_joueurs - UTG_pos + 1 Step -1
            reinit_positions.Add reinit_positions.item(nb_joueurs), Before:=1
            reinit_positions.Remove (nb_joueurs + 1)
        Next j
    ElseIf nb_joueurs = 5 Then
        For j = nb_joueurs To nb_joueurs - UTG_pos Step -1
            reinit_positions.Add reinit_positions.item(nb_joueurs), Before:=1
            reinit_positions.Remove (nb_joueurs + 1)
        Next j
    ElseIf nb_joueurs = 6 Then
        For j = nb_joueurs To nb_joueurs - UTG_pos - 1 Step -1
            reinit_positions.Add reinit_positions.item(nb_joueurs), Before:=1
            reinit_positions.Remove (nb_joueurs + 1)
        Next j
    End If
   
End Function

Sub actualisation_pot(un_jeu As jeu, ws_table As Worksheet)

'Focntion qui actualise la valeur du pot après chaque tour de mis

    Dim joueur_ As joueur
    
    ws_table.Range("pot").Value = 0
    un_jeu.pot = 0

    For Each joueur_ In un_jeu.get_joueurs_en_jeu
        ws_table.Range("pot").Value = ws_table.Range("pot").Value + joueur_.get_mise
        un_jeu.pot = un_jeu.pot + joueur_.get_mise
    Next

End Sub
